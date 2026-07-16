$workdir = $psscriptroot
function Generate-Manifests {
param(
    [hashtable]$Headers = $GraphHeaders,

    [ValidateSet('Graph', 'SharePointV2')]
    [string]$ApiMode = 'Graph',

    # Required for SharePointV2 mode. Example: contoso
    [string]$TenantName,

    [ValidateSet('All', 'Sites', 'Drives', 'Lists')]
    [string[]]$ManifestType = @('Lists'),

    [string]$OutputPath,

    # Document libraries are already enumerated as drives. Enable this only
    # when you also want their listItem/fields representations.
    [switch]$IncludeDocumentLibraryListItems,

    # Use this when you only want list definitions, columns, and content
    # types. List item field payloads are skipped.
    [switch]$ListMetadataOnly
)

function Get-HttpStatusCode {
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    try {
        return [int]$ErrorRecord.Exception.Response.StatusCode
    }
    catch {
        return $null
    }
}

function New-ManifestError {
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    [ordered]@{
        uri        = $Uri
        statusCode = Get-HttpStatusCode -ErrorRecord $ErrorRecord
        message    = $ErrorRecord.Exception.Message
    }
}

function Format-ManifestDuration {
    param(
        [Parameter(Mandatory)]
        [timespan]$Duration
    )

    if ($Duration.TotalHours -ge 1) {
        return '{0:0}h {1:00}m {2:00}s' -f `
            [math]::Floor($Duration.TotalHours),
            $Duration.Minutes,
            $Duration.Seconds
    }

    if ($Duration.TotalMinutes -ge 1) {
        return '{0:0}m {1:00}s' -f `
            [math]::Floor($Duration.TotalMinutes),
            $Duration.Seconds
    }

    return '{0:0}s' -f [math]::Max(0, [math]::Round($Duration.TotalSeconds))
}

function Get-ManifestEtaText {
    param(
        [Parameter(Mandatory)]
        [System.Diagnostics.Stopwatch]$Stopwatch,

        [Parameter(Mandatory)]
        [int]$Completed,

        [Parameter(Mandatory)]
        [int]$Total
    )

    if ($Completed -le 0 -or $Total -le 0 -or $Completed -ge $Total) {
        return $null
    }

    $averageSeconds = $Stopwatch.Elapsed.TotalSeconds / $Completed
    $remaining      = $Total - $Completed
    $eta            = [timespan]::FromSeconds($averageSeconds * $remaining)

    return 'ETA {0}' -f (Format-ManifestDuration -Duration $eta)
}

function Get-ManifestSiteLabel {
    param(
        [Parameter(Mandatory)]
        [object]$Site
    )

    foreach ($propertyName in @('displayName', 'name', 'webUrl', 'id')) {
        if (-not [string]::IsNullOrWhiteSpace($Site.$propertyName)) {
            return $Site.$propertyName
        }
    }

    return '<unknown site>'
}

function Write-ManifestStatus {
    param(
        [Parameter(Mandatory)]
        [System.Diagnostics.Stopwatch]$Stopwatch,

        [Parameter(Mandatory)]
        [string]$Message,

        [ConsoleColor]$ForegroundColor = [ConsoleColor]::Cyan
    )

    Write-Host (
        '[{0}] {1}' -f
        (Format-ManifestDuration -Duration $Stopwatch.Elapsed),
        $Message
    ) -ForegroundColor $ForegroundColor
}

function Invoke-PagedMetadataRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [ValidateRange(0, 20)]
        [int]$MaxRetries = 6,

        [System.Diagnostics.Stopwatch]$StatusStopwatch,

        [string]$StatusLabel,

        [ValidateRange(1, 1000)]
        [int]$StatusPageInterval = 10
    )

    $items     = [System.Collections.Generic.List[object]]::new()
    $nextUri   = $Uri
    $deltaLink = $null
    $pageCount = 0

    while ($nextUri) {
        $attempt = 0

        while ($true) {
            try {
                $response = Invoke-RestMethod `
                    -Method Get `
                    -Uri $nextUri `
                    -Headers $Headers `
                    -ErrorAction Stop

                break
            }
            catch {
                $statusCode = Get-HttpStatusCode -ErrorRecord $_

                $isTransient = $statusCode -in @(
                    429, # Too Many Requests
                    502, # Bad Gateway
                    503, # Service Unavailable
                    504  # Gateway Timeout
                )

                if (-not $isTransient -or $attempt -ge $MaxRetries) {
                    throw
                }

                $delaySeconds = [math]::Min(
                    60,
                    [math]::Pow(2, $attempt + 1)
                )

                Write-Warning (
                    "Request returned HTTP {0}. Retrying in {1} seconds: {2}" -f
                    $statusCode,
                    $delaySeconds,
                    $nextUri
                )

                Start-Sleep -Seconds $delaySeconds
                $attempt++
            }
        }

        if ($null -ne $response.value) {
            foreach ($item in $response.value) {
                $items.Add($item)
            }
        }
        else {
            # Allows the helper to handle a non-collection response too.
            $items.Add($response)
        }

        if ($response.'@odata.deltaLink') {
            $deltaLink = $response.'@odata.deltaLink'
        }

        $pageCount++
        $nextUri = $response.'@odata.nextLink'

        if (
            $StatusStopwatch -and
            -not [string]::IsNullOrWhiteSpace($StatusLabel) -and
            (
                $pageCount -eq 1 -or
                $pageCount % $StatusPageInterval -eq 0 -or
                -not $nextUri
            )
        ) {
            $statusSuffix = if ($nextUri) { 'continuing' } else { 'done' }

            Write-ManifestStatus `
                -Stopwatch $StatusStopwatch `
                -Message (
                    '{0}: page {1}, {2} item(s) so far, {3}' -f
                    $StatusLabel,
                    $pageCount,
                    $items.Count,
                    $statusSuffix
                )
        }
    }

    [pscustomobject]@{
        Items     = $items.ToArray()
        DeltaLink = $deltaLink
    }
}

function Export-SharePointMetadataManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [ValidateSet('Graph', 'SharePointV2')]
        [string]$ApiMode = 'SharePointV2',

        # Required for SharePointV2 mode. Example: contoso
        [string]$TenantName,

        [ValidateSet('All', 'Sites', 'Drives', 'Lists')]
        [string[]]$ManifestType = @('All'),

        [string]$OutputPath = (
            Join-Path $PWD (
                'sharepoint-metadata-{0}.json' -f
                (Get-Date -Format 'yyyyMMdd-HHmmss')
            )
        ),

        # Document libraries are already enumerated as drives. Enable this
        # only when you also want their listItem/fields representations.
        [switch]$IncludeDocumentLibraryListItems,

        # Use this when you only want list definitions, columns, and content
        # types. List item field payloads are skipped.
        [switch]$ListMetadataOnly
    )

    if ($null -eq $Headers -or -not $Headers.ContainsKey('Authorization')) {
        throw 'The supplied headers do not contain an Authorization header.'
    }

    if ($ApiMode -eq 'SharePointV2' -and [string]::IsNullOrWhiteSpace($TenantName)) {
        throw '-TenantName is required when -ApiMode is SharePointV2.'
    }

    $graphBase = 'https://graph.microsoft.com/v1.0'
    $requestedTypes = @($ManifestType | Select-Object -Unique)
    $includeAll     = $requestedTypes -contains 'All'
    $includeDrives  = $includeAll -or ($requestedTypes -contains 'Drives')
    $includeLists   = $includeAll -or ($requestedTypes -contains 'Lists')

    if ($includeAll) {
        $resolvedTypes = @('Sites', 'Drives', 'Lists')
    }
    else {
        $resolvedTypes = @($requestedTypes | Where-Object { $_ -ne 'All' })
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    Write-ManifestStatus `
        -Stopwatch $stopwatch `
        -Message (
            'Starting {0} manifest via {1}. Output: {2}' -f
            ($resolvedTypes -join ', '),
            $ApiMode,
            $OutputPath
        )

    $manifest = [ordered]@{
        schemaVersion  = '1.0'
        generatedAtUtc = [datetime]::UtcNow.ToString('o')
        apiMode        = $ApiMode
        manifestTypes  = $resolvedTypes
        options        = [ordered]@{
            includeDriveItems              = $includeDrives
            includeListSchema              = $includeLists
            includeListItems               = ($includeLists -and -not $ListMetadataOnly)
            includeDocumentLibraryListItems = [bool]$IncludeDocumentLibraryListItems
        }
        discovery      = $null

        counts = [ordered]@{
            sites          = 0
            drives         = 0
            driveItems     = 0
            lists          = 0
            listItems      = 0
            skippedLists   = 0
            errors         = 0
        }

        sites  = [System.Collections.Generic.List[object]]::new()
        errors = [System.Collections.Generic.List[object]]::new()
    }

    #region Site discovery

    if ($ApiMode -eq 'Graph') {
        $siteDiscoveryUri = "$graphBase/sites/getAllSites"

        Write-ManifestStatus `
            -Stopwatch $stopwatch `
            -Message "Discovering sites: $siteDiscoveryUri"

        try {
            $siteResponse = Invoke-PagedMetadataRequest `
                -Uri $siteDiscoveryUri `
                -Headers $Headers `
                -StatusStopwatch $stopwatch `
                -StatusLabel 'Site discovery'

            $manifest.discovery = [ordered]@{
                method = 'getAllSites'
                uri    = $siteDiscoveryUri
            }
        }
        catch {
            $statusCode = Get-HttpStatusCode -ErrorRecord $_

            # getAllSites is application-permission only. A delegated token
            # can instead use the site search endpoint.
            if ($statusCode -notin @(400, 403)) {
                throw
            }

            $siteDiscoveryUri = "$graphBase/sites?search=%2A"

            Write-ManifestStatus `
                -Stopwatch $stopwatch `
                -Message "Falling back to delegated site search: $siteDiscoveryUri" `
                -ForegroundColor Yellow

            $siteResponse = Invoke-PagedMetadataRequest `
                -Uri $siteDiscoveryUri `
                -Headers $Headers `
                -StatusStopwatch $stopwatch `
                -StatusLabel 'Site discovery fallback'

            $manifest.discovery = [ordered]@{
                method = 'search=* fallback'
                uri    = $siteDiscoveryUri
            }
        }
    }
    else {
        $sharePointBase = (
            'https://{0}.sharepoint.com/_api/v2.0' -f $TenantName
        )

        $siteDiscoveryUri = "$sharePointBase/sites"

        Write-ManifestStatus `
            -Stopwatch $stopwatch `
            -Message "Discovering sites: $siteDiscoveryUri"

        $siteResponse = Invoke-PagedMetadataRequest `
            -Uri $siteDiscoveryUri `
            -Headers $Headers `
            -StatusStopwatch $stopwatch `
            -StatusLabel 'Site discovery'

        $manifest.discovery = [ordered]@{
            method = 'SharePoint REST v2 sites'
            uri    = $siteDiscoveryUri
        }
    }

    #endregion Site discovery

    $totalSites = @($siteResponse.Items).Count

    Write-ManifestStatus `
        -Stopwatch $stopwatch `
        -Message "Discovered $totalSites site(s)."

    $siteIndex = 0

    foreach ($site in $siteResponse.Items) {
        $siteIndex++
        $manifest.counts.sites++
        $siteLabel = Get-ManifestSiteLabel -Site $site
        $etaText   = Get-ManifestEtaText `
            -Stopwatch $stopwatch `
            -Completed ($siteIndex - 1) `
            -Total $totalSites

        if ([string]::IsNullOrWhiteSpace($etaText)) {
            $etaText = 'ETA calculating'
        }

        Write-ManifestStatus `
            -Stopwatch $stopwatch `
            -Message (
                'Site {0}/{1}: {2} ({3})' -f
                $siteIndex,
                $totalSites,
                $siteLabel,
                $etaText
            )

        $siteEntry = [ordered]@{
            metadata = $site
            drives   = [System.Collections.Generic.List[object]]::new()
            lists    = [System.Collections.Generic.List[object]]::new()
            errors   = [System.Collections.Generic.List[object]]::new()
        }

        if ($ApiMode -eq 'Graph') {
            $siteApiBase  = "$graphBase/sites/$($site.id)"
            $driveApiBase = $graphBase
        }
        else {
            if ([string]::IsNullOrWhiteSpace($site.webUrl)) {
                $errorData = [ordered]@{
                    uri        = $siteDiscoveryUri
                    statusCode = $null
                    message    = 'The site response did not contain webUrl.'
                }

                Write-ManifestStatus `
                    -Stopwatch $stopwatch `
                    -Message "Skipping site without webUrl: $siteLabel" `
                    -ForegroundColor Yellow

                $siteEntry.errors.Add($errorData)
                $manifest.counts.errors++
                $manifest.sites.Add($siteEntry)
                continue
            }

            # REST v2 is scoped using the individual site's URL.
            $siteWebUrl   = $site.webUrl.TrimEnd('/')
            $siteApiBase  = "$siteWebUrl/_api/v2.0"
            $driveApiBase = $siteApiBase
        }

        #region Drives and drive items

        if ($includeDrives) {
            $drivesUri = "$siteApiBase/drives"

            Write-ManifestStatus `
                -Stopwatch $stopwatch `
                -Message "  Discovering drives for site: $siteLabel"

            try {
                $driveResponse = Invoke-PagedMetadataRequest `
                    -Uri $drivesUri `
                    -Headers $Headers `
                    -StatusStopwatch $stopwatch `
                    -StatusLabel "  Drives for $siteLabel"

                $totalDrives = @($driveResponse.Items).Count

                Write-ManifestStatus `
                    -Stopwatch $stopwatch `
                    -Message "  Found $totalDrives drive(s)."

                $driveIndex = 0

                foreach ($drive in $driveResponse.Items) {
                    $driveIndex++
                    $manifest.counts.drives++
                    $driveLabel = $drive.name

                    if ([string]::IsNullOrWhiteSpace($driveLabel)) {
                        $driveLabel = $drive.id
                    }

                    $driveEntry = [ordered]@{
                        metadata  = $drive
                        items     = @()
                        deltaLink = $null
                        error     = $null
                    }

                    $driveItemsUri = (
                        "$driveApiBase/drives/$($drive.id)/root/delta"
                    )

                    Write-ManifestStatus `
                        -Stopwatch $stopwatch `
                        -Message (
                            '  Drive {0}/{1}: {2} - enumerating metadata' -f
                            $driveIndex,
                            $totalDrives,
                            $driveLabel
                        )

                    try {
                        $driveItemsResponse = Invoke-PagedMetadataRequest `
                            -Uri $driveItemsUri `
                            -Headers $Headers `
                            -StatusStopwatch $stopwatch `
                            -StatusLabel "  Drive metadata for $driveLabel"

                        $driveEntry.items     = $driveItemsResponse.Items
                        $driveEntry.deltaLink = $driveItemsResponse.DeltaLink

                        $manifest.counts.driveItems += (
                            $driveItemsResponse.Items.Count
                        )

                        Write-ManifestStatus `
                            -Stopwatch $stopwatch `
                            -Message (
                                '  Drive {0}/{1}: {2} - {3} item(s)' -f
                                $driveIndex,
                                $totalDrives,
                                $driveLabel,
                                $driveItemsResponse.Items.Count
                            )
                    }
                    catch {
                        $driveEntry.error = New-ManifestError `
                            -Uri $driveItemsUri `
                            -ErrorRecord $_

                        $manifest.counts.errors++

                        Write-ManifestStatus `
                            -Stopwatch $stopwatch `
                            -Message (
                                '  Drive {0}/{1}: {2} - failed to enumerate items' -f
                                $driveIndex,
                                $totalDrives,
                                $driveLabel
                            ) `
                            -ForegroundColor Yellow
                    }

                    $siteEntry.drives.Add($driveEntry)
                }
            }
            catch {
                $siteEntry.errors.Add(
                    (New-ManifestError -Uri $drivesUri -ErrorRecord $_)
                )

                $manifest.counts.errors++

                Write-ManifestStatus `
                    -Stopwatch $stopwatch `
                    -Message "  Failed to discover drives for site: $siteLabel" `
                    -ForegroundColor Yellow
            }
        }

        #endregion Drives and drive items

        #region Lists and list items

        # Including "system" in $select causes system lists to be returned too.
        $listSelect = @(
            'id'
            'name'
            'displayName'
            'description'
            'webUrl'
            'createdDateTime'
            'lastModifiedDateTime'
            'list'
            'system'
            'sharepointIds'
        ) -join ','

        if ($includeLists) {
            $listsUri = "$siteApiBase/lists?`$select=$listSelect"

            Write-ManifestStatus `
                -Stopwatch $stopwatch `
                -Message "  Discovering lists for site: $siteLabel"

            try {
                $listResponse = Invoke-PagedMetadataRequest `
                    -Uri $listsUri `
                    -Headers $Headers `
                    -StatusStopwatch $stopwatch `
                    -StatusLabel "  Lists for $siteLabel"

                $totalLists = @($listResponse.Items).Count

                Write-ManifestStatus `
                    -Stopwatch $stopwatch `
                    -Message "  Found $totalLists list(s)."

                $listIndex = 0

                foreach ($list in $listResponse.Items) {
                    $listIndex++
                    $manifest.counts.lists++
                    $listLabel = $list.displayName

                    if ([string]::IsNullOrWhiteSpace($listLabel)) {
                        $listLabel = $list.name
                    }

                    if ([string]::IsNullOrWhiteSpace($listLabel)) {
                        $listLabel = $list.id
                    }

                    $isDocumentLibrary = (
                        $list.list.template -eq 'documentLibrary'
                    )

                    $skipItems = (
                        $ListMetadataOnly -or
                        (
                            $isDocumentLibrary -and
                            -not $IncludeDocumentLibraryListItems
                        )
                    )

                    $listEntry = [ordered]@{
                        metadata               = $list
                        columns                = @()
                        contentTypes           = @()
                        items                  = @()
                        itemEnumerationSkipped = $skipItems
                        skipReason             = $null
                        errors                 = [System.Collections.Generic.List[object]]::new()
                    }

                    $listColumnsUri = "$siteApiBase/lists/$($list.id)/columns"

                    Write-ManifestStatus `
                        -Stopwatch $stopwatch `
                        -Message (
                            '  List {0}/{1}: {2} - schema metadata' -f
                            $listIndex,
                            $totalLists,
                            $listLabel
                        )

                    try {
                        $listColumnsResponse = Invoke-PagedMetadataRequest `
                            -Uri $listColumnsUri `
                            -Headers $Headers

                        $listEntry.columns = $listColumnsResponse.Items
                    }
                    catch {
                        $listEntry.errors.Add(
                            (New-ManifestError -Uri $listColumnsUri -ErrorRecord $_)
                        )

                        $manifest.counts.errors++

                        Write-ManifestStatus `
                            -Stopwatch $stopwatch `
                            -Message (
                                '  List {0}/{1}: {2} - failed to fetch columns' -f
                                $listIndex,
                                $totalLists,
                                $listLabel
                            ) `
                            -ForegroundColor Yellow
                    }

                    $listContentTypesUri = "$siteApiBase/lists/$($list.id)/contentTypes"

                    try {
                        $listContentTypesResponse = Invoke-PagedMetadataRequest `
                            -Uri $listContentTypesUri `
                            -Headers $Headers

                        $listEntry.contentTypes = $listContentTypesResponse.Items
                    }
                    catch {
                        $listEntry.errors.Add(
                            (New-ManifestError -Uri $listContentTypesUri -ErrorRecord $_)
                        )

                        $manifest.counts.errors++

                        Write-ManifestStatus `
                            -Stopwatch $stopwatch `
                            -Message (
                                '  List {0}/{1}: {2} - failed to fetch content types' -f
                                $listIndex,
                                $totalLists,
                                $listLabel
                            ) `
                            -ForegroundColor Yellow
                    }

                    if ($skipItems) {
                        if ($ListMetadataOnly) {
                            $listEntry.skipReason = 'List item enumeration was not requested.'
                        }
                        else {
                            $listEntry.skipReason = (
                                'Document library list items are skipped by ' +
                                'default because file metadata is normally ' +
                                'captured by the drive manifest. Re-run with ' +
                                '-IncludeDocumentLibraryListItems to include ' +
                                'their listItem/fields payloads too.'
                            )
                        }

                        $manifest.counts.skippedLists++
                        $siteEntry.lists.Add($listEntry)

                        Write-ManifestStatus `
                            -Stopwatch $stopwatch `
                            -Message (
                                '  List {0}/{1}: {2} - item enumeration skipped' -f
                                $listIndex,
                                $totalLists,
                                $listLabel
                            )

                        continue
                    }

                    $listItemsUri = (
                        "$siteApiBase/lists/$($list.id)/items?`$expand=fields"
                    )

                    Write-ManifestStatus `
                        -Stopwatch $stopwatch `
                        -Message (
                            '  List {0}/{1}: {2} - enumerating item fields' -f
                            $listIndex,
                            $totalLists,
                            $listLabel
                        )

                    try {
                        $listItemsResponse = Invoke-PagedMetadataRequest `
                            -Uri $listItemsUri `
                            -Headers $Headers `
                            -StatusStopwatch $stopwatch `
                            -StatusLabel "  Item fields for $listLabel"

                        $listEntry.items = $listItemsResponse.Items

                        $manifest.counts.listItems += (
                            $listItemsResponse.Items.Count
                        )

                        Write-ManifestStatus `
                            -Stopwatch $stopwatch `
                            -Message (
                                '  List {0}/{1}: {2} - {3} item(s)' -f
                                $listIndex,
                                $totalLists,
                                $listLabel,
                                $listItemsResponse.Items.Count
                            )
                    }
                    catch {
                        $listEntry.errors.Add(
                            (New-ManifestError -Uri $listItemsUri -ErrorRecord $_)
                        )

                        $manifest.counts.errors++

                        Write-ManifestStatus `
                            -Stopwatch $stopwatch `
                            -Message (
                                '  List {0}/{1}: {2} - failed to enumerate items' -f
                                $listIndex,
                                $totalLists,
                                $listLabel
                            ) `
                            -ForegroundColor Yellow
                    }

                    $siteEntry.lists.Add($listEntry)
                }
            }
            catch {
                $siteEntry.errors.Add(
                    (New-ManifestError -Uri $listsUri -ErrorRecord $_)
                )

                $manifest.counts.errors++

                Write-ManifestStatus `
                    -Stopwatch $stopwatch `
                    -Message "  Failed to discover lists for site: $siteLabel" `
                    -ForegroundColor Yellow
            }
        }

        #endregion Lists and list items

        $manifest.sites.Add($siteEntry)

        $completedEtaText = Get-ManifestEtaText `
            -Stopwatch $stopwatch `
            -Completed $siteIndex `
            -Total $totalSites

        if ([string]::IsNullOrWhiteSpace($completedEtaText)) {
            $completedEtaText = 'ETA complete'
        }

        Write-ManifestStatus `
            -Stopwatch $stopwatch `
            -Message (
                'Completed site {0}/{1}: {2}. Totals: {3} drive(s), {4} drive item(s), {5} list(s), {6} list item(s), {7} error(s). {8}' -f
                $siteIndex,
                $totalSites,
                $siteLabel,
                $manifest.counts.drives,
                $manifest.counts.driveItems,
                $manifest.counts.lists,
                $manifest.counts.listItems,
                $manifest.counts.errors,
                $completedEtaText
            ) `
            -ForegroundColor Green
    }

    if ([System.IO.Path]::IsPathRooted($OutputPath)) {
        $fullOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
    }
    else {
        $baseOutputDirectory = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
            $PWD.Path
        }
        else {
            $PSScriptRoot
        }

        $fullOutputPath = [System.IO.Path]::GetFullPath(
            (Join-Path -Path $baseOutputDirectory -ChildPath $OutputPath)
        )
    }

    Write-ManifestStatus `
        -Stopwatch $stopwatch `
        -Message "Serializing manifest JSON..."

    $json = $manifest | ConvertTo-Json -Depth 100

    try {
        $outputDirectory = Split-Path -Parent $fullOutputPath

        if (-not (Test-Path -LiteralPath $outputDirectory)) {
            $null = New-Item `
                -ItemType Directory `
                -Path $outputDirectory `
                -Force `
                -ErrorAction Stop
        }

        Write-ManifestStatus `
            -Stopwatch $stopwatch `
            -Message "Writing manifest: $fullOutputPath"

        [System.IO.File]::WriteAllText(
            $fullOutputPath,
            $json,
            [System.Text.UTF8Encoding]::new($false)
        )
    }
    catch {
        $failedOutputPath = $fullOutputPath
        $fallbackDirectory = Join-Path `
            -Path ([System.IO.Path]::GetTempPath()) `
            -ChildPath 'sharepoint-manifests'
        $fallbackFileName = [System.IO.Path]::GetFileName($failedOutputPath)

        if ([string]::IsNullOrWhiteSpace($fallbackFileName)) {
            $fallbackFileName = 'graph-sharepoint-manifest.json'
        }

        $fullOutputPath = Join-Path `
            -Path $fallbackDirectory `
            -ChildPath $fallbackFileName

        Write-ManifestStatus `
            -Stopwatch $stopwatch `
            -Message (
                'Could not write {0}: {1}. Trying fallback: {2}' -f
                $failedOutputPath,
                $_.Exception.Message,
                $fullOutputPath
            ) `
            -ForegroundColor Yellow

        if (-not (Test-Path -LiteralPath $fallbackDirectory)) {
            $null = New-Item `
                -ItemType Directory `
                -Path $fallbackDirectory `
                -Force `
                -ErrorAction Stop
        }

        [System.IO.File]::WriteAllText(
            $fullOutputPath,
            $json,
            [System.Text.UTF8Encoding]::new($false)
        )
    }

    $stopwatch.Stop()

    Write-ManifestStatus `
        -Stopwatch $stopwatch `
        -Message (
            'Done. Sites: {0}; Drives: {1}; Drive items: {2}; Lists: {3}; List items: {4}; Errors: {5}; Output: {6}' -f
            $manifest.counts.sites,
            $manifest.counts.drives,
            $manifest.counts.driveItems,
            $manifest.counts.lists,
            $manifest.counts.listItems,
            $manifest.counts.errors,
            $fullOutputPath
        ) `
        -ForegroundColor Green

    [pscustomobject]@{
        Path          = $fullOutputPath
        ApiMode       = $ApiMode
        ManifestTypes = $resolvedTypes
        Generated     = $manifest.generatedAtUtc
        Sites         = $manifest.counts.sites
        Drives        = $manifest.counts.drives
        DriveItems    = $manifest.counts.driveItems
        Lists         = $manifest.counts.lists
        ListItems     = $manifest.counts.listItems
        Errors        = $manifest.counts.errors
    }
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    if ($ManifestType -contains 'All') {
        $manifestSlug = 'all'
    }
    else {
        $manifestSlug = (
            @($ManifestType | ForEach-Object { $_.ToLowerInvariant() }) -join '-'
        )
    }

    $OutputPath = 'graph-sharepoint-{0}-manifest.json' -f $manifestSlug
}

$exportParams = @{
    Headers                         = $Headers
    ApiMode                         = $ApiMode
    ManifestType                    = $ManifestType
    OutputPath                      = $OutputPath
    IncludeDocumentLibraryListItems = $IncludeDocumentLibraryListItems
    ListMetadataOnly                = $ListMetadataOnly
}

if (-not [string]::IsNullOrWhiteSpace($TenantName)) {
    $exportParams.TenantName = $TenantName
}

Export-SharePointMetadataManifest @exportParams
}