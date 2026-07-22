function Resolve-SharePointManifestDirectory {
    [CmdletBinding()]
    param(
        [string]$ManifestDir = (Join-Path $PWD 'out\sharepoint-manifests'),

        [switch]$Create
    )

    if ([System.IO.Path]::IsPathRooted($ManifestDir)) {
        $resolvedPath = [System.IO.Path]::GetFullPath($ManifestDir)
    }
    else {
        $resolvedPath = [System.IO.Path]::GetFullPath(
            (Join-Path -Path $PWD -ChildPath $ManifestDir)
        )
    }

    if ($Create -and -not (Test-Path -LiteralPath $resolvedPath)) {
        $null = New-Item `
            -ItemType Directory `
            -Path $resolvedPath `
            -Force
    }

    return $resolvedPath
}

function Get-SharePointManifestSlug {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('All', 'Sites', 'Drives', 'Lists')]
        [string[]]$ManifestType
    )

    if ($ManifestType -contains 'All') {
        return 'all'
    }

    return (@($ManifestType | ForEach-Object {
        $_.ToLowerInvariant()
    }) -join '-')
}

function Get-SharePointManifestHttpStatusCode {
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

function New-SharePointManifestError {
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    [ordered]@{
        uri        = $Uri
        statusCode = Get-SharePointManifestHttpStatusCode -ErrorRecord $ErrorRecord
        message    = $ErrorRecord.Exception.Message
    }
}

function Invoke-SharePointManifestPagedRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [scriptblock]$RefreshHeaders,

        [ValidateRange(0, 20)]
        [int]$MaxRetries = 6,

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
                if ($RefreshHeaders) {
                    $Headers = & $RefreshHeaders
                }

                $response = Invoke-RestMethod `
                    -Method Get `
                    -Uri $nextUri `
                    -Headers $Headers `
                    -ErrorAction Stop

                break
            }
            catch {
                $statusCode = Get-SharePointManifestHttpStatusCode -ErrorRecord $_
                $isTransient = $statusCode -in @(429, 502, 503, 504)

                if (-not $isTransient -or $attempt -ge $MaxRetries) {
                    throw
                }

                $delaySeconds = [math]::Min(60, [math]::Pow(2, $attempt + 1))

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
            $items.Add($response)
        }

        if ($response.'@odata.deltaLink') {
            $deltaLink = $response.'@odata.deltaLink'
        }

        $pageCount++
        $nextUri = $response.'@odata.nextLink'

        if (
            -not [string]::IsNullOrWhiteSpace($StatusLabel) -and
            (
                $pageCount -eq 1 -or
                $pageCount % $StatusPageInterval -eq 0 -or
                -not $nextUri
            )
        ) {
            $statusSuffix = if ($nextUri) { 'continuing' } else { 'done' }

            Write-Host (
                "{0}: page {1}, {2} item(s) so far, {3}" -f
                $StatusLabel,
                $pageCount,
                $items.Count,
                $statusSuffix
            ) -ForegroundColor DarkCyan
        }
    }

    [pscustomobject]@{
        Items     = $items.ToArray()
        DeltaLink = $deltaLink
    }
}

function Resolve-SharePointManifestOutputPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    if ([System.IO.Path]::IsPathRooted($OutputPath)) {
        return [System.IO.Path]::GetFullPath($OutputPath)
    }

    [System.IO.Path]::GetFullPath((Join-Path -Path $PWD -ChildPath $OutputPath))
}

function ConvertTo-SharePointManifestSafeFileName {
    [CmdletBinding()]
    param(
        [string]$Value,

        [string]$Fallback = 'site'
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        $Value = $Fallback
    }

    $safeValue = $Value

    foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safeValue = $safeValue.Replace($invalidChar, '-')
    }

    $safeValue = ($safeValue -replace '\s+', '-').Trim('-')

    if ([string]::IsNullOrWhiteSpace($safeValue)) {
        $safeValue = $Fallback
    }

    if ($safeValue.Length -gt 90) {
        $safeValue = $safeValue.Substring(0, 90).Trim('-')
    }

    return $safeValue
}

function Save-SharePointSiteManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$SiteEntry,

        [Parameter(Mandatory)]
        [int]$SiteIndex,

        [Parameter(Mandatory)]
        [int]$TotalSites,

        [Parameter(Mandatory)]
        [string]$SiteManifestDirectory,

        [Parameter(Mandatory)]
        [string]$IndexDirectory,

        [Parameter(Mandatory)]
        [string]$GeneratedAtUtc,

        [Parameter(Mandatory)]
        [string]$ApiMode,

        [Parameter(Mandatory)]
        [string[]]$ManifestTypes
    )

    $site = $SiteEntry.metadata
    $siteLabel = $site.displayName ?? $site.name ?? $site.webUrl ?? $site.id
    $safeSiteLabel = ConvertTo-SharePointManifestSafeFileName `
        -Value $siteLabel `
        -Fallback ('site-{0:0000}' -f $SiteIndex)
    $siteFileName = '{0:0000}-{1}.json' -f $SiteIndex, $safeSiteLabel
    $sitePath = Join-Path -Path $SiteManifestDirectory -ChildPath $siteFileName

    $driveCount = @($SiteEntry.drives).Count
    $driveItemCount = 0
    $driveErrorCount = 0

    foreach ($driveEntry in @($SiteEntry.drives)) {
        $driveItemCount += @($driveEntry.items).Count

        if ($null -ne $driveEntry.error) {
            $driveErrorCount++
        }
    }

    $listCount = @($SiteEntry.lists).Count
    $listItemCount = 0
    $skippedListCount = 0
    $listErrorCount = 0

    foreach ($listEntry in @($SiteEntry.lists)) {
        $listItemCount += @($listEntry.items).Count

        if ($listEntry.itemEnumerationSkipped) {
            $skippedListCount++
        }

        $listErrorCount += @($listEntry.errors).Count
    }

    $siteErrorCount = @($SiteEntry.errors).Count + $driveErrorCount + $listErrorCount
    $siteCounts = [ordered]@{
        sites        = 1
        drives       = $driveCount
        driveItems   = $driveItemCount
        lists        = $listCount
        listItems    = $listItemCount
        skippedLists = $skippedListCount
        errors       = $siteErrorCount
    }

    $siteManifest = [ordered]@{
        schemaVersion  = '1.1'
        layout         = 'Site'
        generatedAtUtc = $GeneratedAtUtc
        apiMode        = $ApiMode
        manifestTypes  = $ManifestTypes
        counts         = $siteCounts
        sites          = @($SiteEntry)
    }

    Write-Host "  Writing site manifest: $sitePath" -ForegroundColor DarkCyan

    [System.IO.File]::WriteAllText(
        $sitePath,
        ($siteManifest | ConvertTo-Json -Depth 100),
        [System.Text.UTF8Encoding]::new($false)
    )

    $relativePath = Join-Path `
        -Path (Split-Path -Leaf $SiteManifestDirectory) `
        -ChildPath $siteFileName

    [pscustomobject]@{
        siteId       = $site.id
        siteName     = $siteLabel
        siteWebUrl   = $site.webUrl
        siteIndex    = $SiteIndex
        totalSites   = $TotalSites
        path         = $sitePath
        relativePath = $relativePath
        counts       = [pscustomobject]$siteCounts
    }
}

function Export-SharePointMetadataManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Headers,

        [scriptblock]$RefreshHeaders,

        [ValidateSet('Graph', 'SharePointV2')]
        [string]$ApiMode = 'Graph',

        [string]$TenantName,

        [ValidateSet('All', 'Sites', 'Drives', 'Lists')]
        [string[]]$ManifestType = @('All'),

        [Parameter(Mandatory)]
        [string]$OutputPath,

        [switch]$IncludeDocumentLibraryListItems,

        [switch]$ListMetadataOnly,

        [ValidateRange(0, [int]::MaxValue)]
        [int]$MaxSites = 0,

        [switch]$FirstSiteOnly
    )

    if ($null -eq $Headers -or -not $Headers.ContainsKey('Authorization')) {
        throw 'Generating SharePoint manifests requires headers with an Authorization value.'
    }

    if ($ApiMode -eq 'SharePointV2' -and [string]::IsNullOrWhiteSpace($TenantName)) {
        throw '-TenantName is required when -ApiMode is SharePointV2.'
    }

    $graphBase      = 'https://graph.microsoft.com/v1.0'
    $requestedTypes = @($ManifestType | Select-Object -Unique)
    $includeAll     = $requestedTypes -contains 'All'
    $includeDrives  = $includeAll -or ($requestedTypes -contains 'Drives')
    $includeLists   = $includeAll -or ($requestedTypes -contains 'Lists')
    $resolvedTypes  = if ($includeAll) {
        @('Sites', 'Drives', 'Lists')
    }
    else {
        @($requestedTypes | Where-Object { $_ -ne 'All' })
    }
    $fullOutputPath = Resolve-SharePointManifestOutputPath -OutputPath $OutputPath
    $outputDirectory = Split-Path -Parent $fullOutputPath
    $outputBaseName = [System.IO.Path]::GetFileNameWithoutExtension($fullOutputPath)
    $siteManifestDirectory = Join-Path `
        -Path $outputDirectory `
        -ChildPath ('{0}-sites' -f $outputBaseName)

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        $null = New-Item -ItemType Directory -Path $outputDirectory -Force
    }

    if (-not (Test-Path -LiteralPath $siteManifestDirectory)) {
        $null = New-Item -ItemType Directory -Path $siteManifestDirectory -Force
    }

    $manifest = [ordered]@{
        schemaVersion  = '1.1'
        layout         = 'PerSite'
        generatedAtUtc = [datetime]::UtcNow.ToString('o')
        apiMode        = $ApiMode
        manifestTypes  = $resolvedTypes
        options        = [ordered]@{
            includeDriveItems               = $includeDrives
            includeListSchema               = $includeLists
            includeListItems                = ($includeLists -and -not $ListMetadataOnly)
            includeDocumentLibraryListItems = [bool]$IncludeDocumentLibraryListItems
        }
        discovery = $null
        counts    = [ordered]@{
            sites        = 0
            drives       = 0
            driveItems   = 0
            lists        = 0
            listItems    = 0
            skippedLists = 0
            errors       = 0
        }
        siteManifestDirectory = $siteManifestDirectory
        siteManifests         = [System.Collections.Generic.List[object]]::new()
        sites                 = @()
        errors                = [System.Collections.Generic.List[object]]::new()
    }

    Write-Host "Discovering SharePoint sites..." -ForegroundColor Cyan

    if ($ApiMode -eq 'Graph') {
        $siteDiscoveryUri = "$graphBase/sites/getAllSites"

        try {
            $siteResponse = Invoke-SharePointManifestPagedRequest `
                -Uri $siteDiscoveryUri `
                -Headers $Headers `
                -RefreshHeaders $RefreshHeaders `
                -StatusLabel 'Site discovery'

            $manifest.discovery = [ordered]@{
                method = 'getAllSites'
                uri    = $siteDiscoveryUri
            }
        }
        catch {
            $statusCode = Get-SharePointManifestHttpStatusCode -ErrorRecord $_

            if ($statusCode -notin @(400, 403)) {
                throw
            }

            $siteDiscoveryUri = "$graphBase/sites?search=%2A"
            Write-Host "Falling back to delegated site search." -ForegroundColor Yellow

            $siteResponse = Invoke-SharePointManifestPagedRequest `
                -Uri $siteDiscoveryUri `
                -Headers $Headers `
                -RefreshHeaders $RefreshHeaders `
                -StatusLabel 'Site discovery fallback'

            $manifest.discovery = [ordered]@{
                method = 'search=* fallback'
                uri    = $siteDiscoveryUri
            }
        }
    }
    else {
        $sharePointBase  = 'https://{0}.sharepoint.com/_api/v2.0' -f $TenantName
        $siteDiscoveryUri = "$sharePointBase/sites"

        $siteResponse = Invoke-SharePointManifestPagedRequest `
            -Uri $siteDiscoveryUri `
            -Headers $Headers `
            -RefreshHeaders $RefreshHeaders `
            -StatusLabel 'Site discovery'

        $manifest.discovery = [ordered]@{
            method = 'SharePoint REST v2 sites'
            uri    = $siteDiscoveryUri
        }
    }

    if ($FirstSiteOnly) {
        $MaxSites = 1
    }

    $discoveredSiteCount = @($siteResponse.Items).Count
    $sitesToProcess = if ($MaxSites -gt 0) {
        @($siteResponse.Items | Select-Object -First $MaxSites)
    }
    else {
        @($siteResponse.Items)
    }

    $totalSites = @($sitesToProcess).Count

    if ($MaxSites -gt 0) {
        Write-Host "Discovered $discoveredSiteCount site(s). Test run will process first $totalSites." -ForegroundColor Yellow
    }
    else {
        Write-Host "Discovered $totalSites site(s)." -ForegroundColor Cyan
    }

    $siteIndex = 0

    foreach ($site in $sitesToProcess) {
        $siteIndex++
        $manifest.counts.sites++

        $siteLabel = $site.displayName ?? $site.name ?? $site.webUrl ?? $site.id
        Write-Host "Site $siteIndex/$totalSites`: $siteLabel" -ForegroundColor Cyan

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
                $siteEntry.errors.Add([ordered]@{
                    uri        = $siteDiscoveryUri
                    statusCode = $null
                    message    = 'The site response did not contain webUrl.'
                })
                $manifest.counts.errors++
                $siteManifestIndex = Save-SharePointSiteManifest `
                    -SiteEntry $siteEntry `
                    -SiteIndex $siteIndex `
                    -TotalSites $totalSites `
                    -SiteManifestDirectory $siteManifestDirectory `
                    -IndexDirectory $outputDirectory `
                    -GeneratedAtUtc $manifest.generatedAtUtc `
                    -ApiMode $ApiMode `
                    -ManifestTypes $resolvedTypes
                $manifest.siteManifests.Add($siteManifestIndex)
                continue
            }

            $siteWebUrl   = $site.webUrl.TrimEnd('/')
            $siteApiBase  = "$siteWebUrl/_api/v2.0"
            $driveApiBase = $siteApiBase
        }

        if ($includeDrives) {
            $drivesUri = "$siteApiBase/drives"

            try {
                $driveResponse = Invoke-SharePointManifestPagedRequest `
                    -Uri $drivesUri `
                    -Headers $Headers `
                    -RefreshHeaders $RefreshHeaders `
                    -StatusLabel "Drives for $siteLabel"

                $totalDrives = @($driveResponse.Items).Count
                $driveIndex = 0

                foreach ($drive in $driveResponse.Items) {
                    $driveIndex++
                    $manifest.counts.drives++
                    $driveLabel = $drive.name ?? $drive.id

                    Write-Host "  Drive $driveIndex/$totalDrives`: $driveLabel" -ForegroundColor DarkCyan

                    $driveEntry = [ordered]@{
                        metadata  = $drive
                        items     = @()
                        deltaLink = $null
                        error     = $null
                    }

                    $driveItemsUri = "$driveApiBase/drives/$($drive.id)/root/delta"

                    try {
                        $driveItemsResponse = Invoke-SharePointManifestPagedRequest `
                            -Uri $driveItemsUri `
                            -Headers $Headers `
                            -RefreshHeaders $RefreshHeaders `
                            -StatusLabel "Drive metadata for $driveLabel"

                        $driveEntry.items     = $driveItemsResponse.Items
                        $driveEntry.deltaLink = $driveItemsResponse.DeltaLink
                        $manifest.counts.driveItems += $driveItemsResponse.Items.Count
                    }
                    catch {
                        $driveEntry.error = New-SharePointManifestError `
                            -Uri $driveItemsUri `
                            -ErrorRecord $_

                        $manifest.counts.errors++
                    }

                    $siteEntry.drives.Add($driveEntry)
                }
            }
            catch {
                $siteEntry.errors.Add(
                    (New-SharePointManifestError -Uri $drivesUri -ErrorRecord $_)
                )
                $manifest.counts.errors++
            }
        }

        if ($includeLists) {
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

            $listsUri = "$siteApiBase/lists?`$select=$listSelect"

            try {
                $listResponse = Invoke-SharePointManifestPagedRequest `
                    -Uri $listsUri `
                    -Headers $Headers `
                    -RefreshHeaders $RefreshHeaders `
                    -StatusLabel "Lists for $siteLabel"

                $totalLists = @($listResponse.Items).Count
                $listIndex = 0

                foreach ($list in $listResponse.Items) {
                    $listIndex++
                    $manifest.counts.lists++
                    $listLabel = $list.displayName ?? $list.name ?? $list.id

                    Write-Host "  List $listIndex/$totalLists`: $listLabel" -ForegroundColor DarkCyan

                    $isDocumentLibrary = $list.list.template -eq 'documentLibrary'
                    $skipItems = (
                        $ListMetadataOnly -or
                        ($isDocumentLibrary -and -not $IncludeDocumentLibraryListItems)
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
                    try {
                        $listEntry.columns = (
                            Invoke-SharePointManifestPagedRequest `
                                -Uri $listColumnsUri `
                                -Headers $Headers `
                                -RefreshHeaders $RefreshHeaders
                        ).Items
                    }
                    catch {
                        $listEntry.errors.Add(
                            (New-SharePointManifestError -Uri $listColumnsUri -ErrorRecord $_)
                        )
                        $manifest.counts.errors++
                    }

                    $listContentTypesUri = "$siteApiBase/lists/$($list.id)/contentTypes"
                    try {
                        $listEntry.contentTypes = (
                            Invoke-SharePointManifestPagedRequest `
                                -Uri $listContentTypesUri `
                                -Headers $Headers `
                                -RefreshHeaders $RefreshHeaders
                        ).Items
                    }
                    catch {
                        $listEntry.errors.Add(
                            (New-SharePointManifestError -Uri $listContentTypesUri -ErrorRecord $_)
                        )
                        $manifest.counts.errors++
                    }

                    if ($skipItems) {
                        $listEntry.skipReason = if ($ListMetadataOnly) {
                            'List item enumeration was not requested.'
                        }
                        else {
                            'Document library list items are skipped by default because file metadata is normally captured by the drive manifest.'
                        }

                        $manifest.counts.skippedLists++
                        $siteEntry.lists.Add($listEntry)
                        continue
                    }

                    $listItemsUri = "$siteApiBase/lists/$($list.id)/items?`$expand=fields"
                    try {
                        $listItemsResponse = Invoke-SharePointManifestPagedRequest `
                            -Uri $listItemsUri `
                            -Headers $Headers `
                            -RefreshHeaders $RefreshHeaders `
                            -StatusLabel "Item fields for $listLabel"

                        $listEntry.items = $listItemsResponse.Items
                        $manifest.counts.listItems += $listItemsResponse.Items.Count
                    }
                    catch {
                        $listEntry.errors.Add(
                            (New-SharePointManifestError -Uri $listItemsUri -ErrorRecord $_)
                        )
                        $manifest.counts.errors++
                    }

                    $siteEntry.lists.Add($listEntry)
                }
            }
            catch {
                $siteEntry.errors.Add(
                    (New-SharePointManifestError -Uri $listsUri -ErrorRecord $_)
                )
                $manifest.counts.errors++
            }
        }

        $siteManifestIndex = Save-SharePointSiteManifest `
            -SiteEntry $siteEntry `
            -SiteIndex $siteIndex `
            -TotalSites $totalSites `
            -SiteManifestDirectory $siteManifestDirectory `
            -IndexDirectory $outputDirectory `
            -GeneratedAtUtc $manifest.generatedAtUtc `
            -ApiMode $ApiMode `
            -ManifestTypes $resolvedTypes
        $manifest.siteManifests.Add($siteManifestIndex)
    }

    $fullOutputPath = Resolve-SharePointManifestOutputPath -OutputPath $OutputPath
    $outputDirectory = Split-Path -Parent $fullOutputPath

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        $null = New-Item -ItemType Directory -Path $outputDirectory -Force
    }

    Write-Host "Writing SharePoint manifest: $fullOutputPath" -ForegroundColor Cyan

    [System.IO.File]::WriteAllText(
        $fullOutputPath,
        ($manifest | ConvertTo-Json -Depth 100),
        [System.Text.UTF8Encoding]::new($false)
    )

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

function Import-SharePointManifestJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "SharePoint manifest not found: $Path"
    }

    Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json
}

function Import-SharePointManifestSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ManifestPath,

        [string]$ManifestDir
    )

    $paths = @($ManifestPath | ForEach-Object {
        [System.IO.Path]::GetFullPath($_)
    })

    $manifestList = [System.Collections.Generic.List[object]]::new()
    $countSourceManifestList = [System.Collections.Generic.List[object]]::new()
    $allPathList = [System.Collections.Generic.List[string]]::new()

    foreach ($path in $paths) {
        $manifest = Import-SharePointManifestJson -Path $path
        $allPathList.Add($path)

        if ($manifest.layout -eq 'PerSite' -and $manifest.siteManifests) {
            $countSourceManifestList.Add($manifest)
            $indexDirectory = Split-Path -Parent $path

            foreach ($siteManifestRef in @($manifest.siteManifests)) {
                $siteManifestPath = $siteManifestRef.path

                if (
                    [string]::IsNullOrWhiteSpace($siteManifestPath) -or
                    -not (Test-Path -LiteralPath $siteManifestPath -PathType Leaf)
                ) {
                    $siteManifestPath = Join-Path `
                        -Path $indexDirectory `
                        -ChildPath $siteManifestRef.relativePath
                }

                $siteManifestPath = [System.IO.Path]::GetFullPath($siteManifestPath)
                $siteManifest = Import-SharePointManifestJson -Path $siteManifestPath

                $manifestList.Add($siteManifest)
                $allPathList.Add($siteManifestPath)
            }
        }
        else {
            $manifestList.Add($manifest)
            $countSourceManifestList.Add($manifest)
        }
    }

    $counts = [ordered]@{
        sites      = 0
        drives     = 0
        driveItems = 0
        lists      = 0
        listItems  = 0
        errors     = 0
    }

    foreach ($manifest in $countSourceManifestList) {
        if ($manifest.counts) {
            $counts.sites      += [int]($manifest.counts.sites ?? 0)
            $counts.drives     += [int]($manifest.counts.drives ?? 0)
            $counts.driveItems += [int]($manifest.counts.driveItems ?? 0)
            $counts.lists      += [int]($manifest.counts.lists ?? 0)
            $counts.listItems  += [int]($manifest.counts.listItems ?? 0)
            $counts.errors     += [int]($manifest.counts.errors ?? 0)
        }
    }

    [pscustomobject]@{
        ManifestDir = $ManifestDir
        Paths       = $allPathList.ToArray()
        Manifests   = $manifestList.ToArray()
        Counts      = [pscustomobject]$counts
    }
}

function Initialize-SharePointManifestSet {
    [CmdletBinding()]
    param(
        [hashtable]$Headers = $GraphHeaders,

        [scriptblock]$RefreshHeaders,

        [ValidateSet('Graph', 'SharePointV2')]
        [string]$ApiMode = 'Graph',

        [string]$TenantName,

        [ValidateSet('All', 'Sites', 'Drives', 'Lists')]
        [string[]]$ManifestType = @('All'),

        [ValidateSet('Auto', 'Generate', 'UseExisting')]
        [string]$ManifestMode = 'Auto',

        [string]$ManifestDir = (Join-Path $PWD 'out\sharepoint-manifests'),

        [Alias('SharePointManifestSet', 'ManifestGeneratorPath')]
        [string]$GeneratorPath,

        [switch]$IncludeDocumentLibraryListItems,

        [switch]$ListMetadataOnly,

        [ValidateRange(0, [int]::MaxValue)]
        [int]$MaxSites = 0,

        [switch]$FirstSiteOnly,

        [switch]$Force
    )

    $manifestDirPath = Resolve-SharePointManifestDirectory `
        -ManifestDir $ManifestDir `
        -Create
    $manifestSlug = Get-SharePointManifestSlug -ManifestType $ManifestType
    $manifestPath = Join-Path `
        -Path $manifestDirPath `
        -ChildPath ('graph-sharepoint-{0}-manifest.json' -f $manifestSlug)

    $hasExistingManifest = Test-Path -LiteralPath $manifestPath -PathType Leaf
    $shouldGenerate = (
        $ManifestMode -eq 'Generate' -or
        $Force -or
        ($ManifestMode -eq 'Auto' -and -not $hasExistingManifest)
    )

    if ($ManifestMode -eq 'UseExisting' -and -not $hasExistingManifest) {
        throw "No cached SharePoint manifest exists at: $manifestPath"
    }

    if ($shouldGenerate) {
        if ($null -eq $Headers -or -not $Headers.ContainsKey('Authorization')) {
            throw 'Generating SharePoint manifests requires headers with an Authorization value.'
        }

        Write-Host "Generating SharePoint manifest: $manifestPath" -ForegroundColor Cyan

        if (-not [string]::IsNullOrWhiteSpace($GeneratorPath)) {
            Write-Warning (
                '-GeneratorPath/-SharePointManifestSet is ignored. ' +
                'Manifest generation is now handled by Export-SharePointMetadataManifest.'
            )
        }

        $generatorParams = @{
            Headers                         = $Headers
            RefreshHeaders                  = $RefreshHeaders
            ApiMode                         = $ApiMode
            ManifestType                    = $ManifestType
            OutputPath                      = $manifestPath
            IncludeDocumentLibraryListItems = $IncludeDocumentLibraryListItems
            ListMetadataOnly                = $ListMetadataOnly
            MaxSites                        = $MaxSites
            FirstSiteOnly                   = $FirstSiteOnly
        }

        if (-not [string]::IsNullOrWhiteSpace($TenantName)) {
            $generatorParams.TenantName = $TenantName
        }

        $generatorResult = Export-SharePointMetadataManifest @generatorParams
        $generatorResult | Out-Host

        if (-not [string]::IsNullOrWhiteSpace($generatorResult.Path)) {
            $manifestPath = [System.IO.Path]::GetFullPath($generatorResult.Path)
        }
    }
    else {
        Write-Host "Using cached SharePoint manifest: $manifestPath" -ForegroundColor Cyan
    }

    Import-SharePointManifestSet `
        -ManifestPath $manifestPath `
        -ManifestDir $manifestDirPath
}

function ConvertFrom-SharePointManifestSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$ManifestSet,

        [ValidateSet('All', 'DriveItems', 'ListItems')]
        [string[]]$ItemType = @('All'),

        [switch]$IncludeDriveFolders
    )

    $includeAll        = $ItemType -contains 'All'
    $includeDriveItems = $includeAll -or ($ItemType -contains 'DriveItems')
    $includeListItems  = $includeAll -or ($ItemType -contains 'ListItems')

    foreach ($manifest in @($ManifestSet.Manifests)) {
        foreach ($siteEntry in @($manifest.sites)) {
            $site = $siteEntry.metadata

            if ($includeDriveItems) {
                foreach ($driveEntry in @($siteEntry.drives)) {
                    $drive = $driveEntry.metadata

                    foreach ($item in @($driveEntry.items)) {
                        $isFolder = $null -ne $item.folder

                        if ($isFolder -and -not $IncludeDriveFolders) {
                            continue
                        }

                        $kind = if ($isFolder) {
                            'DriveFolder'
                        }
                        elseif ($null -ne $item.file) {
                            'DriveFile'
                        }
                        else {
                            'DriveItem'
                        }

                        [pscustomobject]@{
                            SourceKey          = 'sharepoint:driveItem:{0}' -f $item.id
                            SourceETag         = $item.eTag
                            Kind               = $kind
                            Name               = $item.name
                            WebUrl             = $item.webUrl
                            Size               = $item.size
                            CreatedDateTime    = $item.createdDateTime
                            LastModified       = $item.lastModifiedDateTime
                            IsContainer        = [bool]$isFolder
                            SiteId             = $site.id
                            SiteName           = $site.displayName
                            SiteWebUrl         = $site.webUrl
                            DriveId            = $drive.id
                            DriveName          = $drive.name
                            ListId             = $null
                            ListName           = $null
                            ItemId             = $item.id
                            ParentReference    = $item.parentReference
                            Raw                = $item
                        }
                    }
                }
            }

            if ($includeListItems) {
                foreach ($listEntry in @($siteEntry.lists)) {
                    $list = $listEntry.metadata

                    foreach ($item in @($listEntry.items)) {
                        [pscustomobject]@{
                            SourceKey       = 'sharepoint:listItem:{0}:{1}' -f $list.id, $item.id
                            SourceETag      = $item.eTag
                            Kind            = 'ListItem'
                            Name            = $item.fields.Title ?? $item.webUrl ?? $item.id
                            WebUrl          = $item.webUrl
                            Size            = $null
                            CreatedDateTime = $item.createdDateTime
                            LastModified    = $item.lastModifiedDateTime
                            IsContainer     = $false
                            SiteId          = $site.id
                            SiteName        = $site.displayName
                            SiteWebUrl      = $site.webUrl
                            DriveId         = $null
                            DriveName       = $null
                            ListId          = $list.id
                            ListName        = $list.displayName ?? $list.name
                            ItemId          = $item.id
                            Fields          = $item.fields
                            ContentType     = $item.contentType
                            Raw             = $item
                        }
                    }
                }
            }
        }
    }
}

function Export-SharePointWorkQueue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]]$InputObject,

        [Parameter(Mandatory)]
        [string]$Path
    )

    begin {
        $items = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($item in $InputObject) {
            $items.Add(($item | ConvertTo-Json -Depth 100 -Compress))
        }
    }

    end {
        $fullPath = [System.IO.Path]::GetFullPath($Path)
        $directory = Split-Path -Parent $fullPath

        if (-not (Test-Path -LiteralPath $directory)) {
            $null = New-Item -ItemType Directory -Path $directory -Force
        }

        [System.IO.File]::WriteAllLines(
            $fullPath,
            $items.ToArray(),
            [System.Text.UTF8Encoding]::new($false)
        )

        [pscustomobject]@{
            Path  = $fullPath
            Count = $items.Count
        }
    }
}

function Import-SharePointMigrationState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $state = @{}

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $state
    }

    foreach ($line in Get-Content -LiteralPath $Path) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $entry = $line | ConvertFrom-Json

        if (-not [string]::IsNullOrWhiteSpace($entry.sourceKey)) {
            $state[$entry.sourceKey] = $entry
        }
    }

    return $state
}

function Test-SharePointItemAlreadyMigrated {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Item,

        [Parameter(Mandatory)]
        [hashtable]$State,

        [switch]$IgnoreETag
    )

    if ([string]::IsNullOrWhiteSpace($Item.SourceKey)) {
        return $false
    }

    if (-not $State.ContainsKey($Item.SourceKey)) {
        return $false
    }

    $entry = $State[$Item.SourceKey]

    if ($entry.status -ne 'Completed') {
        return $false
    }

    if ($IgnoreETag) {
        return $true
    }

    if ([string]::IsNullOrWhiteSpace($Item.SourceETag)) {
        return $true
    }

    return ($entry.sourceETag -eq $Item.SourceETag)
}

function Write-SharePointMigrationStateEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [object]$Item,

        [Parameter(Mandatory)]
        [ValidateSet('Completed', 'Skipped', 'Failed')]
        [string]$Status,

        [string]$HuduType,

        [object]$HuduId,

        [string]$Message
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $directory = Split-Path -Parent $fullPath

    if (-not (Test-Path -LiteralPath $directory)) {
        $null = New-Item -ItemType Directory -Path $directory -Force
    }

    $entry = [ordered]@{
        sourceKey      = $Item.SourceKey
        sourceETag     = $Item.SourceETag
        status         = $Status
        huduType       = $HuduType
        huduId         = $HuduId
        message        = $Message
        completedAtUtc = [datetime]::UtcNow.ToString('o')
    }

    Add-Content `
        -LiteralPath $fullPath `
        -Value (($entry | ConvertTo-Json -Depth 20 -Compress)) `
        -Encoding utf8

    [pscustomobject]$entry
}
