param(
    [hashtable]$Headers = $SharePointHeaders,

    [string]$OutputDirectory = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'out') 'sharepoint-list-csv'),

    # Optional Graph site IDs. When omitted, the job enumerates sites with /sites?search=*.
    [string[]]$SiteIds = @(),

    # Optional displayName/name filter applied after site enumeration.
    [string[]]$SiteNames = @(),

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxSites = 0,

    [switch]$IncludeHidden,

    [switch]$IncludeSystem,

    # Document libraries are skipped by default because file metadata is usually better handled as drive data.
    [switch]$IncludeDocumentLibraries,

    [switch]$IncludeEmptyLists,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Write-SharePointListCsvLog {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [string]$Color = 'White'
    )

    if (Get-Command Set-PrintAndLog -ErrorAction SilentlyContinue) {
        Set-PrintAndLog -message $Message -Color $Color
    } else {
        Write-Host $Message -ForegroundColor $Color
    }
}

function Get-SharePointListCsvHeaders {
    if (Get-Command Update-SharePointAccessToken -ErrorAction SilentlyContinue) {
        return Update-SharePointAccessToken
    }

    if ($null -eq $script:Headers -or -not $script:Headers.ContainsKey('Authorization')) {
        throw "SharePoint Graph headers are not available. Run your environment/auth setup first, or pass -Headers."
    }

    return $script:Headers
}

function Invoke-SharePointListCsvRequest {
    param(
        [Parameter(Mandatory)] [string]$Uri,
        [ValidateRange(0, 20)] [int]$MaxRetries = 6
    )

    $attempt = 0

    while ($true) {
        try {
            return Invoke-RestMethod `
                -Method Get `
                -Uri $Uri `
                -Headers (Get-SharePointListCsvHeaders) `
                -ErrorAction Stop
        }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

            $isTransient = $statusCode -in @(429, 502, 503, 504)
            if (-not $isTransient -or $attempt -ge $MaxRetries) {
                throw
            }

            $delaySeconds = [math]::Min(60, [math]::Pow(2, $attempt + 1))
            Write-SharePointListCsvLog -Message "Request returned HTTP $statusCode. Retrying in $delaySeconds second(s): $Uri" -Color Yellow
            Start-Sleep -Seconds $delaySeconds
            $attempt++
        }
    }
}

function Invoke-SharePointListCsvCollection {
    param(
        [Parameter(Mandatory)] [string]$Uri,
        [string]$StatusLabel
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    $pageCount = 0

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-SharePointListCsvRequest -Uri $nextUri

        if ($null -ne $response.value) {
            foreach ($item in @($response.value)) {
                $items.Add($item)
            }
        } else {
            $items.Add($response)
        }

        $pageCount++
        if (-not [string]::IsNullOrWhiteSpace($StatusLabel) -and ($pageCount -eq 1 -or $pageCount % 10 -eq 0)) {
            Write-SharePointListCsvLog -Message "$StatusLabel`: $($items.Count) item(s) fetched..." -Color DarkGray
        }

        $nextUri = $response.'@odata.nextLink'
    }

    return @($items)
}

function Get-SharePointListCsvSafeName {
    param(
        [string]$Name,
        [string]$Fallback = 'unnamed'
    )

    $value = if ([string]::IsNullOrWhiteSpace($Name)) { $Fallback } else { $Name }
    $safe = (($value -replace '[\\/:*?"<>|]', '_') -replace '\s{2,}', ' ').Trim()

    if ([string]::IsNullOrWhiteSpace($safe)) {
        return $Fallback
    }

    return $safe
}

function ConvertTo-SharePointListCsvValue {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string] -or $Value -is [ValueType]) {
        return $Value
    }

    return ($Value | ConvertTo-Json -Compress -Depth 20)
}

function ConvertTo-SharePointListCsvRow {
    param(
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] $List,
        [Parameter(Mandatory)] $Item,
        [string[]]$FieldNames
    )

    $row = [ordered]@{
        SiteName = $Site.displayName
        SiteId   = $Site.id
        SiteUrl  = $Site.webUrl
        ListName = $List.displayName
        ListId   = $List.id
        ListUrl  = $List.webUrl
        ItemId   = $Item.id
        ItemUrl  = $Item.webUrl
    }

    $fields = $Item.fields
    foreach ($fieldName in $FieldNames) {
        if ($row.Contains($fieldName)) {
            continue
        }

        $value = $null
        if ($fields -and ($fields.PSObject.Properties.Name -contains $fieldName)) {
            $value = $fields.PSObject.Properties[$fieldName].Value
        }

        $row[$fieldName] = ConvertTo-SharePointListCsvValue -Value $value
    }

    return [pscustomobject]$row
}

function New-SharePointListCsvEmptyRow {
    param(
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] $List,
        [string[]]$FieldNames
    )

    $row = [ordered]@{
        SiteName = $Site.displayName
        SiteId   = $Site.id
        SiteUrl  = $Site.webUrl
        ListName = $List.displayName
        ListId   = $List.id
        ListUrl  = $List.webUrl
        ItemId   = $null
        ItemUrl  = $null
    }

    foreach ($fieldName in $FieldNames) {
        if (-not $row.Contains($fieldName)) {
            $row[$fieldName] = $null
        }
    }

    return [pscustomobject]$row
}

function Get-SharePointListCsvFieldNames {
    param(
        [object[]]$Columns,
        [object[]]$Items
    )

    $seen = [ordered]@{}

    foreach ($column in @($Columns)) {
        $name = [string]$column.name
        if (-not [string]::IsNullOrWhiteSpace($name) -and -not $seen.Contains($name)) {
            $seen[$name] = $true
        }
    }

    foreach ($item in @($Items)) {
        if (-not $item.fields) {
            continue
        }

        foreach ($property in @($item.fields.PSObject.Properties)) {
            if (-not $seen.Contains($property.Name)) {
                $seen[$property.Name] = $true
            }
        }
    }

    return @($seen.Keys)
}

function Test-SharePointListCsvNameFilter {
    param(
        [Parameter(Mandatory)] $Site,
        [string[]]$Names
    )

    if ($null -eq $Names -or $Names.Count -lt 1) {
        return $true
    }

    $siteNamesToCheck = @(
        [string]$Site.displayName
        [string]$Site.name
        [string]$Site.webUrl
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($configuredName in $Names) {
        foreach ($siteName in $siteNamesToCheck) {
            if ($siteName -like $configuredName -or $siteName -eq $configuredName) {
                return $true
            }
        }
    }

    return $false
}

function Get-SharePointListCsvUniquePath {
    param(
        [Parameter(Mandatory)] [string]$Directory,
        [Parameter(Mandatory)] [string]$BaseName,
        [string]$Extension = '',
        [Parameter(Mandatory)] [hashtable]$UsedPaths
    )

    $candidate = Join-Path $Directory "$BaseName$Extension"
    $counter = 2

    while ($UsedPaths.ContainsKey($candidate)) {
        $candidate = Join-Path $Directory "$BaseName-$counter$Extension"
        $counter++
    }

    $UsedPaths[$candidate] = $true
    return $candidate
}

$resolvedOutputDirectory = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    [System.IO.Path]::GetFullPath($OutputDirectory)
} else {
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $OutputDirectory))
}

if ((Test-Path -LiteralPath $resolvedOutputDirectory) -and -not $Force) {
    Write-SharePointListCsvLog -Message "Output directory already exists; files may be overwritten list-by-list: $resolvedOutputDirectory" -Color Yellow
}

$null = New-Item -ItemType Directory -Path $resolvedOutputDirectory -Force

$sites = @()
if ($SiteIds.Count -gt 0) {
    foreach ($siteId in $SiteIds) {
        if ([string]::IsNullOrWhiteSpace($siteId)) {
            continue
        }

        $sites += Invoke-SharePointListCsvRequest -Uri "https://graph.microsoft.com/v1.0/sites/$siteId"
    }
} else {
    $sites = @(
        Invoke-SharePointListCsvCollection `
            -Uri "https://graph.microsoft.com/v1.0/sites?search=*" `
            -StatusLabel "SharePoint sites"
    )
}

$sites = @($sites | Where-Object { Test-SharePointListCsvNameFilter -Site $_ -Names $SiteNames })
if ($MaxSites -gt 0) {
    $sites = @($sites | Select-Object -First $MaxSites)
}

Write-SharePointListCsvLog -Message "Exporting SharePoint lists from $($sites.Count) site(s) to $resolvedOutputDirectory" -Color Cyan

$indexRows = [System.Collections.Generic.List[object]]::new()
$totalLists = 0
$totalItems = 0
$usedSiteDirectories = @{}
$usedCsvPaths = @{}

$siteSelect = @(
    'id'
    'name'
    'displayName'
    'webUrl'
) -join ','

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

$siteNumber = 0
foreach ($site in $sites) {
    $siteNumber++
    $siteId = [string]$site.id
    if ([string]::IsNullOrWhiteSpace($siteId)) {
        continue
    }

    if ([string]::IsNullOrWhiteSpace([string]$site.webUrl) -or [string]::IsNullOrWhiteSpace([string]$site.displayName)) {
        $site = Invoke-SharePointListCsvRequest -Uri "https://graph.microsoft.com/v1.0/sites/$siteId?`$select=$siteSelect"
    }

    $siteLabel = [string]($site.displayName ?? $site.name ?? $site.id)
    $siteDirectoryBaseName = Get-SharePointListCsvSafeName -Name $siteLabel -Fallback $siteId
    $siteDirectory = Get-SharePointListCsvUniquePath `
        -Directory $resolvedOutputDirectory `
        -BaseName $siteDirectoryBaseName `
        -Extension '' `
        -UsedPaths $usedSiteDirectories
    $null = New-Item -ItemType Directory -Path $siteDirectory -Force

    Write-SharePointListCsvLog -Message "Site $siteNumber/$($sites.Count): $siteLabel" -Color Cyan

    $listsUri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists?`$select=$listSelect"
    $lists = @(Invoke-SharePointListCsvCollection -Uri $listsUri -StatusLabel "Lists for $siteLabel")
    $listNumber = 0

    foreach ($list in $lists) {
        $listNumber++
        $listLabel = [string]($list.displayName ?? $list.name ?? $list.id)
        $isHidden = [bool]$list.list.hidden
        $isSystem = $null -ne $list.system
        $isDocumentLibrary = [string]$list.list.template -eq 'documentLibrary'

        if ($isHidden -and -not $IncludeHidden) {
            continue
        }

        if ($isSystem -and -not $IncludeSystem) {
            continue
        }

        if ($isDocumentLibrary -and -not $IncludeDocumentLibraries) {
            continue
        }

        Write-SharePointListCsvLog -Message "  List $listNumber/$($lists.Count): $listLabel" -Color DarkCyan

        $columnsUri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$($list.id)/columns"
        $itemsUri = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$($list.id)/items?`$expand=fields"

        $columns = @(Invoke-SharePointListCsvCollection -Uri $columnsUri)
        $items = @(Invoke-SharePointListCsvCollection -Uri $itemsUri -StatusLabel "Items for $siteLabel / $listLabel")

        if ($items.Count -lt 1 -and -not $IncludeEmptyLists) {
            $indexRows.Add([pscustomobject]@{
                SiteName     = $siteLabel
                SiteId       = $siteId
                SiteUrl      = $site.webUrl
                ListName     = $listLabel
                ListId       = $list.id
                ListUrl      = $list.webUrl
                Template     = $list.list.template
                Hidden       = $isHidden
                System       = $isSystem
                ItemCount    = 0
                CsvPath      = $null
                Exported     = $false
                SkipReason   = 'List was empty.'
            })
            continue
        }

        $fieldNames = @(Get-SharePointListCsvFieldNames -Columns $columns -Items $items)
        $rows = @(
            foreach ($item in $items) {
                ConvertTo-SharePointListCsvRow -Site $site -List $list -Item $item -FieldNames $fieldNames
            }
        )

        $listFileBaseName = Get-SharePointListCsvSafeName -Name $listLabel -Fallback $list.id
        $csvPath = Get-SharePointListCsvUniquePath `
            -Directory $siteDirectory `
            -BaseName $listFileBaseName `
            -Extension '.csv' `
            -UsedPaths $usedCsvPaths

        if ($rows.Count -gt 0) {
            $rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        } else {
            New-SharePointListCsvEmptyRow -Site $site -List $list -FieldNames $fieldNames |
                Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        }

        $totalLists++
        $totalItems += $items.Count

        $indexRows.Add([pscustomobject]@{
            SiteName     = $siteLabel
            SiteId       = $siteId
            SiteUrl      = $site.webUrl
            ListName     = $listLabel
            ListId       = $list.id
            ListUrl      = $list.webUrl
            Template     = $list.list.template
            Hidden       = $isHidden
            System       = $isSystem
            ItemCount    = $items.Count
            CsvPath      = $csvPath
            Exported     = $true
            SkipReason   = $null
        })
    }
}

$indexPath = Join-Path $resolvedOutputDirectory '_index.csv'
$indexRows | Export-Csv -LiteralPath $indexPath -NoTypeInformation -Encoding UTF8

$summary = [pscustomobject]@{
    OutputDirectory = $resolvedOutputDirectory
    IndexPath       = $indexPath
    Sites           = $sites.Count
    ExportedLists   = $totalLists
    ExportedItems   = $totalItems
    IndexedLists    = $indexRows.Count
}

Write-SharePointListCsvLog -Message "SharePoint list CSV export complete: $totalItems item(s) from $totalLists list(s)." -Color Green
Write-SharePointListCsvLog -Message "Index: $indexPath" -Color DarkMagenta

$summary
