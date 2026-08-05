<#
.SYNOPSIS
Dump SharePoint drive-item client/company metadata and resolve lookup IDs against site lists.

.DESCRIPTION
This is a diagnostic helper for figuring out which SharePoint metadata field/list backs
client attribution on files. It writes:
- one CSV row per file, including client/company-ish fields and lookup IDs
- one CSV row per lookup candidate found in matching SharePoint lists
- one CSV row per lookup column on the selected document libraries
#>

param(
    [hashtable]$Headers = $SharePointHeaders,

    [string]$SiteId,
    [string]$SiteUrl,

    [string[]]$DriveIds = @(),
    [string[]]$DriveNames = @(),

    [string]$OutputPath = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'logs') 'sharepoint-file-client-metadata.csv'),
    [string]$LookupOutputPath = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'logs') 'sharepoint-file-client-lookup-candidates.csv'),
    [string]$ColumnOutputPath = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'logs') 'sharepoint-file-client-lookup-columns.csv'),

    [string]$ListNamePattern = '(?i)client|compan|customer|account',
    [string[]]$LookupListIds = @(),
    [string[]]$LookupListNames = @(),
    [switch]$ResolveAllLists,
    [switch]$NoLookupResolution,
    [switch]$IncludeAllLookupListItems,
    [switch]$SkipFileScan,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxItems = 0
)

$ErrorActionPreference = 'Stop'

function Write-ClientMetadataLog {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [string]$Color = 'White'
    )

    if (Get-Command Set-PrintAndLog -ErrorAction SilentlyContinue) {
        try {
            Set-PrintAndLog -message $Message -Color $Color
            return
        } catch {}
    }

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message" -ForegroundColor $Color
}

function ConvertTo-ClientMetadataKey {
    param($Value)

    if ($null -eq $Value) { return "" }
    $text = ([string]$Value).Normalize([Text.NormalizationForm]::FormD).ToLowerInvariant()
    $text = $text -replace '\p{Mn}', ''
    try {
        $text = [System.Web.HttpUtility]::HtmlDecode($text)
    } catch {
        $text = [System.Net.WebUtility]::HtmlDecode($text)
    }
    $text = $text -replace '&', ' and '
    $text = $text -replace '[^a-z0-9]+', ' '
    return ($text -replace '\s+', ' ').Trim()
}

function ConvertFrom-ClientMetadataInternalFieldName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return $Name }

    return [regex]::Replace($Name, '_x(?<hex>[0-9a-fA-F]{4})_', {
        param($match)
        [string][char][int]("0x$($match.Groups['hex'].Value)")
    })
}

function ConvertTo-ClientMetadataText {
    param($Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [string] -or $Value -is [ValueType]) {
        return ([string]$Value).Trim()
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $parts = @(
            foreach ($entry in @($Value)) {
                ConvertTo-ClientMetadataText -Value $entry
            }
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        if ($parts.Count -gt 0) {
            return ($parts -join '; ')
        }
    }

    foreach ($propertyName in @('LookupValue', 'lookupValue', 'Title', 'title', 'Name', 'name', 'Value', 'value')) {
        $property = $Value.PSObject.Properties[$propertyName]
        if ($property -and $null -ne $property.Value) {
            $text = ConvertTo-ClientMetadataText -Value $property.Value
            if (-not [string]::IsNullOrWhiteSpace($text)) { return $text }
        }
    }

    $json = $Value | ConvertTo-Json -Compress -Depth 12
    if ($null -eq $json) { return '' }
    return ([string]$json).Trim()
}

function Select-ClientMetadataFirstValue {
    param(
        [Parameter(ValueFromRemainingArguments)]
        [object[]]$Values
    )

    foreach ($value in @($Values)) {
        if ($null -eq $value) { continue }
        $text = [string]$value
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            return $text
        }
    }

    return $null
}

function Get-ClientMetadataHeaders {
    if (Get-Command Update-SharePointAccessToken -ErrorAction SilentlyContinue) {
        return Update-SharePointAccessToken
    }

    if ($null -eq $script:Headers -or -not $script:Headers.ContainsKey('Authorization')) {
        throw "SharePoint Graph headers are not available. Run auth setup first, or pass -Headers."
    }

    return $script:Headers
}

function Invoke-ClientMetadataGraphRequest {
    param([Parameter(Mandatory)] [string]$Uri)

    Invoke-RestMethod -Method Get -Uri $Uri -Headers (Get-ClientMetadataHeaders) -ErrorAction Stop
}

function Invoke-ClientMetadataGraphCollection {
    param([Parameter(Mandatory)] [string]$Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-ClientMetadataGraphRequest -Uri $nextUri
        foreach ($item in @($response.value)) {
            $items.Add($item)
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return @($items)
}

function Resolve-ClientMetadataSite {
    param(
        [string]$GraphSiteId,
        [string]$SharePointSiteUrl
    )

    if (-not [string]::IsNullOrWhiteSpace($GraphSiteId)) {
        return Invoke-ClientMetadataGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$GraphSiteId"
    }

    if ([string]::IsNullOrWhiteSpace($SharePointSiteUrl)) {
        throw "Provide -SiteId or -SiteUrl."
    }

    $parsedUrl = [uri]$SharePointSiteUrl
    $hostname = $parsedUrl.Host
    $serverRelativePath = $parsedUrl.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($serverRelativePath)) { $serverRelativePath = '/' }

    Invoke-ClientMetadataGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/${hostname}:$serverRelativePath"
}

function Get-ClientMetadataDriveItems {
    param(
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] $Drive,
        [string]$FolderId = 'root',
        [string[]]$FolderPath = @()
    )

    $itemsUri = if ($FolderId -eq 'root') {
        "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($Drive.id)/root/children"
    } else {
        "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($Drive.id)/items/$FolderId/children"
    }
    $items = @(Invoke-ClientMetadataGraphCollection -Uri $itemsUri)

    foreach ($item in $items) {
        if ($item.folder) {
            $nextPath = @($FolderPath) + @([string]$item.name)
            foreach ($child in @(Get-ClientMetadataDriveItems -Site $Site -Drive $Drive -FolderId $item.id -FolderPath $nextPath)) {
                $child
            }
            continue
        }

        if (-not $item.file) { continue }

        [PSCustomObject]@{
            Item       = $item
            FolderPath = @($FolderPath)
            Drive      = $Drive
            Site       = $Site
        }
    }
}

function Get-ClientMetadataListItemFields {
    param(
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] $Drive,
        [Parameter(Mandatory)] $DriveItem
    )

    $uri = "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($Drive.id)/items/$($DriveItem.id)/listItem?`$expand=fields"
    $listItem = Invoke-ClientMetadataGraphRequest -Uri $uri
    return $listItem.fields
}

function Get-ClientMetadataFieldSummary {
    param($Fields)

    $interesting = [ordered]@{}
    $lookups = [ordered]@{}
    $lookupIds = [System.Collections.Generic.List[string]]::new()
    $textCandidates = [System.Collections.Generic.List[string]]::new()

    if (-not $Fields) {
        return [PSCustomObject]@{
            Interesting    = $interesting
            Lookups        = $lookups
            LookupIds      = @()
            TextCandidates = @()
        }
    }

    foreach ($property in @($Fields.PSObject.Properties)) {
        $decodedName = ConvertFrom-ClientMetadataInternalFieldName -Name $property.Name
        $nameKey = ConvertTo-ClientMetadataKey -Value $decodedName
        $isLookup = [string]$property.Name -match 'LookupId$'
        $isInteresting = $isLookup -or $nameKey -match '\b(client|company|customer|account|tenant|site)\b'

        if (-not $isInteresting) { continue }

        $valueText = ConvertTo-ClientMetadataText -Value $property.Value
        $interesting[$property.Name] = $valueText

        if ($isLookup) {
            $lookupId = [string]$property.Value
            if (-not [string]::IsNullOrWhiteSpace($lookupId)) {
                $lookups[$property.Name] = $lookupId
                if (-not $lookupIds.Contains($lookupId)) {
                    $lookupIds.Add($lookupId)
                }
            }
        } elseif (-not [string]::IsNullOrWhiteSpace($valueText)) {
            $textCandidates.Add("$decodedName=$valueText")
        }
    }

    [PSCustomObject]@{
        Interesting    = $interesting
        Lookups        = $lookups
        LookupIds      = @($lookupIds)
        TextCandidates = @($textCandidates)
    }
}

function Get-ClientMetadataPreferredTitle {
    param($Fields)

    if (-not $Fields) { return $null }
    foreach ($fieldName in @('Title', 'LinkTitle', 'Client Name', 'ClientName', 'Client_x0020_Name', 'Company', 'Name')) {
        $property = $Fields.PSObject.Properties[$fieldName]
        if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return ConvertTo-ClientMetadataText -Value $property.Value
        }
    }

    foreach ($property in @($Fields.PSObject.Properties)) {
        $decodedName = ConvertFrom-ClientMetadataInternalFieldName -Name $property.Name
        if ((ConvertTo-ClientMetadataKey -Value $decodedName) -match '\b(title|client|company|name)\b') {
            $value = ConvertTo-ClientMetadataText -Value $property.Value
            if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
        }
    }

    return $null
}

function Get-ClientMetadataSiteLists {
    param([Parameter(Mandatory)] $Site)

    try {
        return @(Invoke-ClientMetadataGraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$($Site.id)/lists")
    } catch {
        Write-ClientMetadataLog -Message "Failed to list site lists: $($_.Exception.Message)" -Color Yellow
        return @()
    }
}

function Get-ClientMetadataDriveList {
    param(
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] $Drive
    )

    try {
        return Invoke-ClientMetadataGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($Drive.id)/list"
    } catch {
        return $null
    }
}

function Get-ClientMetadataLookupColumnRows {
    param(
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] $Drive,
        [hashtable]$ListNameById = @{}
    )

    $driveLabel = [string](Select-ClientMetadataFirstValue $Drive.name $Drive.id)
    $driveList = Get-ClientMetadataDriveList -Site $Site -Drive $Drive
    if (-not $driveList) {
        Write-ClientMetadataLog -Message "Could not resolve backing list for drive '$driveLabel'; lookup-column dump skipped for that drive." -Color Yellow
        return @()
    }

    $columns = @(Invoke-ClientMetadataGraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$($Site.id)/lists/$($driveList.id)/columns")
    foreach ($column in $columns) {
        $lookupInfo = $column.lookup
        $decodedName = ConvertFrom-ClientMetadataInternalFieldName -Name ([string]$column.name)
        $nameKey = ConvertTo-ClientMetadataKey -Value (Select-ClientMetadataFirstValue $column.displayName $decodedName)
        $looksClientish = $nameKey -match '\b(client|company|customer|account|tenant|site)\b'
        if (-not $lookupInfo -and -not $looksClientish) { continue }

        $sourceListId = [string](Select-ClientMetadataFirstValue $lookupInfo.listId)
        $sourceListName = $null
        if (-not [string]::IsNullOrWhiteSpace($sourceListId) -and $ListNameById.ContainsKey($sourceListId)) {
            $sourceListName = $ListNameById[$sourceListId]
        }

        [PSCustomObject]@{
            DriveName          = $driveLabel
            DriveId            = $Drive.id
            DocumentLibraryId  = $driveList.id
            DocumentLibraryUrl = $driveList.webUrl
            ColumnName         = $column.name
            DecodedColumnName  = $decodedName
            DisplayName        = $column.displayName
            IsLookup           = [bool]$lookupInfo
            LookupListId       = $sourceListId
            LookupListName     = $sourceListName
            LookupColumnName   = [string](Select-ClientMetadataFirstValue $lookupInfo.columnName)
            ColumnJson         = ($column | ConvertTo-Json -Compress -Depth 12)
        }
    }
}

function Select-ClientMetadataTargetLists {
    param(
        [object[]]$Lists,
        [string[]]$Ids = @(),
        [string[]]$Names = @(),
        [string]$NamePattern,
        [switch]$All
    )

    if ($All) { return @($Lists) }

    if ($Ids.Count -gt 0) {
        $idSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($id in @($Ids)) {
            if (-not [string]::IsNullOrWhiteSpace($id)) {
                [void]$idSet.Add([string]$id)
            }
        }

        return @($Lists | Where-Object { $idSet.Contains([string]$_.id) })
    }

    if ($Names.Count -gt 0) {
        $nameKeys = @($Names | ForEach-Object { ConvertTo-ClientMetadataKey -Value $_ })
        return @($Lists | Where-Object {
            $displayNameKey = ConvertTo-ClientMetadataKey -Value $_.displayName
            $nameKey = ConvertTo-ClientMetadataKey -Value $_.name
            $nameKeys -contains $displayNameKey -or $nameKeys -contains $nameKey
        })
    }

    return @($Lists | Where-Object {
        [string](Select-ClientMetadataFirstValue $_.displayName $_.name) -match $NamePattern
    })
}

$resolvedOutputPath = if ([System.IO.Path]::IsPathRooted($OutputPath)) {
    [System.IO.Path]::GetFullPath($OutputPath)
} else {
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $OutputPath))
}

$resolvedLookupOutputPath = if ([System.IO.Path]::IsPathRooted($LookupOutputPath)) {
    [System.IO.Path]::GetFullPath($LookupOutputPath)
} else {
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $LookupOutputPath))
}

$resolvedColumnOutputPath = if ([System.IO.Path]::IsPathRooted($ColumnOutputPath)) {
    [System.IO.Path]::GetFullPath($ColumnOutputPath)
} else {
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $ColumnOutputPath))
}

foreach ($path in @($resolvedOutputPath, $resolvedLookupOutputPath, $resolvedColumnOutputPath)) {
    $directory = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        $null = New-Item -ItemType Directory -Path $directory -Force
    }
}

$site = Resolve-ClientMetadataSite -GraphSiteId $SiteId -SharePointSiteUrl $SiteUrl
$siteLabel = [string](Select-ClientMetadataFirstValue $site.displayName $site.name $site.id)
$siteLists = @(Get-ClientMetadataSiteLists -Site $site)
$listNameById = @{}
foreach ($list in @($siteLists)) {
    $listId = [string]$list.id
    if ([string]::IsNullOrWhiteSpace($listId)) { continue }
    $listNameById[$listId] = [string](Select-ClientMetadataFirstValue $list.displayName $list.name $list.id)
}

$drives = @(Invoke-ClientMetadataGraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives")
if ($DriveIds.Count -gt 0) {
    $driveIdSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($driveId in @($DriveIds)) { [void]$driveIdSet.Add([string]$driveId) }
    $drives = @($drives | Where-Object { $driveIdSet.Contains([string]$_.id) })
}
if ($DriveNames.Count -gt 0) {
    $driveNameKeys = @($DriveNames | ForEach-Object { ConvertTo-ClientMetadataKey -Value $_ })
    $drives = @($drives | Where-Object {
        $driveNameKey = ConvertTo-ClientMetadataKey -Value $_.name
        $driveNameKeys -contains $driveNameKey
    })
}

Write-ClientMetadataLog -Message "Dumping SharePoint file client metadata for '$siteLabel'. Drives=$($drives.Count); ResolveAllLists=$ResolveAllLists; NoLookupResolution=$NoLookupResolution; SkipFileScan=$SkipFileScan." -Color Cyan

$fileRows = [System.Collections.Generic.List[object]]::new()
$columnRows = [System.Collections.Generic.List[object]]::new()
$lookupIdSet = [System.Collections.Generic.HashSet[string]]::new()
$processed = 0

foreach ($drive in @($drives)) {
    $driveLabel = [string](Select-ClientMetadataFirstValue $drive.name $drive.id)
    Write-ClientMetadataLog -Message "Scanning drive '$driveLabel'." -Color DarkCyan

    foreach ($columnRow in @(Get-ClientMetadataLookupColumnRows -Site $site -Drive $drive -ListNameById $listNameById)) {
        $columnRows.Add($columnRow)
    }

    if ($SkipFileScan) { continue }

    foreach ($entry in @(Get-ClientMetadataDriveItems -Site $site -Drive $drive)) {
        if ($MaxItems -gt 0 -and $processed -ge $MaxItems) { break }
        $processed++

        $item = $entry.Item
        $fields = $null
        $errorMessage = $null
        try {
            $fields = Get-ClientMetadataListItemFields -Site $site -Drive $drive -DriveItem $item
        } catch {
            $errorMessage = $_.Exception.Message
        }

        $summary = Get-ClientMetadataFieldSummary -Fields $fields
        foreach ($lookupId in @($summary.LookupIds)) {
            [void]$lookupIdSet.Add([string]$lookupId)
        }

        $fileRows.Add([PSCustomObject]@{
            SiteName              = $siteLabel
            SiteId                = $site.id
            DriveName             = $driveLabel
            DriveId               = $drive.id
            SharePointItemId      = $item.id
            SharePointName        = $item.name
            ArticleTitle          = [System.IO.Path]::GetFileNameWithoutExtension([string]$item.name)
            SharePointUrl         = $item.webUrl
            RelativeFolderPath    = (@($entry.FolderPath) -join '\')
            ParentDrivePath       = $item.parentReference.path
            MetadataLookupIds     = (@($summary.LookupIds) -join '; ')
            MetadataLookupFields  = ($summary.Lookups | ConvertTo-Json -Compress -Depth 12)
            MetadataTextCandidates = (@($summary.TextCandidates) -join '; ')
            InterestingFieldsJson = ($summary.Interesting | ConvertTo-Json -Compress -Depth 12)
            MetadataReadError     = $errorMessage
        })
    }
}

$lookupCandidateRows = [System.Collections.Generic.List[object]]::new()
$lookupCandidateById = @{}

if (-not $NoLookupResolution -and ($lookupIdSet.Count -gt 0 -or $IncludeAllLookupListItems)) {
    $lists = $siteLists
    if ($lists.Count -lt 1) {
        $lists = @(Invoke-ClientMetadataGraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists")
    }
    $targetLists = @(Select-ClientMetadataTargetLists `
        -Lists $lists `
        -Ids $LookupListIds `
        -Names $LookupListNames `
        -NamePattern $ListNamePattern `
        -All:$ResolveAllLists)

    Write-ClientMetadataLog -Message "Resolving $($lookupIdSet.Count) unique lookup id(s) across $($targetLists.Count) list(s). IncludeAllLookupListItems=$IncludeAllLookupListItems." -Color DarkCyan

    foreach ($list in @($targetLists)) {
        $listLabel = [string](Select-ClientMetadataFirstValue $list.displayName $list.name $list.id)
        try {
            $items = @(Invoke-ClientMetadataGraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/items?`$expand=fields&`$top=999")
            foreach ($listItem in $items) {
                $itemId = [string](Select-ClientMetadataFirstValue $listItem.id $listItem.fields.id $listItem.fields.ID)
                if ([string]::IsNullOrWhiteSpace($itemId)) { continue }
                if (-not $IncludeAllLookupListItems -and -not $lookupIdSet.Contains($itemId)) { continue }

                $title = Get-ClientMetadataPreferredTitle -Fields $listItem.fields
                $row = [PSCustomObject]@{
                    LookupId       = $itemId
                    IsUsedByScannedFiles = $lookupIdSet.Contains($itemId)
                    ListName       = $listLabel
                    ListId         = $list.id
                    CandidateTitle = $title
                    CandidateUrl   = $listItem.webUrl
                    CandidateFieldsJson = ($listItem.fields | ConvertTo-Json -Compress -Depth 12)
                }

                $lookupCandidateRows.Add($row)
                if (-not $lookupCandidateById.ContainsKey($itemId)) {
                    $lookupCandidateById[$itemId] = [System.Collections.Generic.List[object]]::new()
                }
                $lookupCandidateById[$itemId].Add($row)
            }
        } catch {
            Write-ClientMetadataLog -Message "Failed to resolve lookup ids from list '$listLabel': $($_.Exception.Message)" -Color Yellow
        }
    }
}

$enrichedRows = foreach ($row in $fileRows) {
    $candidateParts = [System.Collections.Generic.List[string]]::new()
    foreach ($lookupId in @([string]$row.MetadataLookupIds -split ';\s*' | Where-Object { $_ })) {
        if (-not $lookupCandidateById.ContainsKey($lookupId)) { continue }
        foreach ($candidate in @($lookupCandidateById[$lookupId])) {
            $candidateParts.Add("$lookupId => $($candidate.ListName): $($candidate.CandidateTitle)")
        }
    }

    $row | Add-Member -NotePropertyName LookupCandidates -NotePropertyValue (@($candidateParts) -join '; ') -Force
    $row
}

$enrichedRows | Export-Csv -LiteralPath $resolvedOutputPath -NoTypeInformation -Encoding UTF8
$lookupCandidateRows | Export-Csv -LiteralPath $resolvedLookupOutputPath -NoTypeInformation -Encoding UTF8
$columnRows | Export-Csv -LiteralPath $resolvedColumnOutputPath -NoTypeInformation -Encoding UTF8

$summary = [PSCustomObject]@{
    SiteName          = $siteLabel
    SiteId            = $site.id
    Drives            = $drives.Count
    ProcessedFiles    = $processed
    UniqueLookupIds   = $lookupIdSet.Count
    LookupCandidates  = $lookupCandidateRows.Count
    LookupColumns     = $columnRows.Count
    OutputPath        = $resolvedOutputPath
    LookupOutputPath  = $resolvedLookupOutputPath
    ColumnOutputPath  = $resolvedColumnOutputPath
}

Write-ClientMetadataLog -Message "Client metadata dump complete: files=$processed, uniqueLookupIds=$($lookupIdSet.Count), lookupCandidates=$($lookupCandidateRows.Count). Output: $resolvedOutputPath" -Color Cyan

$summary
