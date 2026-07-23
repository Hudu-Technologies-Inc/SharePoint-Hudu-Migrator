##### Optional maintenance job, import structured-list JSON bundles as Hudu assets

$structuredAssetImportDryRun = [bool]($HuduStructuredAssetImportDryRun ?? $true)
$structuredAssetImportSkipExisting = [bool]($HuduStructuredAssetImportSkipExisting ?? $true)
$structuredAssetImportResolveCompanyByName = [bool]($HuduStructuredAssetImportResolveCompanyByName ?? $true)
$structuredAssetImportFallbackCompanyId = [int]($HuduStructuredAssetImportFallbackCompanyId ?? 0)
$structuredAssetImportMaxItems = [int]($HuduStructuredAssetImportMaxItems ?? 0)
$structuredAssetImportListNames = @(
    if ($null -ne $HuduStructuredAssetImportListNames) {
        @($HuduStructuredAssetImportListNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    } else {
        @()
    }
)

$structuredAssetImportRoot = if (-not [string]::IsNullOrWhiteSpace([string]$workdir)) {
    [string]$workdir
} elseif (-not [string]::IsNullOrWhiteSpace([string]$PSScriptRoot)) {
    (Split-Path -Parent $PSScriptRoot)
} else {
    (Get-Location).Path
}

foreach ($helperName in @('attribution.ps1', 'structuredlists.ps1')) {
    $helperPath = Join-Path (Join-Path $structuredAssetImportRoot 'helpers') $helperName
    if (Test-Path -LiteralPath $helperPath -PathType Leaf) {
        . $helperPath
    }
}

function Write-HuduStructuredAssetImportLog {
    param (
        [Parameter(Mandatory)] [string]$Message,
        [string]$Color = 'White'
    )

    if (Get-Command Set-PrintAndLog -ErrorAction SilentlyContinue) {
        Set-PrintAndLog -message $Message -Color $Color
    } else {
        Write-Host $Message -ForegroundColor $Color
    }
}

function ConvertTo-HuduStructuredAssetImportKey {
    param ($Value)

    if (Get-Command ConvertTo-AttributionNormalizedText -ErrorAction SilentlyContinue) {
        return (ConvertTo-AttributionNormalizedText $Value)
    }

    if ($null -eq $Value) { return "" }
    return (([string]$Value).ToLowerInvariant() -replace '[^a-z0-9]+', ' ').Trim()
}

function Get-HuduStructuredAssetImportBaseListName {
    param ($Name)

    if (Get-Command Get-SharePointStructuredListBaseName -ErrorAction SilentlyContinue) {
        return (Get-SharePointStructuredListBaseName $Name)
    }

    if ($null -eq $Name) { return "" }
    return ([string]$Name -replace '_\[[^\]]+\]$', '' -replace '\[[^\]]+\]$', '').Trim()
}

function ConvertFrom-HuduStructuredAssetImportFieldName {
    param ([string]$Name)

    if (Get-Command ConvertFrom-SharePointInternalFieldName -ErrorAction SilentlyContinue) {
        return (ConvertFrom-SharePointInternalFieldName $Name)
    }

    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    return ([regex]::Replace($Name, '_x(?<hex>[0-9a-fA-F]{4})_', {
        param($Match)
        [string][char][Convert]::ToInt32($Match.Groups['hex'].Value, 16)
    }))
}

function Resolve-HuduStructuredAssetImportPath {
    param ([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if ([System.IO.Path]::IsPathRooted($Path)) { return [System.IO.Path]::GetFullPath($Path) }
    return [System.IO.Path]::GetFullPath((Join-Path $structuredAssetImportRoot $Path))
}

$structuredAssetImportJsonPath = Resolve-HuduStructuredAssetImportPath (
    $HuduStructuredAssetImportJsonPath ??
    $RunSummary.OutputJsonFiles.StructuredListJsonDir ??
    (Join-Path 'logs' 'structured-list-json')
)
$structuredAssetImportReportPath = Resolve-HuduStructuredAssetImportPath (
    $HuduStructuredAssetImportReportPath ??
    (Join-Path (Join-Path 'logs' 'structured-list-assets') 'structured-list-asset-import.csv')
)
$structuredAssetImportMapsPath = Resolve-HuduStructuredAssetImportPath (
    $HuduStructuredAssetImportMapsPath ??
    (Join-Path $structuredAssetImportRoot 'structured-asset-maps.ps1')
)

if ($structuredAssetImportMapsPath -and (Test-Path -LiteralPath $structuredAssetImportMapsPath -PathType Leaf)) {
    . $structuredAssetImportMapsPath
}

if ($null -eq $HuduStructuredAssetImportMaps -or @($HuduStructuredAssetImportMaps.Keys).Count -lt 1) {
    throw "No structured asset import maps are configured. Set `$HuduStructuredAssetImportMaps or create $structuredAssetImportMapsPath."
}

function Get-HuduStructuredAssetImportBundlePaths {
    param ([Parameter(Mandatory)] [string]$Path)

    if (Test-Path -LiteralPath $Path -PathType Container) {
        return @(
            Get-ChildItem -LiteralPath $Path -Recurse -Filter '*.json' -File |
                Sort-Object FullName |
                ForEach-Object { $_.FullName }
        )
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Structured asset import path was not found: $Path"
    }

    if ([System.IO.Path]::GetExtension($Path) -ieq '.csv') {
        $baseDir = Split-Path -Parent $Path
        return @(
            Import-Csv -LiteralPath $Path |
                ForEach-Object { $_.OutputPath } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique |
                ForEach-Object {
                    if ([System.IO.Path]::IsPathRooted($_)) {
                        [System.IO.Path]::GetFullPath($_)
                    } else {
                        [System.IO.Path]::GetFullPath((Join-Path $baseDir $_))
                    }
                }
        )
    }

    return @([System.IO.Path]::GetFullPath($Path))
}

function Import-HuduStructuredAssetImportBundles {
    param ([Parameter(Mandatory)] [string[]]$Path)

    foreach ($bundlePath in @($Path)) {
        $json = Get-Content -LiteralPath $bundlePath -Raw | ConvertFrom-Json
        foreach ($bundle in @($json)) {
            if (-not $bundle.Items) { continue }
            $bundle | Add-Member -MemberType NoteProperty -Name SourcePath -Value $bundlePath -Force
            $bundle
        }
    }
}

function Resolve-HuduStructuredAssetImportMap {
    param (
        [Parameter(Mandatory)] [hashtable]$Maps,
        [Parameter(Mandatory)] [string]$ListName
    )

    $keysToTry = @(
        $ListName,
        (Get-HuduStructuredAssetImportBaseListName $ListName)
    ) |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

    foreach ($key in $keysToTry) {
        if ($Maps.ContainsKey($key)) { return $Maps[$key] }
    }

    $normalizedKeysToTry = @($keysToTry | ForEach-Object { ConvertTo-HuduStructuredAssetImportKey $_ })
    foreach ($mapKey in @($Maps.Keys)) {
        if ($normalizedKeysToTry -contains (ConvertTo-HuduStructuredAssetImportKey $mapKey)) {
            return $Maps[$mapKey]
        }
    }

    return $null
}

function Get-HuduStructuredAssetImportFieldValue {
    param (
        $Item,
        $Fields,
        [string]$FieldName
    )

    if ([string]::IsNullOrWhiteSpace($FieldName)) { return $null }

    foreach ($source in @($Item, $Fields)) {
        if (-not $source) { continue }
        if ($source.PSObject.Properties[$FieldName]) {
            return $source.PSObject.Properties[$FieldName].Value
        }
    }

    $targetKey = ConvertTo-HuduStructuredAssetImportKey $FieldName
    foreach ($source in @($Item, $Fields)) {
        if (-not $source) { continue }
        foreach ($property in @($source.PSObject.Properties)) {
            $propertyKeys = @(
                (ConvertTo-HuduStructuredAssetImportKey $property.Name),
                (ConvertTo-HuduStructuredAssetImportKey (ConvertFrom-HuduStructuredAssetImportFieldName $property.Name))
            )
            if ($propertyKeys -contains $targetKey) {
                return $property.Value
            }
        }
    }

    return $null
}

function ConvertTo-HuduStructuredAssetImportValue {
    param ($Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [array]) {
        $parts = @($Value | ForEach-Object { ConvertTo-HuduStructuredAssetImportValue $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($parts.Count -lt 1) { return $null }
        return ($parts -join '; ')
    }

    if ($Value -is [string] -or $Value -is [ValueType]) {
        if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
        return $Value
    }

    foreach ($propertyName in @('LookupValue', 'Value', 'Title', 'Name', 'Email', 'DisplayName')) {
        if ($Value.PSObject.Properties[$propertyName] -and -not [string]::IsNullOrWhiteSpace([string]$Value.$propertyName)) {
            return [string]$Value.$propertyName
        }
    }

    $json = $Value | ConvertTo-Json -Depth 12 -Compress
    if ([string]::IsNullOrWhiteSpace($json) -or $json -eq '{}') { return $null }
    return $json
}

function ConvertTo-HuduStructuredAssetImportFields {
    param (
        $Item,
        [System.Collections.IDictionary]$FieldMap,
        [switch]$AutoMapUnmappedFields,
        [string[]]$ExcludeFields = @()
    )

    $huduFields = [System.Collections.Generic.List[object]]::new()
    $mappedSourceKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($sourceFieldName in @($FieldMap.Keys)) {
        $huduLabel = [string]$FieldMap[$sourceFieldName]
        if ([string]::IsNullOrWhiteSpace($huduLabel)) { continue }

        $value = ConvertTo-HuduStructuredAssetImportValue (Get-HuduStructuredAssetImportFieldValue -Item $Item -Fields $Item.Fields -FieldName $sourceFieldName)
        [void]$mappedSourceKeys.Add((ConvertTo-HuduStructuredAssetImportKey $sourceFieldName))
        [void]$mappedSourceKeys.Add((ConvertTo-HuduStructuredAssetImportKey (ConvertFrom-HuduStructuredAssetImportFieldName $sourceFieldName)))

        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
            $huduFields.Add(@{ $huduLabel = $value })
        }
    }

    if ($AutoMapUnmappedFields -and $Item.Fields) {
        $excludedKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($excludedField in @($ExcludeFields + @('@odata.etag', 'ContentType', 'Attachments', 'Edit', 'LinkTitleNoMenu'))) {
            [void]$excludedKeys.Add((ConvertTo-HuduStructuredAssetImportKey $excludedField))
        }

        foreach ($property in @($Item.Fields.PSObject.Properties)) {
            if ($property.Name -like '@odata*') { continue }
            $decodedName = ConvertFrom-HuduStructuredAssetImportFieldName $property.Name
            $propertyKeys = @(
                (ConvertTo-HuduStructuredAssetImportKey $property.Name),
                (ConvertTo-HuduStructuredAssetImportKey $decodedName)
            )
            if (($propertyKeys | Where-Object { $mappedSourceKeys.Contains($_) -or $excludedKeys.Contains($_) }).Count -gt 0) {
                continue
            }

            $value = ConvertTo-HuduStructuredAssetImportValue $property.Value
            if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { continue }
            $huduFields.Add(@{ $decodedName = $value })
        }
    }

    return @($huduFields)
}

function Get-HuduStructuredAssetImportName {
    param (
        $Item,
        $Map
    )

    $nameFields = @(
        $Map.NameFields ??
        $Map.NameField ??
        $Map.AssetNameFields ??
        $Map.AssetNameField ??
        @('Title', 'LinkTitle', 'Name', 'Client Name')
    )

    foreach ($fieldName in @($nameFields)) {
        $value = ConvertTo-HuduStructuredAssetImportValue (Get-HuduStructuredAssetImportFieldValue -Item $Item -Fields $Item.Fields -FieldName $fieldName)
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            return [string]$value
        }
    }

    $id = $Item.SharePointItemId ?? $Item.id
    if ($id) { return "SharePoint item $id" }
    return "Imported SharePoint item"
}

function Resolve-HuduStructuredAssetImportLayoutId {
    param (
        $Map,
        [hashtable]$LayoutNameCache
    )

    $layoutId = $Map.AssetLayoutId ?? $Map.LayoutId
    if ($layoutId) { return [int]$layoutId }

    $layoutName = [string]($Map.AssetLayoutName ?? $Map.LayoutName)
    if ([string]::IsNullOrWhiteSpace($layoutName)) { return $null }

    $key = ConvertTo-HuduStructuredAssetImportKey $layoutName
    if ($LayoutNameCache.ContainsKey($key)) { return $LayoutNameCache[$key] }

    $layout = @(Get-HuduAssetLayouts -Name $layoutName) |
        Where-Object { (ConvertTo-HuduStructuredAssetImportKey $_.name) -eq $key } |
        Select-Object -First 1

    if (-not $layout) {
        $layout = Get-HuduAssetLayouts -Name $layoutName | Select-Object -First 1
    }

    $LayoutNameCache[$key] = ($layout.id ?? $null)
    return $LayoutNameCache[$key]
}

function Resolve-HuduStructuredAssetImportCompanyId {
    param (
        $Bundle,
        [switch]$ResolveByName
    )

    if ($Bundle.CompanyId) { return [int]$Bundle.CompanyId }
    if (-not $ResolveByName -or [string]::IsNullOrWhiteSpace([string]$Bundle.CompanyName) -or $Bundle.CompanyName -eq 'Unattributed') {
        return $null
    }

    $company = Get-HuduCompanies -Name $Bundle.CompanyName | Select-Object -First 1
    return ($company.id ?? $company.Id)
}

function Get-HuduStructuredAssetImportPrimaryValues {
    param (
        $Item,
        $Map
    )

    $primaryValues = @{}
    $primaryFieldMap = $Map.PrimaryFieldMap ?? $Map.PrimaryFields
    if (-not $primaryFieldMap) { return $primaryValues }

    foreach ($primaryName in @('PrimarySerial', 'PrimaryMail', 'PrimaryModel', 'PrimaryManufacturer')) {
        $sourceFieldName = $primaryFieldMap[$primaryName]
        $value = ConvertTo-HuduStructuredAssetImportValue (Get-HuduStructuredAssetImportFieldValue -Item $Item -Fields $Item.Fields -FieldName $sourceFieldName)
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            $primaryValues[$primaryName] = [string]$value
        }
    }

    return $primaryValues
}

function Test-HuduStructuredAssetImportListFilter {
    param ([string]$ListName)

    if ($structuredAssetImportListNames.Count -lt 1) { return $true }

    $listKey = ConvertTo-HuduStructuredAssetImportKey $ListName
    $baseKey = ConvertTo-HuduStructuredAssetImportKey (Get-HuduStructuredAssetImportBaseListName $ListName)
    foreach ($filterName in @($structuredAssetImportListNames)) {
        $filterKey = ConvertTo-HuduStructuredAssetImportKey $filterName
        if ($filterKey -eq $listKey -or $filterKey -eq $baseKey) { return $true }
    }

    return $false
}

function Get-HuduStructuredAssetImportExistingAssetIndex {
    param (
        [Parameter(Mandatory)] [int]$CompanyId,
        [Parameter(Mandatory)] [int]$AssetLayoutId,
        [Parameter(Mandatory)] [hashtable]$Cache
    )

    $cacheKey = "$CompanyId|$AssetLayoutId"
    if ($Cache.ContainsKey($cacheKey)) { return $Cache[$cacheKey] }

    $index = @{}
    foreach ($asset in @(Get-HuduAssets -CompanyId $CompanyId -AssetLayoutId $AssetLayoutId)) {
        $assetName = [string]($asset.name ?? $asset.Name)
        if ([string]::IsNullOrWhiteSpace($assetName)) { continue }
        $index[(ConvertTo-HuduStructuredAssetImportKey $assetName)] = $asset
    }

    $Cache[$cacheKey] = $index
    return $index
}

$bundlePaths = @(Get-HuduStructuredAssetImportBundlePaths -Path $structuredAssetImportJsonPath)
if ($bundlePaths.Count -lt 1) {
    throw "No structured-list JSON bundle files were found at $structuredAssetImportJsonPath."
}

$reportDirectory = Split-Path -Parent $structuredAssetImportReportPath
if (-not (Test-Path -LiteralPath $reportDirectory -PathType Container)) {
    $null = New-Item -ItemType Directory -Path $reportDirectory -Force
}

Write-HuduStructuredAssetImportLog -Message "Importing structured-list asset bundles from $structuredAssetImportJsonPath. DryRun=$structuredAssetImportDryRun; SkipExisting=$structuredAssetImportSkipExisting." -Color Cyan

$layoutNameCache = @{}
$existingAssetCache = @{}
$report = [System.Collections.Generic.List[object]]::new()
$createdCount = 0
$dryRunCount = 0
$skippedCount = 0
$failedCount = 0
$processedCount = 0

foreach ($bundle in @(Import-HuduStructuredAssetImportBundles -Path $bundlePaths)) {
    $listName = [string]$bundle.ListName
    if (-not (Test-HuduStructuredAssetImportListFilter -ListName $listName)) { continue }

    $map = Resolve-HuduStructuredAssetImportMap -Maps $HuduStructuredAssetImportMaps -ListName $listName
    $companyId = Resolve-HuduStructuredAssetImportCompanyId -Bundle $bundle -ResolveByName:$structuredAssetImportResolveCompanyByName
    if (-not $companyId -and $structuredAssetImportFallbackCompanyId -gt 0) {
        $companyId = $structuredAssetImportFallbackCompanyId
    }
    $layoutId = if ($map) { Resolve-HuduStructuredAssetImportLayoutId -Map $map -LayoutNameCache $layoutNameCache } else { $null }

    foreach ($item in @($bundle.Items)) {
        if ($structuredAssetImportMaxItems -gt 0 -and $processedCount -ge $structuredAssetImportMaxItems) { break }
        $processedCount++

        $assetName = if ($map) { Get-HuduStructuredAssetImportName -Item $item -Map $map } else { "SharePoint item $($item.SharePointItemId)" }
        $record = [ordered]@{
            Status           = $null
            CompanyId        = $companyId
            CompanyName      = $bundle.CompanyName
            ListName         = $listName
            AssetLayoutId    = $layoutId
            AssetName        = $assetName
            AssetId          = $null
            FieldCount       = 0
            SharePointItemId = $item.SharePointItemId
            MatchStatus      = $item.MatchStatus
            MatchAlias       = $item.MatchAlias
            MatchConfidence  = $item.MatchConfidence
            SourcePath       = $bundle.SourcePath
            WebUrl           = $item.WebUrl
            Error            = $null
        }

        if (-not $map) {
            $record.Status = 'SkippedNoMap'
            $skippedCount++
            $report.Add([PSCustomObject]$record)
            continue
        }
        if (-not $layoutId) {
            $record.Status = 'SkippedNoAssetLayout'
            $skippedCount++
            $report.Add([PSCustomObject]$record)
            continue
        }
        if (-not $companyId) {
            $record.Status = 'SkippedUnattributed'
            $skippedCount++
            $report.Add([PSCustomObject]$record)
            continue
        }

        try {
            $fields = ConvertTo-HuduStructuredAssetImportFields `
                -Item $item `
                -FieldMap ($map.FieldMap ?? @{}) `
                -AutoMapUnmappedFields:([bool]($map.AutoMapUnmappedFields ?? $false)) `
                -ExcludeFields @($map.ExcludeFields)
            $record.FieldCount = @($fields).Count

            $existingAsset = $null
            if ($structuredAssetImportSkipExisting) {
                $existingIndex = Get-HuduStructuredAssetImportExistingAssetIndex `
                    -CompanyId $companyId `
                    -AssetLayoutId $layoutId `
                    -Cache $existingAssetCache
                $assetKey = ConvertTo-HuduStructuredAssetImportKey $assetName
                if ($existingIndex.ContainsKey($assetKey)) {
                    $existingAsset = $existingIndex[$assetKey]
                }
            }

            if ($existingAsset) {
                $record.Status = 'SkippedExisting'
                $record.AssetId = $existingAsset.id ?? $existingAsset.Id
                $skippedCount++
                $report.Add([PSCustomObject]$record)
                continue
            }

            if ($structuredAssetImportDryRun) {
                $record.Status = 'DryRunCreate'
                $dryRunCount++
                $report.Add([PSCustomObject]$record)
                continue
            }

            $primaryValues = Get-HuduStructuredAssetImportPrimaryValues -Item $item -Map $map
            $newAssetParams = @{
                Name          = $assetName
                CompanyId     = $companyId
                AssetLayoutId = $layoutId
                Fields        = $fields
            }
            foreach ($primaryName in @($primaryValues.Keys)) {
                $newAssetParams[$primaryName] = $primaryValues[$primaryName]
            }

            $created = New-HuduAsset @newAssetParams
            $asset = $created.asset ?? $created
            $record.Status = 'Created'
            $record.AssetId = $asset.id ?? $asset.Id
            $createdCount++

            if ($structuredAssetImportSkipExisting) {
                $existingIndex = Get-HuduStructuredAssetImportExistingAssetIndex -CompanyId $companyId -AssetLayoutId $layoutId -Cache $existingAssetCache
                $existingIndex[(ConvertTo-HuduStructuredAssetImportKey $assetName)] = $asset
            }
        } catch {
            $record.Status = 'Failed'
            $record.Error = $_.Exception.Message
            $failedCount++
        }

        $report.Add([PSCustomObject]$record)
    }
}

$report | Export-Csv -LiteralPath $structuredAssetImportReportPath -NoTypeInformation -Encoding UTF8

Write-HuduStructuredAssetImportLog -Message "Structured asset import complete: $createdCount created, $dryRunCount dry-run create(s), $skippedCount skipped, $failedCount failed. Report: $structuredAssetImportReportPath" -Color Cyan
