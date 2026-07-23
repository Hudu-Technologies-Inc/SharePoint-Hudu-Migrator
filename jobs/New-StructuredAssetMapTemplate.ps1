##### Optional helper, generate a starter Hudu structured asset map from exported list JSON

$structuredAssetTemplateRoot = if (-not [string]::IsNullOrWhiteSpace([string]$workdir)) {
    [string]$workdir
} elseif (-not [string]::IsNullOrWhiteSpace([string]$PSScriptRoot)) {
    (Split-Path -Parent $PSScriptRoot)
} else {
    (Get-Location).Path
}

foreach ($helperName in @('attribution.ps1', 'structuredlists.ps1')) {
    $helperPath = Join-Path (Join-Path $structuredAssetTemplateRoot 'helpers') $helperName
    if (Test-Path -LiteralPath $helperPath -PathType Leaf) {
        . $helperPath
    }
}

function ConvertTo-HuduStructuredAssetTemplateKey {
    param ($Value)

    if (Get-Command ConvertTo-AttributionNormalizedText -ErrorAction SilentlyContinue) {
        return (ConvertTo-AttributionNormalizedText $Value)
    }

    if ($null -eq $Value) { return "" }
    return (([string]$Value).ToLowerInvariant() -replace '[^a-z0-9]+', ' ').Trim()
}

function Get-HuduStructuredAssetTemplateBaseListName {
    param ($Name)

    if (Get-Command Get-SharePointStructuredListBaseName -ErrorAction SilentlyContinue) {
        return (Get-SharePointStructuredListBaseName $Name)
    }

    if ($null -eq $Name) { return "" }
    return ([string]$Name -replace '_\[[^\]]+\]$', '' -replace '\[[^\]]+\]$', '').Trim()
}

function ConvertFrom-HuduStructuredAssetTemplateFieldName {
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

function ConvertTo-HuduStructuredAssetTemplatePath {
    param ([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if ([System.IO.Path]::IsPathRooted($Path)) { return [System.IO.Path]::GetFullPath($Path) }
    return [System.IO.Path]::GetFullPath((Join-Path $structuredAssetTemplateRoot $Path))
}

function ConvertTo-HuduStructuredAssetTemplateLiteral {
    param ([string]$Value)

    if ($null -eq $Value) { $Value = "" }
    "'" + ([string]$Value).Replace("'", "''") + "'"
}

function Get-HuduStructuredAssetTemplateBundlePaths {
    param ([Parameter(Mandatory)] [string]$Path)

    if (Test-Path -LiteralPath $Path -PathType Container) {
        return @(
            Get-ChildItem -LiteralPath $Path -Recurse -Filter '*.json' -File |
                Sort-Object FullName |
                ForEach-Object { $_.FullName }
        )
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Structured-list JSON path was not found: $Path"
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

$structuredAssetTemplateJsonPath = ConvertTo-HuduStructuredAssetTemplatePath (
    $HuduStructuredAssetImportJsonPath ??
    $RunSummary.OutputJsonFiles.StructuredListJsonDir ??
    (Join-Path 'logs' 'structured-list-json')
)
$structuredAssetTemplateOutputPath = ConvertTo-HuduStructuredAssetTemplatePath (
    $HuduStructuredAssetMapTemplateOutputPath ??
    (Join-Path $structuredAssetTemplateRoot 'structured-asset-maps.generated.ps1')
)

$listFieldMap = @{}
foreach ($bundlePath in @(Get-HuduStructuredAssetTemplateBundlePaths -Path $structuredAssetTemplateJsonPath)) {
    $json = Get-Content -LiteralPath $bundlePath -Raw | ConvertFrom-Json
    foreach ($bundle in @($json)) {
        if (-not $bundle.Items) { continue }

        $listName = Get-HuduStructuredAssetTemplateBaseListName $bundle.ListName
        $listKey = ConvertTo-HuduStructuredAssetTemplateKey $listName
        if (-not $listKey) { continue }

        if (-not $listFieldMap.ContainsKey($listKey)) {
            $listFieldMap[$listKey] = [ordered]@{
                ListName = $listName
                Fields   = [ordered]@{}
            }
        }

        foreach ($item in @($bundle.Items)) {
            if (-not $item.Fields) { continue }
            foreach ($property in @($item.Fields.PSObject.Properties)) {
                if ($property.Name -like '@odata*') { continue }
                $decodedName = ConvertFrom-HuduStructuredAssetTemplateFieldName $property.Name
                if ([string]::IsNullOrWhiteSpace($decodedName)) { continue }
                if (-not $listFieldMap[$listKey].Fields.Contains($property.Name)) {
                    $listFieldMap[$listKey].Fields[$property.Name] = $decodedName
                }
            }
        }
    }
}

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# Generated starter map. Review Hudu layout names/IDs and field labels before live import.')
$lines.Add('$HuduStructuredAssetImportMaps = @{')

foreach ($list in @($listFieldMap.Values | Sort-Object ListName)) {
    $defaultNameField = if ($list.Fields.Contains('Title')) {
        'Title'
    } elseif ($list.Fields.Contains('LinkTitle')) {
        'LinkTitle'
    } else {
        @($list.Fields.Keys | Select-Object -First 1)
    }

    $lines.Add("    $(ConvertTo-HuduStructuredAssetTemplateLiteral $list.ListName) = @{")
    $lines.Add("        AssetLayoutName = $(ConvertTo-HuduStructuredAssetTemplateLiteral $list.ListName)")
    $lines.Add("        NameField       = $(ConvertTo-HuduStructuredAssetTemplateLiteral $defaultNameField)")
    $lines.Add('        FieldMap        = [ordered]@{')

    foreach ($fieldName in @($list.Fields.Keys | Sort-Object)) {
        $lines.Add("            $(ConvertTo-HuduStructuredAssetTemplateLiteral $fieldName) = $(ConvertTo-HuduStructuredAssetTemplateLiteral $list.Fields[$fieldName])")
    }

    $lines.Add('        }')
    $lines.Add('    }')
    $lines.Add('')
}

$lines.Add('}')

$outputDirectory = Split-Path -Parent $structuredAssetTemplateOutputPath
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    $null = New-Item -ItemType Directory -Path $outputDirectory -Force
}

$lines | Out-File -LiteralPath $structuredAssetTemplateOutputPath -Encoding UTF8
Write-Host "Wrote structured asset map template: $structuredAssetTemplateOutputPath" -ForegroundColor Cyan
