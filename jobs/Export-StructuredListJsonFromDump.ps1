##### Optional helper, build structured-list JSON bundles from an existing SharePoint manifest dump

$structuredDumpRoot = if (-not [string]::IsNullOrWhiteSpace([string]$workdir)) {
    [string]$workdir
} elseif (-not [string]::IsNullOrWhiteSpace([string]$PSScriptRoot)) {
    (Split-Path -Parent $PSScriptRoot)
} else {
    (Get-Location).Path
}

foreach ($helperPath in @(
    (Join-Path (Join-Path $structuredDumpRoot 'helpers') 'attribution.ps1'),
    (Join-Path (Join-Path $structuredDumpRoot 'helpers') 'structuredlists.ps1'),
    (Join-Path (Join-Path (Join-Path $structuredDumpRoot 'helpers') 'sharepoint') 'manifests.ps1')
)) {
    if (Test-Path -LiteralPath $helperPath -PathType Leaf) {
        . $helperPath
    }
}

function Write-StructuredDumpExportLog {
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

function Resolve-StructuredDumpExportPath {
    param ([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if ([System.IO.Path]::IsPathRooted($Path)) { return [System.IO.Path]::GetFullPath($Path) }
    return [System.IO.Path]::GetFullPath((Join-Path $structuredDumpRoot $Path))
}

function ConvertTo-StructuredDumpDesignationLookup {
    param ($DesignationMap)

    if (-not $DesignationMap) { return $null }

    $sites = @($DesignationMap.Sites)
    $lists = @($DesignationMap.Lists)
    $siteById = @{}
    $listByKey = @{}

    foreach ($site in $sites) {
        $key = [string]($site.Key ?? $site.SiteId)
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $siteById[$key] = $site
        }
    }

    foreach ($list in $lists) {
        $key = [string]($list.Key ?? $list.ListKey)
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $listByKey[$key] = $list
        }
    }

    [PSCustomObject]@{
        Sites     = $sites
        Lists     = $lists
        SiteById  = $siteById
        ListByKey = $listByKey
    }
}

$structuredDumpManifestPath = Resolve-StructuredDumpExportPath (
    $SharePointStructuredListJsonManifestPath ??
    (Join-Path (Join-Path 'out' 'sharepoint-manifests') 'graph-sharepoint-all-manifest.json')
)
$structuredDumpManifestDir = Resolve-StructuredDumpExportPath (
    $SharePointStructuredListJsonManifestDir ??
    (Split-Path -Parent $structuredDumpManifestPath)
)
$structuredDumpOutputDirectory = Resolve-StructuredDumpExportPath (
    $SharePointStructuredListJsonOutputDirectory ??
    $RunSummary.OutputJsonFiles.StructuredListJsonDir ??
    (Join-Path 'logs' 'structured-list-json')
)
$structuredDumpIndexPath = Resolve-StructuredDumpExportPath (
    $SharePointStructuredListJsonIndexPath ??
    $RunSummary.OutputJsonFiles.StructuredListJsonIndex ??
    (Join-Path 'logs' 'structured-list-json-index.csv')
)
$structuredDumpAttributionMapPath = Resolve-StructuredDumpExportPath (
    $SharePointStructuredListJsonAttributionMapPath ??
    (Join-Path 'logs' 'client-attribution-map.json')
)
$structuredDumpDesignationMapPath = Resolve-StructuredDumpExportPath (
    $SharePointStructuredListJsonDesignationMapPath ??
    (Join-Path 'logs' 'client-designation-map.json')
)
$structuredDumpListNames = @(
    if ($null -ne $SharePointStructuredListJsonNames) {
        @($SharePointStructuredListJsonNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    } else {
        @()
    }
)
$structuredDumpPrimaryAttributionFieldNames = @(
    if ($null -ne $SharePointClientAttributionFieldNames) {
        @($SharePointClientAttributionFieldNames)
    } elseif ($RunSummary -and $RunSummary.SetupInfo -and $RunSummary.SetupInfo.ClientAttributionFieldNames) {
        @($RunSummary.SetupInfo.ClientAttributionFieldNames)
    } else {
        @("Select a Client", "Client", "Customer", "Company", "LinkTitle")
    }
)
$structuredDumpUseListDesignation = [bool]($SharePointClientAttributionUseListDesignations ?? $RunSummary.SetupInfo.ClientAttributionUseListDesignations ?? $true)
$structuredDumpUseSiteDesignation = [bool]($SharePointClientAttributionUseSiteDesignations ?? $RunSummary.SetupInfo.ClientAttributionUseSiteDesignations ?? $true)

if ($structuredDumpListNames.Count -lt 1) {
    throw "No SharePoint structured list names are configured. Set `$SharePointStructuredListJsonNames before running this job."
}

if (-not (Test-Path -LiteralPath $structuredDumpManifestPath -PathType Leaf)) {
    throw "SharePoint manifest dump was not found: $structuredDumpManifestPath"
}

$manifestSet = Import-SharePointManifestSet -ManifestPath $structuredDumpManifestPath -ManifestDir $structuredDumpManifestDir

$attributionMap = @()
if (Test-Path -LiteralPath $structuredDumpAttributionMapPath -PathType Leaf) {
    $attributionMap = @(Get-Content -LiteralPath $structuredDumpAttributionMapPath -Raw | ConvertFrom-Json)
    if ($attributionMap.Count -gt 0) {
        $attributionMap = Resolve-SharePointClientAttributionLookup -AttributionMap $attributionMap
    }
} else {
    Write-StructuredDumpExportLog -Message "No client attribution map found at $structuredDumpAttributionMapPath; export will be unattributed unless designation map applies." -Color Yellow
}

$designationMap = $null
if (Test-Path -LiteralPath $structuredDumpDesignationMapPath -PathType Leaf) {
    $designationMap = ConvertTo-StructuredDumpDesignationLookup (Get-Content -LiteralPath $structuredDumpDesignationMapPath -Raw | ConvertFrom-Json)
}

Write-StructuredDumpExportLog -Message "Exporting structured-list JSON from dump for list(s): $($structuredDumpListNames -join ', ')" -Color Cyan

$result = Export-SharePointStructuredListJson `
    -ManifestSet $manifestSet `
    -ListNames $structuredDumpListNames `
    -AttributionMap $attributionMap `
    -ClientDesignationMap $designationMap `
    -PrimaryAttributionFieldNames $structuredDumpPrimaryAttributionFieldNames `
    -UseListDesignation:$structuredDumpUseListDesignation `
    -UseSiteDesignation:$structuredDumpUseSiteDesignation `
    -OutputDirectory $structuredDumpOutputDirectory `
    -IndexPath $structuredDumpIndexPath

Write-StructuredDumpExportLog -Message "Structured-list JSON export from dump complete: $($result.Items) item(s) in $($result.Bundles) bundle(s)." -Color Cyan
Write-StructuredDumpExportLog -Message "Structured-list JSON directory: $($result.OutputDirectory)" -Color DarkMagenta
Write-StructuredDumpExportLog -Message "Structured-list JSON index: $($result.IndexPath)" -Color DarkMagenta

$result
