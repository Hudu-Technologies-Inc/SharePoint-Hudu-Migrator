param(
    [hashtable]$Headers = $GraphHeaders,

    [ValidateSet('Graph', 'SharePointV2')]
    [string]$ApiMode = 'Graph',

    # Required for SharePointV2 mode. Example: contoso
    [string]$TenantName,

    [ValidateSet('All', 'Sites', 'Drives', 'Lists')]
    [string[]]$ManifestType = @('All'),

    [string]$ManifestDir = (Join-Path $PSScriptRoot 'out\sharepoint-manifests'),

    [string]$GeneratorPath,

    [string]$WorkQueuePath,

    [switch]$IncludeDocumentLibraryListItems,

    [switch]$IncludeDriveFolders,

    [switch]$ListMetadataOnly,

    [switch]$Force,

    [switch]$SkipWorkQueue
)

. (Join-Path $PSScriptRoot 'helpers\sharepoint\manifests.ps1')

$manifestSetParams = @{
    Headers                         = $Headers
    ApiMode                         = $ApiMode
    TenantName                      = $TenantName
    ManifestType                    = $ManifestType
    ManifestMode                    = 'Generate'
    ManifestDir                     = $ManifestDir
    IncludeDocumentLibraryListItems = $IncludeDocumentLibraryListItems
    ListMetadataOnly                = $ListMetadataOnly
    Force                           = $Force
}

if (-not [string]::IsNullOrWhiteSpace($GeneratorPath)) {
    $manifestSetParams.GeneratorPath = $GeneratorPath
}

$manifestSet = Initialize-SharePointManifestSet @manifestSetParams

$workQueueResult = $null
$workItemCount = 0

if (-not $SkipWorkQueue) {
    if ([string]::IsNullOrWhiteSpace($WorkQueuePath)) {
        $WorkQueuePath = Join-Path `
            -Path $manifestSet.ManifestDir `
            -ChildPath 'sharepoint-workqueue.jsonl'
    }

    $workItems = @(
        ConvertFrom-SharePointManifestSet `
            -ManifestSet $manifestSet `
            -IncludeDriveFolders:$IncludeDriveFolders
    )

    $workItemCount = $workItems.Count

    $workQueueResult = $workItems | Export-SharePointWorkQueue `
        -Path $WorkQueuePath
}

[pscustomobject]@{
    ManifestDir   = $manifestSet.ManifestDir
    ManifestPaths = $manifestSet.Paths
    ManifestCounts = $manifestSet.Counts
    WorkQueuePath = $workQueueResult.Path
    WorkItemCount = $workItemCount
}
