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

function Get-DefaultSharePointManifestGeneratorPath {
    [CmdletBinding()]
    param()

    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

    Join-Path `
        -Path $repoRoot `
        -ChildPath 'One-Offs\dump-sharepoint-manifest.ps1\dump-manifest.ps1'
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

    $manifests = foreach ($path in $paths) {
        Import-SharePointManifestJson -Path $path
    }

    $counts = [ordered]@{
        sites      = 0
        drives     = 0
        driveItems = 0
        lists      = 0
        listItems  = 0
        errors     = 0
    }

    foreach ($manifest in $manifests) {
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
        Paths       = $paths
        Manifests   = @($manifests)
        Counts      = [pscustomobject]$counts
    }
}

function Initialize-SharePointManifestSet {
    [CmdletBinding()]
    param(
        [hashtable]$Headers = $GraphHeaders,

        [ValidateSet('Graph', 'SharePointV2')]
        [string]$ApiMode = 'Graph',

        [string]$TenantName,

        [ValidateSet('All', 'Sites', 'Drives', 'Lists')]
        [string[]]$ManifestType = @('All'),

        [ValidateSet('Auto', 'Generate', 'UseExisting')]
        [string]$ManifestMode = 'Auto',

        [string]$ManifestDir = (Join-Path $PWD 'out\sharepoint-manifests'),

        [string]$GeneratorPath = (Get-DefaultSharePointManifestGeneratorPath),

        [switch]$IncludeDocumentLibraryListItems,

        [switch]$ListMetadataOnly,

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

        if (-not (Test-Path -LiteralPath $GeneratorPath -PathType Leaf)) {
            throw "SharePoint manifest generator not found: $GeneratorPath"
        }

        Write-Host "Generating SharePoint manifest: $manifestPath" -ForegroundColor Cyan

        $generatorParams = @{
            Headers                         = $Headers
            ApiMode                         = $ApiMode
            ManifestType                    = $ManifestType
            OutputPath                      = $manifestPath
            IncludeDocumentLibraryListItems = $IncludeDocumentLibraryListItems
            ListMetadataOnly                = $ListMetadataOnly
        }

        if (-not [string]::IsNullOrWhiteSpace($TenantName)) {
            $generatorParams.TenantName = $TenantName
        }

        $generatorResult = & $GeneratorPath @generatorParams
            $generatorResult | Out-Host

            $writtenManifestPath = @($generatorResult |
                Where-Object { $_.PSObject.Properties.Name -contains 'Path' } |
                Select-Object -Last 1).Path

            if (-not [string]::IsNullOrWhiteSpace($writtenManifestPath)) {
                $manifestPath = $writtenManifestPath
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