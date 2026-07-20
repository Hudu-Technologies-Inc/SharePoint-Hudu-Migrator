function Get-SharePointStructuredListBaseName {
    param ($Name)

    if ($null -eq $Name) { return "" }
    $baseName = [string]$Name
    $baseName = $baseName -replace '_\[[^\]]+\]$', ''
    $baseName = $baseName -replace '\[[^\]]+\]$', ''
    return $baseName.Trim()
}

function Test-SharePointStructuredListName {
    param (
        [Parameter(Mandatory)] [string]$ListName,
        [Parameter(Mandatory)] [string[]]$ConfiguredNames
    )

    $normalizedListName = ConvertTo-AttributionNormalizedText $ListName
    $normalizedBaseName = ConvertTo-AttributionNormalizedText (Get-SharePointStructuredListBaseName $ListName)

    foreach ($configuredName in @($ConfiguredNames)) {
        $normalizedConfigured = ConvertTo-AttributionNormalizedText $configuredName
        if (-not $normalizedConfigured) { continue }

        if ($normalizedListName -eq $normalizedConfigured -or $normalizedBaseName -eq $normalizedConfigured) {
            return $true
        }
    }

    return $false
}

function Get-SharePointListItemAttributionSourceText {
    param (
        $SiteEntry,
        $ListEntry,
        $Item
    )

    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($value in @(
        $SiteEntry.metadata.displayName
        $SiteEntry.metadata.name
        $ListEntry.metadata.displayName
        $ListEntry.metadata.name
        $Item.webUrl
    )) {
        if ($value) { $parts.Add([string]$value) }
    }

    if ($Item.fields) {
        foreach ($property in $Item.fields.PSObject.Properties) {
            if ($property.Name -like '@odata*') { continue }
            if ($null -eq $property.Value) { continue }
            if ($property.Value -is [string] -or $property.Value -is [ValueType]) {
                $parts.Add([string]$property.Value)
            }
        }
    }

    return ($parts -join ' ')
}

function Get-SharePointListItemPrimaryAttributionSourceText {
    param ($Item)

    if (-not $Item.fields) { return $null }
    if (-not ($Item.fields.PSObject.Properties.Name -contains 'LinkTitle')) { return $null }

    $linkTitle = $Item.fields.LinkTitle
    if ([string]::IsNullOrWhiteSpace([string]$linkTitle)) { return $null }

    return [string]$linkTitle
}

function Get-SafeStructuredListPathName {
    param (
        [string]$Name,
        [int]$MaxLength = 90
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { $Name = "unnamed" }
    $safe = ($Name -replace '[\\/:*?"<>|]', '_') -replace '\s+', ' '
    $safe = $safe.Trim()
    if ($safe.Length -gt $MaxLength) {
        $safe = $safe.Substring(0, $MaxLength).Trim()
    }
    return $safe
}

function Export-SharePointStructuredListJson {
    param (
        [Parameter(Mandatory)] $ManifestSet,
        [Parameter(Mandatory)] [string[]]$ListNames,
        $AttributionMap = @(),
        [Parameter(Mandatory)] [string]$OutputDirectory,
        [Parameter(Mandatory)] [string]$IndexPath
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        $null = New-Item -ItemType Directory -Path $OutputDirectory -Force
    }

    $recordsByKey = @{}
    $indexRows = [System.Collections.Generic.List[object]]::new()
    $primaryAttributionCache = @{}
    $hasAttributionMap = if (
        $AttributionMap.PSObject.Properties.Name -contains 'IsSharePointClientAttributionLookup' -and
        $AttributionMap.IsSharePointClientAttributionLookup
    ) {
        @($AttributionMap.Entries).Count -gt 0
    } else {
        @($AttributionMap).Count -gt 0
    }

    foreach ($manifest in @($ManifestSet.Manifests)) {
        foreach ($siteEntry in @($manifest.sites)) {
            foreach ($listEntry in @($siteEntry.lists)) {
                $listName = $listEntry.metadata.displayName ?? $listEntry.metadata.name
                if (-not (Test-SharePointStructuredListName -ListName $listName -ConfiguredNames $ListNames)) {
                    continue
                }

                $listBaseName = Get-SharePointStructuredListBaseName $listName
                foreach ($item in @($listEntry.items)) {
                    $match = $null
                    if ($hasAttributionMap) {
                        $primarySourceText = Get-SharePointListItemPrimaryAttributionSourceText -Item $item
                        if (-not [string]::IsNullOrWhiteSpace([string]$primarySourceText)) {
                            $primaryCacheKey = ConvertTo-AttributionNormalizedText $primarySourceText
                            if ($primaryAttributionCache.ContainsKey($primaryCacheKey)) {
                                $match = $primaryAttributionCache[$primaryCacheKey]
                            } else {
                                $match = Resolve-HuduCompanyFromSharePointAttributionMap -SourceText $primarySourceText -AttributionMap $AttributionMap -AutoOnly
                                $primaryAttributionCache[$primaryCacheKey] = $match
                            }
                        }

                        if (-not $match) {
                            $sourceText = Get-SharePointListItemAttributionSourceText -SiteEntry $siteEntry -ListEntry $listEntry -Item $item
                            $match = Resolve-HuduCompanyFromSharePointAttributionMap -SourceText $sourceText -AttributionMap $AttributionMap -AutoOnly
                        }
                    }

                    $companyId = $match.Entry.HuduCompanyId
                    $companyName = $match.Entry.HuduCompanyName
                    $matchStatus = if ($match) { 'Auto' } else { 'Unattributed' }
                    if (-not $companyName) { $companyName = 'Unattributed' }

                    $companyFolderName = if ($companyId) {
                        "{0} [{1}]" -f (Get-SafeStructuredListPathName $companyName), $companyId
                    } else {
                        "Unattributed"
                    }

                    $companyFolder = Join-Path $OutputDirectory $companyFolderName
                    if (-not (Test-Path -LiteralPath $companyFolder)) {
                        $null = New-Item -ItemType Directory -Path $companyFolder -Force
                    }

                    $safeListName = Get-SafeStructuredListPathName $listBaseName
                    $outputPath = Join-Path $companyFolder "$safeListName.json"
                    $key = "$companyFolderName|$safeListName"

                    if (-not $recordsByKey.ContainsKey($key)) {
                        $recordsByKey[$key] = [PSCustomObject]@{
                            CompanyId   = $companyId
                            CompanyName = $companyName
                            ListName    = $listBaseName
                            SourceLists = [System.Collections.Generic.List[object]]::new()
                            Items       = [System.Collections.Generic.List[object]]::new()
                            OutputPath  = $outputPath
                        }
                    }

                    $bundle = $recordsByKey[$key]
                    $sourceListKey = "$($siteEntry.metadata.id)|$($listEntry.metadata.id)"
                    if (-not @($bundle.SourceLists | ForEach-Object { $_.SourceListKey }).Contains($sourceListKey)) {
                        $bundle.SourceLists.Add([PSCustomObject]@{
                            SourceListKey = $sourceListKey
                            SiteId        = $siteEntry.metadata.id
                            SiteName      = $siteEntry.metadata.displayName
                            SiteWebUrl    = $siteEntry.metadata.webUrl
                            ListId        = $listEntry.metadata.id
                            ListName      = $listName
                            ListWebUrl    = $listEntry.metadata.webUrl
                        })
                    }

                    $bundle.Items.Add([PSCustomObject]@{
                        SharePointItemId = $item.id
                        WebUrl           = $item.webUrl
                        CreatedDateTime  = $item.createdDateTime
                        LastModified     = $item.lastModifiedDateTime
                        MatchStatus      = $matchStatus
                        MatchAlias       = $match.Alias
                        MatchConfidence  = $match.Confidence
                        Fields           = $item.fields
                    })

                    $indexRows.Add([PSCustomObject]@{
                        CompanyId        = $companyId
                        CompanyName      = $companyName
                        ListName         = $listBaseName
                        SharePointItemId = $item.id
                        MatchStatus      = $matchStatus
                        MatchAlias       = $match.Alias
                        OutputPath       = $outputPath
                    })
                }
            }
        }
    }

    foreach ($bundle in $recordsByKey.Values) {
        $bundle |
            Select-Object CompanyId, CompanyName, ListName, SourceLists, Items |
            ConvertTo-Json -Depth 60 |
            Out-File -LiteralPath $bundle.OutputPath -Encoding UTF8
    }

    $indexRows |
        Export-Csv -Path $IndexPath -NoTypeInformation -Encoding UTF8

    [PSCustomObject]@{
        Bundles = $recordsByKey.Count
        Items   = $indexRows.Count
        OutputDirectory = $OutputDirectory
        IndexPath = $IndexPath
    }
}
