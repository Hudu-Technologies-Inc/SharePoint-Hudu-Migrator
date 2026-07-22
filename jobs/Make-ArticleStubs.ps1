##### Step 5, Stub Articles
$FolderResolutionCache = @{}
$BaseSitePath = (Get-Item $allSitesfolder).FullName
Set-PrintAndLog -message "BaseSitePath set to $BaseSitePath" -Color Cyan
$docsToStub = @($successConverted | Where-Object { $_.PSObject.Properties.Name -contains 'ContentPreview' -and $_.ContentPreview })

function Get-SharePointDocAttributionSourceText {
    param ($Doc)

    @(
        $Doc.SiteName
        $Doc.DriveName
        $Doc.parentDrivePath
        $Doc.RelativePath
        $Doc.LocalPath
        $Doc.webViewUrl
        $Doc.title
    ) -join ' '
}

function Resolve-SharePointDocClientAttribution {
    param (
        [Parameter(Mandatory)] [string]$SourceText
    )

    if (-not $RunSummary.SetupInfo.ClientAttributionAutoApply -or $ClientAttributionMap.Count -lt 1) {
        return $null
    }

    $attributionMatch = Resolve-HuduCompanyFromSharePointAttributionMap `
        -SourceText $SourceText `
        -AttributionMap ($ClientAttributionResolver ?? $ClientAttributionMap) `
        -AutoOnly `
        -AllowUnmatchedClientEntry:$RunSummary.SetupInfo.ClientAttributionCreateMissing `
        -MinScore $RunSummary.SetupInfo.ClientAttributionListItemMinScore `
        -MinGap $RunSummary.SetupInfo.ClientAttributionListItemMinGap

    if (-not $attributionMatch) { return $null }

    try {
        $attributionEntry = Confirm-HuduCompanyForSharePointAttributionMatch `
            -AttributionMatch $attributionMatch `
            -AttributionMap $ClientAttributionMap `
            -CreateMissing:$RunSummary.SetupInfo.ClientAttributionCreateMissing
    } catch {
        Set-PrintAndLog -message "Failed to create Hudu company for client list item '$($attributionMatch.Entry.ClientName)': $($_.Exception.Message)" -Color Red
        return $null
    }

    if (-not $attributionEntry -or -not $attributionEntry.HuduCompanyId) { return $null }

    [PSCustomObject]@{
        Entry = $attributionEntry
        Match = $attributionMatch
    }
}

function Resolve-SharePointDocClientDesignation {
    param ($Doc)

    Resolve-HuduCompanyFromClientDesignationMap `
        -SiteId $Doc.SiteId `
        -ListId $Doc.sharepointListId `
        -ClientDesignationMap $ClientDesignationMap `
        -UseSiteDesignation:$RunSummary.SetupInfo.ClientAttributionUseSiteDesignations `
        -UseListDesignation:$RunSummary.SetupInfo.ClientAttributionUseListDesignations
}

$docIDX=0
foreach ($doc in $docsToStub) {
    $docIDX += 1
    $completionPercentage = Get-PercentDone -Current $docIDX -Total $docsToStub.count

    # Determine CompanyId
    switch ([int]$RunSummary.JobInfo.MigrationDest.Identifier) {
        0 { $doc.CompanyId = $SingleCompanyChoice.id }
        1 { $doc.CompanyId = $null }
        3 {
            $clientDesignation = if ($RunSummary.SetupInfo.PreferClientAttributionOverSiteCompany) {
                Resolve-SharePointDocClientDesignation -Doc $doc
            } else {
                $null
            }

            $clientAttribution = if (-not $clientDesignation -and $RunSummary.SetupInfo.PreferClientAttributionOverSiteCompany) {
                Resolve-SharePointDocClientAttribution -SourceText (Get-SharePointDocAttributionSourceText -Doc $doc)
            } else {
                $null
            }

            if ($clientDesignation) {
                $doc.CompanyId = $clientDesignation.HuduCompanyId
                $doc | Add-Member -NotePropertyName ClientDesignationMatch -NotePropertyValue $clientDesignation -Force
                Set-PrintAndLog -message "Assigned '$($doc.title)' to per-site client designation '$($clientDesignation.HuduCompanyName)' ($($clientDesignation.Votes)/$($clientDesignation.TotalVotes) votes)." -Color Cyan
            } elseif ($clientAttribution) {
                $doc.CompanyId = $clientAttribution.Entry.HuduCompanyId
                $doc | Add-Member -NotePropertyName AttributionMatch -NotePropertyValue $clientAttribution.Entry -Force
                Set-PrintAndLog -message "Auto-attributed '$($doc.title)' to client list item '$($clientAttribution.Entry.RawTitle)' => Hudu company '$($clientAttribution.Entry.HuduCompanyName)' via '$($clientAttribution.Match.Alias)' ($($clientAttribution.Match.Confidence)%)." -Color Cyan
            } else {
                $siteCompany = Resolve-HuduCompanyFromSiteCompanyMap -SiteId $doc.SiteId -SiteName $doc.SiteName -SiteCompanyMap $SiteCompanyMap
                if ($siteCompany -and $siteCompany.HuduCompanyId) {
                $doc.CompanyId = $siteCompany.HuduCompanyId
                $doc | Add-Member -NotePropertyName SiteCompanyMatch -NotePropertyValue $siteCompany -Force
                Set-PrintAndLog -message "Assigned '$($doc.title)' to per-site Hudu company '$($siteCompany.HuduCompanyName)'." -Color Cyan
                } else {
                    Set-PrintAndLog -message "No client or per-site Hudu company available for '$($doc.SiteName)'; falling back to manual company selection." -Color Yellow
                    $doc.CompanyId = (
                        Select-ObjectFromList `
                            -message "Migrating Article: $($doc.ContentPreview ?? "no preview")... Which company to migrate into?" `
                            -objects $Attribution_Options
                    ).CompanyId
                }
            }
        }
        default {
            $clientDesignation = Resolve-SharePointDocClientDesignation -Doc $doc
            $clientAttribution = if (-not $clientDesignation) {
                Resolve-SharePointDocClientAttribution -SourceText (Get-SharePointDocAttributionSourceText -Doc $doc)
            } else {
                $null
            }

            if ($clientDesignation) {
                $doc.CompanyId = $clientDesignation.HuduCompanyId
                $doc | Add-Member -NotePropertyName ClientDesignationMatch -NotePropertyValue $clientDesignation -Force
                Set-PrintAndLog -message "Assigned '$($doc.title)' to per-site client designation '$($clientDesignation.HuduCompanyName)' ($($clientDesignation.Votes)/$($clientDesignation.TotalVotes) votes)." -Color Cyan
            } elseif ($clientAttribution) {
                $doc.CompanyId = $clientAttribution.Entry.HuduCompanyId
                $doc | Add-Member -NotePropertyName AttributionMatch -NotePropertyValue $clientAttribution.Entry -Force
                Set-PrintAndLog -message "Auto-attributed '$($doc.title)' to client list item '$($clientAttribution.Entry.RawTitle)' => Hudu company '$($clientAttribution.Entry.HuduCompanyName)' via '$($clientAttribution.Match.Alias)' ($($clientAttribution.Match.Confidence)%)." -Color Cyan
            } else {
                $doc.CompanyId = (
                    Select-ObjectFromList `
                        -message "Migrating Article: $($doc.ContentPreview ?? "no preview")... Which company to migrate into?" `
                        -objects $Attribution_Options
                ).CompanyId
            }
        }
    }

    if ($RunSummary.SetupInfo.SkipExistingArticles -and ($null -eq $doc.CompanyId -or $doc.CompanyId -ge 0)) {
        $existingArticle = Get-HuduExistingArticleByExactName -Title $doc.title -CompanyId $doc.CompanyId
        if ($existingArticle) {
            $existingArticleId = $existingArticle.id ?? $existingArticle.Id
            $existingArticleUrl = $existingArticle.url ?? $existingArticle.Url
            Set-PrintAndLog -message "Skipping '$($doc.title)' because Hudu article already exists in target company/global KB: $existingArticleUrl" -Color Yellow

            $doc | Add-Member -NotePropertyName ExistingHuduArticle -NotePropertyValue $existingArticle -Force
            $RunSummary.JobInfo.ArticlesSkipped++
            $RunSummary.Warnings += @{
                Message       = "Skipped SharePoint file because matching Hudu article already exists"
                Title         = $doc.title
                CompanyId     = $doc.CompanyId
                ExistingId    = $existingArticleId
                ExistingUrl   = $existingArticleUrl
                SharePointKey = $doc.SourceKey
            }

            if ($existingArticleUrl -or $existingArticleId) {
                $AllNewLinks.Add([PSCustomObject]@{
                    PageId    = $doc.id
                    PageTitle = $doc.title
                    HuduUrl   = $existingArticleUrl
                    ArticleId = $existingArticleId
                })
            }

            Write-SharePointExistingArticleSkipState `
                -Doc $doc `
                -ExistingArticle $existingArticle

            Write-Progress -Activity "Stubbing $($doc.title)" -Status "$completionPercentage%" -PercentComplete $completionPercentage
            continue
        }
    }

    # Resolve relative folder path
    $relativeFolderPath = $null
    if (-not [string]::IsNullOrWhiteSpace([string]$doc.RelativeFolderPath)) {
        $relativeFolderPath = [string]$doc.RelativeFolderPath
    } elseif ($doc.LocalPath -and $BaseSitePath) {
        $docDirectory = Split-Path -Path $doc.LocalPath -Parent
        $resolvedDocDirectory = [System.IO.Path]::GetFullPath($docDirectory)
        $resolvedBaseSitePath = [System.IO.Path]::GetFullPath($BaseSitePath).TrimEnd('\')

        if ($resolvedDocDirectory.StartsWith($resolvedBaseSitePath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $relativeFolderPath = $resolvedDocDirectory.Substring($resolvedBaseSitePath.Length).TrimStart('\')
        }
    }

    # Build key and resolve/create folder via cache
    $key = "$($doc.CompanyId)-$relativeFolderPath"
    if (-not $FolderResolutionCache.ContainsKey($key) -and $relativeFolderPath) {
        $folderParts = $relativeFolderPath -split '\\'
        $resolvedFolder = Initialize-HuduFolder -FolderPath $folderParts -CompanyId $doc.CompanyId

        if ($resolvedFolder) {
            $FolderResolutionCache[$key] = $resolvedFolder
            Set-PrintAndLog -message "Created folder for path: $relativeFolderPath with ID $($resolvedFolder.id)" -Color Cyan
        } else {
            Set-PrintAndLog -message "Failed to create folder for: $relativeFolderPath" -Color Red
        }
    }

    if ($FolderResolutionCache.ContainsKey($key)) {
        $doc.HuduFolder = $FolderResolutionCache[$key]
        $doc.HuduFolderId = $doc.HuduFolder.id
    }

    # Stub article
    if ($null -eq $doc.CompanyId -or $doc.CompanyId -eq 0) {
        Set-PrintAndLog -message "Stubbing global KB article" -Color Yellow
        $doc.stub = New-HuduStubArticle -Title $doc.title -Content "$($doc.ContentPreview)" -FolderId $doc.HuduFolderId
    }
    elseif ($doc.CompanyId -lt 0) {
        Set-PrintAndLog -message "Skipping doc/article transfer for $($doc.title)" -Color Gray
        $RunSummary.Warnings += @{
            Message     = "User elected to skip doc/article transfer for $($doc.title)"
            docSkipped  = "doc with ID $($doc.id), titled $($doc.title) was skipped. $($doc.FullUrl ?? '')"
        }
        $RunSummary.JobInfo.Skipped++
        continue
    }
    else {
        Set-PrintAndLog -message "Stubbing KB article for Hudu company ID: $($doc.CompanyId)" -Color Yellow
        $doc.stub = New-HuduStubArticle -Title $doc.title -Content "$($doc.ContentPreview)" -CompanyId $doc.CompanyId -FolderId $doc.HuduFolderId
    }

    # Post-processing
    Set-PrintAndLog -message "Article $($doc.title)  with id $($doc.stub.id); $($doc.stub | ConvertTo-Json -Depth 3)" -Color Green

    if (-not $doc.stub) {
        $ErrorObject = @{
            Error = "Error stubbing article with id $($doc.id), title $($doc.title)"
        }
        Write-ErrorObjectsToFile -name "Stub-$($doc.title)" -ErrorObject $ErrorObject
        $RunSummary.Errors.Add($ErrorObject)
        $RunSummary.JobInfo.ArticlesErrored++
        continue
    }

    $RunSummary.JobInfo.ArticlesCreated++
    $RunSummary.JobInfo.LinksCreated++
    $AllNewLinks.Add([PSCustomObject]@{
        PageId    = $doc.id
        PageTitle = $doc.title
        HuduUrl   = $doc.stub.url
        ArticleId = $doc.stub.id
    })

    $StubbedArticles += $doc
    Write-Progress -Activity "Stubbing $($doc.title)" -Status "$completionPercentage%" -PercentComplete $completionPercentage
}
