##### Step 4-pre, Skip already-existing Hudu articles before conversion

if (
    -not $RunSummary.SetupInfo.SkipExistingArticles -or
    $null -eq $AllDiscoveredFiles -or
    $AllDiscoveredFiles.Count -lt 1
) {
    return
}

function Get-SharePointEarlyArticleTitle {
    param ($File)

    $extension = if ($File.LocalPath) {
        [System.IO.Path]::GetExtension([string]$File.LocalPath).ToLowerInvariant()
    } else {
        [System.IO.Path]::GetExtension([string]$File.Name).ToLowerInvariant()
    }

    if ($extension -eq ".pdf" -and $RunSummary.SetupInfo.PdfUploadAsFile) {
        return [System.IO.Path]::GetFileName([string]($File.LocalPath ?? $File.Name))
    }

    return [string]($File.title ?? (Get-SafeTitle -Name $File.Name))
}

function Test-SharePointEarlySkipEligibleFile {
    param ($File)

    if (-not $File.LocalPath) { return $false }

    $extension = [System.IO.Path]::GetExtension([string]$File.LocalPath).ToLowerInvariant()
    $indexOnlyExtensions = @($RunSummary.SetupInfo.IndexOnlyExtensions) | ForEach-Object {
        $configuredExtension = ([string]$_).Trim().ToLowerInvariant()
        if ($configuredExtension -and -not $configuredExtension.StartsWith(".")) {
            ".$configuredExtension"
        } else {
            $configuredExtension
        }
    }

    return ($indexOnlyExtensions -notcontains $extension)
}

function Get-SharePointEarlyDocAttributionSourceText {
    param ($File)

    @(
        $File.SiteName
        $File.DriveName
        $File.parentDrivePath
        $File.RelativePath
        $File.LocalPath
        $File.webViewUrl
        (Get-SharePointEarlyArticleTitle -File $File)
    ) -join ' '
}

function Resolve-SharePointEarlyClientAttribution {
    param ($File)

    if (
        -not $RunSummary.SetupInfo.ClientAttributionAutoApply -or
        $null -eq $ClientAttributionMap -or
        $ClientAttributionMap.Count -lt 1
    ) {
        return $null
    }

    $attributionMatch = Resolve-HuduCompanyFromSharePointAttributionMap `
        -SourceText (Get-SharePointEarlyDocAttributionSourceText -File $File) `
        -AttributionMap ($ClientAttributionResolver ?? $ClientAttributionMap) `
        -AutoOnly `
        -MinScore $RunSummary.SetupInfo.ClientAttributionListItemMinScore `
        -MinGap $RunSummary.SetupInfo.ClientAttributionListItemMinGap

    if (-not $attributionMatch -or -not $attributionMatch.Entry.HuduCompanyId) {
        return $null
    }

    return $attributionMatch.Entry
}

function Resolve-SharePointEarlyClientDesignation {
    param ($File)

    Resolve-HuduCompanyFromClientDesignationMap `
        -SiteId $File.SiteId `
        -ListId $File.sharepointListId `
        -ClientDesignationMap $ClientDesignationMap `
        -UseSiteDesignation:$RunSummary.SetupInfo.ClientAttributionUseSiteDesignations `
        -UseListDesignation:$RunSummary.SetupInfo.ClientAttributionUseListDesignations
}

function Resolve-SharePointEarlyExistingArticleTarget {
    param ($File)

    switch ([int]$RunSummary.JobInfo.MigrationDest.Identifier) {
        0 {
            return [PSCustomObject]@{
                Resolved  = $true
                CompanyId = $SingleCompanyChoice.id
                Reason    = "single_company"
            }
        }
        1 {
            return [PSCustomObject]@{
                Resolved  = $true
                CompanyId = $null
                Reason    = "global_kb"
            }
        }
        3 {
            $clientDesignation = if ($RunSummary.SetupInfo.PreferClientAttributionOverSiteCompany) {
                Resolve-SharePointEarlyClientDesignation -File $File
            } else {
                $null
            }

            if ($clientDesignation -and $clientDesignation.HuduCompanyId) {
                return [PSCustomObject]@{
                    Resolved  = $true
                    CompanyId = $clientDesignation.HuduCompanyId
                    Reason    = "client_designation"
                }
            }

            $clientAttribution = if ($RunSummary.SetupInfo.PreferClientAttributionOverSiteCompany) {
                Resolve-SharePointEarlyClientAttribution -File $File
            } else {
                $null
            }

            if ($clientAttribution -and $clientAttribution.HuduCompanyId) {
                return [PSCustomObject]@{
                    Resolved  = $true
                    CompanyId = $clientAttribution.HuduCompanyId
                    Reason    = "client_attribution"
                }
            }

            $siteCompany = Resolve-HuduCompanyFromSiteCompanyMap -SiteId $File.SiteId -SiteName $File.SiteName -SiteCompanyMap $SiteCompanyMap
            if ($siteCompany -and $siteCompany.HuduCompanyId) {
                return [PSCustomObject]@{
                    Resolved  = $true
                    CompanyId = $siteCompany.HuduCompanyId
                    Reason    = "site_company"
                }
            }
        }
        default {
            $clientDesignation = Resolve-SharePointEarlyClientDesignation -File $File
            if ($clientDesignation -and $clientDesignation.HuduCompanyId) {
                return [PSCustomObject]@{
                    Resolved  = $true
                    CompanyId = $clientDesignation.HuduCompanyId
                    Reason    = "client_designation"
                }
            }

            $clientAttribution = Resolve-SharePointEarlyClientAttribution -File $File
            if ($clientAttribution -and $clientAttribution.HuduCompanyId) {
                return [PSCustomObject]@{
                    Resolved  = $true
                    CompanyId = $clientAttribution.HuduCompanyId
                    Reason    = "client_attribution"
                }
            }
        }
    }

    [PSCustomObject]@{
        Resolved  = $false
        CompanyId = $null
        Reason    = "unresolved"
    }
}

$remainingFiles = [System.Collections.ArrayList]@()
$earlySkippedCount = 0

foreach ($file in @($AllDiscoveredFiles)) {
    if (-not (Test-SharePointEarlySkipEligibleFile -File $file)) {
        [void]$remainingFiles.Add($file)
        continue
    }

    $target = Resolve-SharePointEarlyExistingArticleTarget -File $file
    if (-not $target.Resolved -or ($null -ne $target.CompanyId -and $target.CompanyId -lt 0)) {
        [void]$remainingFiles.Add($file)
        continue
    }

    $articleTitle = Get-SharePointEarlyArticleTitle -File $file
    $existingArticle = Get-HuduExistingArticleByExactName -Title $articleTitle -CompanyId $target.CompanyId
    if (-not $existingArticle) {
        [void]$remainingFiles.Add($file)
        continue
    }

    $existingArticleId = $existingArticle.id ?? $existingArticle.Id
    $existingArticleUrl = $existingArticle.url ?? $existingArticle.Url
    Set-PrintAndLog -message "Early skip '$articleTitle' because Hudu article already exists in target company/global KB: $existingArticleUrl" -Color Yellow

    $file | Add-Member -NotePropertyName CompanyId -NotePropertyValue $target.CompanyId -Force
    $file | Add-Member -NotePropertyName ExistingHuduArticle -NotePropertyValue $existingArticle -Force
    $RunSummary.JobInfo.ArticlesSkipped++
    $RunSummary.Warnings += @{
        Message       = "Skipped SharePoint file before conversion because matching Hudu article already exists"
        Title         = $articleTitle
        CompanyId     = $target.CompanyId
        ExistingId    = $existingArticleId
        ExistingUrl   = $existingArticleUrl
        SharePointKey = $file.SourceKey
        Reason        = $target.Reason
    }

    if ($existingArticleUrl -or $existingArticleId) {
        [void]$AllNewLinks.Add([PSCustomObject]@{
            PageId    = $file.id
            PageTitle = $articleTitle
            HuduUrl   = $existingArticleUrl
            ArticleId = $existingArticleId
        })
    }

    Write-SharePointExistingArticleSkipState `
        -Doc $file `
        -ExistingArticle $existingArticle `
        -Message "Skipped before conversion because matching Hudu article already exists"

    if ($file.LocalPath -and (Test-Path -LiteralPath $file.LocalPath -PathType Leaf)) {
        try {
            Remove-Item -LiteralPath $file.LocalPath -Force -ErrorAction Stop
        } catch {
            Set-PrintAndLog -message "Failed to remove early-skipped local file '$($file.LocalPath)': $($_.Exception.Message)" -Color DarkYellow
        }
    }

    $earlySkippedCount++
}

$AllDiscoveredFiles = $remainingFiles
if ($earlySkippedCount -gt 0) {
    Set-PrintAndLog -message "Early existing-article skip complete: $earlySkippedCount file(s) removed before conversion; $($AllDiscoveredFiles.Count) file(s) remain in this batch." -Color Green
}
