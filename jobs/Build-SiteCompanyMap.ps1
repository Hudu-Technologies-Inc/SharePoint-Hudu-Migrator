##### Step 2C, Build SharePoint site to Hudu company map

$SiteCompanyMap = @()

if ([int]$RunSummary.JobInfo.MigrationDest.Identifier -ne 3) {
    Set-PrintAndLog -message "Per-site company destination mode not selected; skipping site company map." -Color DarkGray
    return
}

if (
    $RunSummary.SetupInfo.SiteCompanyUseCachedMap -and
    -not $RunSummary.SetupInfo.SiteCompanyForceRebuildMap -and
    (Test-Path -LiteralPath $RunSummary.OutputJsonFiles.SiteCompanyMap -PathType Leaf)
) {
    try {
        $SiteCompanyMap = @(Get-Content -LiteralPath $RunSummary.OutputJsonFiles.SiteCompanyMap -Raw | ConvertFrom-Json)
        Set-PrintAndLog -message "Loaded cached SharePoint site company map: $($SiteCompanyMap.Count) item(s) from $($RunSummary.OutputJsonFiles.SiteCompanyMap)" -Color Cyan
        return
    } catch {
        Set-PrintAndLog -message "Failed to load cached SharePoint site company map; rebuilding. $($_.Exception.Message)" -Color Yellow
        $SiteCompanyMap = @()
    }
}

Set-PrintAndLog -message "Building SharePoint site to Hudu company map. Missing companies will $(if ($RunSummary.SetupInfo.SiteCompanyCreateMissing) { 'be created' } else { 'not be created' })." -Color Cyan

$siteCompanyMapItems = [System.Collections.Generic.List[object]]::new()

foreach ($site in @($userSelectedSites)) {
    $siteCompanyName = ($site.displayName ?? $site.name)
    $candidates = @(Get-HuduCompanySiteCandidates -Site $site -Companies $AllCompanies)
    $best = $candidates | Select-Object -First 1
    $second = $candidates | Select-Object -Skip 1 -First 1
    $gap = if ($second) { [double]$best.Score - [double]$second.Score } else { [double]$best.Score }
    $matched = ($best -and [double]$best.Score -ge $RunSummary.SetupInfo.SiteCompanyMinScore -and $gap -ge $RunSummary.SetupInfo.SiteCompanyMinGap)
    $company = $null
    $status = 'Unmatched'
    $reason = $best.Reason

    if ($matched) {
        $company = $AllCompanies | Where-Object { $_.Id -eq $best.CompanyId } | Select-Object -First 1
        $status = 'Matched'
        Set-PrintAndLog -message "Matched SharePoint site '$siteCompanyName' to Hudu company '$($company.Name)' ($($best.Score)%, gap $gap)." -Color Cyan
    }
    elseif ($RunSummary.SetupInfo.SiteCompanyCreateMissing) {
        Set-PrintAndLog -message "No confident Hudu company match for SharePoint site '$siteCompanyName'; creating company." -Color Yellow
        try {
            $created = New-HuduCompany -Name $siteCompanyName
            $company = $created.company ?? $created
            $status = 'Created'
            $reason = 'created_missing_company'

            if ($company -and $company.Id) {
                $AllCompanies += $company
                Set-PrintAndLog -message "Created Hudu company '$($company.Name)' with ID $($company.Id)." -Color Green
            } else {
                throw "New-HuduCompany did not return a company id for '$siteCompanyName'."
            }
        } catch {
            $status = 'CreateFailed'
            $RunSummary.Errors.Add(@{
                Step = 'Build site company map'
                Site = $siteCompanyName
                Error = $_.Exception.Message
            })
            Set-PrintAndLog -message "Failed to create Hudu company for SharePoint site '$siteCompanyName': $($_.Exception.Message)" -Color Red
        }
    }

    $siteCompanyMapItems.Add([PSCustomObject]@{
        SiteId            = $site.id
        SiteName          = $siteCompanyName
        SiteSlug          = $site.name
        SiteWebUrl        = $site.webUrl
        HuduCompanyId     = $company.Id
        HuduCompanyName   = $company.Name
        Status            = $status
        Confidence        = if ($best) { [double]$best.Score } else { 0 }
        ConfidenceGap     = [double]$gap
        MatchReason       = $reason
        TopCandidates     = @($candidates | Select-Object -First 5)
    })
}

$SiteCompanyMap = @($siteCompanyMapItems)

$SiteCompanyMap |
    ConvertTo-Json -Depth 20 |
    Out-File -FilePath $RunSummary.OutputJsonFiles.SiteCompanyMap -Encoding UTF8

$SiteCompanyMap |
    ForEach-Object {
        [PSCustomObject]@{
            Status          = $_.Status
            SiteName        = $_.SiteName
            SiteSlug        = $_.SiteSlug
            SiteWebUrl      = $_.SiteWebUrl
            HuduCompanyId   = $_.HuduCompanyId
            HuduCompanyName = $_.HuduCompanyName
            Confidence      = $_.Confidence
            ConfidenceGap   = $_.ConfidenceGap
            MatchReason     = $_.MatchReason
            Candidate2      = @($_.TopCandidates | Select-Object -Skip 1 -First 1).CompanyName
            Candidate2Score = @($_.TopCandidates | Select-Object -Skip 1 -First 1).Score
            Candidate3      = @($_.TopCandidates | Select-Object -Skip 2 -First 1).CompanyName
            Candidate3Score = @($_.TopCandidates | Select-Object -Skip 2 -First 1).Score
        }
    } |
    Export-Csv -Path $RunSummary.OutputJsonFiles.SiteCompanyReview -NoTypeInformation -Encoding UTF8

Set-PrintAndLog -message "Site company map complete: $(@($SiteCompanyMap | Where-Object { $_.Status -eq 'Matched' }).Count) matched, $(@($SiteCompanyMap | Where-Object { $_.Status -eq 'Created' }).Count) created, $(@($SiteCompanyMap | Where-Object { $_.Status -eq 'CreateFailed' -or $_.Status -eq 'Unmatched' }).Count) unresolved." -Color Cyan
Set-PrintAndLog -message "Wrote site company map: $($RunSummary.OutputJsonFiles.SiteCompanyMap)" -Color DarkMagenta
Set-PrintAndLog -message "Wrote site company review CSV: $($RunSummary.OutputJsonFiles.SiteCompanyReview)" -Color DarkMagenta
