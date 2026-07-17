##### Step 2C, Build client attribution map

if (-not $RunSummary.SetupInfo.ClientAttributionEnabled) {
    Set-PrintAndLog -message "Client attribution matching disabled." -Color DarkGray
    $ClientAttributionMap = @()
    return
}

if ($null -eq $manifestSet) {
    Set-PrintAndLog -message "No SharePoint manifest set is available; skipping client attribution matching." -Color Yellow
    $ClientAttributionMap = @()
    return
}

if ($null -eq $AllCompanies -or $AllCompanies.Count -eq 0) {
    Set-PrintAndLog -message "No Hudu companies are available; skipping client attribution matching." -Color Yellow
    $ClientAttributionMap = @()
    return
}

Set-PrintAndLog -message "Building SharePoint client attribution map from list(s): $($RunSummary.SetupInfo.ClientAttributionListNames -join ', ')" -Color Cyan

$ClientAttributionMap = @(
    New-SharePointClientAttributionMap `
        -ManifestSet $manifestSet `
        -Companies $AllCompanies `
        -ListNames $RunSummary.SetupInfo.ClientAttributionListNames `
        -MinScore $RunSummary.SetupInfo.ClientAttributionMinScore `
        -MinGap $RunSummary.SetupInfo.ClientAttributionMinGap
)

$autoCount = @($ClientAttributionMap | Where-Object { $_.AutoMatched }).Count
$reviewCount = @($ClientAttributionMap | Where-Object { -not $_.AutoMatched }).Count

Set-PrintAndLog -message "Client attribution map built: $($ClientAttributionMap.Count) SharePoint client entries; $autoCount auto-match(es); $reviewCount review item(s)." -Color Cyan

$ClientAttributionMap |
    ConvertTo-Json -Depth 20 |
    Out-File -FilePath $RunSummary.OutputJsonFiles.ClientAttributionMap -Encoding UTF8

$ClientAttributionMap |
    ForEach-Object {
        [PSCustomObject]@{
            MatchStatus       = $_.MatchStatus
            AutoMatched       = $_.AutoMatched
            Confidence        = $_.Confidence
            ConfidenceGap     = $_.ConfidenceGap
            SharePointTitle   = $_.RawTitle
            ClientName        = $_.ClientName
            ClientCode        = $_.ClientCode
            Provider          = $_.Provider
            HuduCompanyId     = $_.HuduCompanyId
            HuduCompanyName   = $_.HuduCompanyName
            MatchReason       = $_.MatchReason
            Candidate2        = @($_.TopCandidates | Select-Object -Skip 1 -First 1).CompanyName
            Candidate2Score   = @($_.TopCandidates | Select-Object -Skip 1 -First 1).Score
            Candidate3        = @($_.TopCandidates | Select-Object -Skip 2 -First 1).CompanyName
            Candidate3Score   = @($_.TopCandidates | Select-Object -Skip 2 -First 1).Score
        }
    } |
    Export-Csv -Path $RunSummary.OutputJsonFiles.ClientAttributionReview -NoTypeInformation -Encoding UTF8

Set-PrintAndLog -message "Wrote client attribution map: $($RunSummary.OutputJsonFiles.ClientAttributionMap)" -Color DarkMagenta
Set-PrintAndLog -message "Wrote client attribution review CSV: $($RunSummary.OutputJsonFiles.ClientAttributionReview)" -Color DarkMagenta
