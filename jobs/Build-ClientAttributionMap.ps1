##### Step 2C, Build client attribution map

if (-not $RunSummary.SetupInfo.ClientAttributionEnabled) {
    Set-PrintAndLog -message "Client attribution matching disabled." -Color DarkGray
    $ClientAttributionMap = @()
    $ClientAttributionLookup = $null
    $ClientAttributionResolver = @()
    return
}

if ($null -eq $AllCompanies -or $AllCompanies.Count -eq 0) {
    Set-PrintAndLog -message "No Hudu companies are available; skipping client attribution matching." -Color Yellow
    $ClientAttributionMap = @()
    $ClientAttributionLookup = $null
    $ClientAttributionResolver = @()
    return
}

$clientAttributionEntries = @()
$clientAttributionSource = $null
$clientsPath = $RunSummary.SetupInfo.ClientAttributionClientsPath
$resolvedClientsPath = $null

if (-not [string]::IsNullOrWhiteSpace([string]$clientsPath)) {
    $resolvedClientsPath = if ([System.IO.Path]::IsPathRooted([string]$clientsPath)) {
        [string]$clientsPath
    } else {
        Join-Path $workdir ([string]$clientsPath)
    }
}

$clientMapCacheIsStale = $false
if (
    -not [string]::IsNullOrWhiteSpace([string]$resolvedClientsPath) -and
    (Test-Path -LiteralPath $resolvedClientsPath -PathType Leaf) -and
    (Test-Path -LiteralPath $RunSummary.OutputJsonFiles.ClientAttributionMap -PathType Leaf)
) {
    $clientFileLastWriteUtc = (Get-Item -LiteralPath $resolvedClientsPath).LastWriteTimeUtc
    $mapLastWriteUtc = (Get-Item -LiteralPath $RunSummary.OutputJsonFiles.ClientAttributionMap).LastWriteTimeUtc
    $clientMapCacheIsStale = $clientFileLastWriteUtc -gt $mapLastWriteUtc
}

if (
    $RunSummary.SetupInfo.ClientAttributionUseCachedMap -and
    -not $RunSummary.SetupInfo.ClientAttributionForceRebuildMap -and
    -not $clientMapCacheIsStale -and
    (Test-Path -LiteralPath $RunSummary.OutputJsonFiles.ClientAttributionMap -PathType Leaf)
) {
    try {
        $ClientAttributionMap = @(Get-Content -LiteralPath $RunSummary.OutputJsonFiles.ClientAttributionMap -Raw | ConvertFrom-Json)
        $ClientAttributionLookup = New-SharePointClientAttributionLookup -AttributionMap $ClientAttributionMap
        $ClientAttributionResolver = $ClientAttributionLookup
        Set-PrintAndLog -message "Loaded cached SharePoint client attribution map: $($ClientAttributionMap.Count) item(s) from $($RunSummary.OutputJsonFiles.ClientAttributionMap)" -Color Cyan
        return
    } catch {
        Set-PrintAndLog -message "Failed to load cached SharePoint client attribution map; rebuilding. $($_.Exception.Message)" -Color Yellow
        $ClientAttributionMap = @()
        $ClientAttributionLookup = $null
        $ClientAttributionResolver = @()
    }
}

if ($clientMapCacheIsStale) {
    Set-PrintAndLog -message "Cached SharePoint client attribution map is older than clients.json; rebuilding." -Color Yellow
}

if (-not [string]::IsNullOrWhiteSpace([string]$resolvedClientsPath)) {
    if (Test-Path -LiteralPath $resolvedClientsPath -PathType Leaf) {
        Set-PrintAndLog -message "Loading predetermined SharePoint client list: $resolvedClientsPath" -Color Cyan
        $clientAttributionEntries = @(Import-SharePointClientAttributionClientFile -Path $resolvedClientsPath)
        if ($clientAttributionEntries.Count -gt 0) {
            $clientAttributionSource = "client file: $resolvedClientsPath"
            foreach ($clientAttributionEntry in $clientAttributionEntries) {
                $clientAttributionEntry | Add-Member -MemberType NoteProperty -Name AttributionSource -Value $clientAttributionSource -Force
            }
        } else {
            Set-PrintAndLog -message "Predetermined SharePoint client list was empty; falling back to manifest list(s)." -Color Yellow
        }
    } else {
        Set-PrintAndLog -message "Predetermined SharePoint client list not found at $resolvedClientsPath; falling back to manifest list(s)." -Color DarkGray
    }
}

if ($clientAttributionEntries.Count -gt 0) {
    Set-PrintAndLog -message "Building SharePoint client attribution map from predetermined client list ($($clientAttributionEntries.Count) entries)." -Color Cyan
    $ClientAttributionMap = @(
        New-HuduClientAttributionMapFromEntries `
            -Entries $clientAttributionEntries `
            -Companies $AllCompanies `
            -MinScore $RunSummary.SetupInfo.ClientAttributionMinScore `
            -MinGap $RunSummary.SetupInfo.ClientAttributionMinGap
    )
} else {
    if ($null -eq $manifestSet) {
        Set-PrintAndLog -message "No SharePoint manifest set is available; skipping client attribution matching." -Color Yellow
        $ClientAttributionMap = @()
        $ClientAttributionLookup = $null
        $ClientAttributionResolver = @()
        return
    }

    Set-PrintAndLog -message "Building SharePoint client attribution map from list(s): $($RunSummary.SetupInfo.ClientAttributionListNames -join ', ')" -Color Cyan
    $clientAttributionSource = "manifest list(s): $($RunSummary.SetupInfo.ClientAttributionListNames -join ', ')"
    $ClientAttributionMap = @(
        New-SharePointClientAttributionMap `
            -ManifestSet $manifestSet `
            -Companies $AllCompanies `
            -ListNames $RunSummary.SetupInfo.ClientAttributionListNames `
            -MinScore $RunSummary.SetupInfo.ClientAttributionMinScore `
            -MinGap $RunSummary.SetupInfo.ClientAttributionMinGap
    )
}

foreach ($mapEntry in @($ClientAttributionMap)) {
    if ([string]::IsNullOrWhiteSpace([string]$mapEntry.AttributionSource)) {
        $mapEntry | Add-Member -MemberType NoteProperty -Name AttributionSource -Value $clientAttributionSource -Force
    }
}

$ClientAttributionLookup = New-SharePointClientAttributionLookup -AttributionMap $ClientAttributionMap
$ClientAttributionResolver = $ClientAttributionLookup

$autoCount = @($ClientAttributionMap | Where-Object { $_.AutoMatched }).Count
$reviewCount = @($ClientAttributionMap | Where-Object { -not $_.AutoMatched }).Count

Set-PrintAndLog -message "Client attribution map built from $clientAttributionSource`: $($ClientAttributionMap.Count) client entries; $autoCount auto-match(es); $reviewCount review item(s)." -Color Cyan

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
            AttributionSource = $_.AttributionSource
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
