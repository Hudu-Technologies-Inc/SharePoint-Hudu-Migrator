##### Step 2A Select Source Options
$allSitesResponse = Invoke-RestMethod -Headers $SharePointHeaders -Uri "https://graph.microsoft.com/v1.0/sites?search=*" -Method Get
$allSites = $allSitesResponse.value
foreach ($site in $allSites) {
    $site | Add-Member -NotePropertyName FetchedBy     -NotePropertyValue $localhost_name -Force
    $site | Add-Member -NotePropertyName CompanyId     -NotePropertyValue $null -Force
}

$RunSummary.JobInfo.MigrationSource=$(Select-ObjectFromList -Objects @(
    [PSCustomObject]@{
        OptionMessage=  "From a specific SharePoint Site"
        Identifier = 0
    }, 
    [PSCustomObject]@{
        OptionMessage= "From Some SharePoint Sites"
        Identifier = 1
    },
    [PSCustomObject]@{
        OptionMessage= "From All SharePoint Sites ($($allSites.count) total)"
        Identifier = 2
    }) -message "Configure Source (Sharepoint-Side) Options- Migrate from which Site?" -allowNull $false)
if ($RunSummary.JobInfo.MigrationSource.Identifier -eq 0) {
    $selectedSite = Select-ObjectFromList -objects $allSites -message "From which site?" -allowNull $true
    [void]$userSelectedSites.Add($selectedSite)
} elseif ($RunSummary.JobInfo.MigrationSource.Identifier -eq 1) {
    while ($true) {
    $selectedSite = Select-ObjectFromList -objects $allSites -message "From which site?" -allowNull $true
        if ($null -eq $selection) { break }
        [void]$userSelectedSites.Add($selectedSite)
    }
} else {
    [void]$userSelectedSites.AddRange($allSites)
    }
    if ($userSelectedSites.count -lt 1){
        Write-Error "No Sites Selected. Please try again." -ForegroundColor Red
}
