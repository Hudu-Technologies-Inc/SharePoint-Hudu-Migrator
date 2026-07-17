##### Step 2A Select Source Options
function Get-SharePointSourceOptionSiteKey {
    param ($Site)

    foreach ($value in @($Site.id, $Site.webUrl, $Site.name, $Site.displayName)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            return [string]$value
        }
    }

    return [guid]::NewGuid().ToString()
}

$graphSites = @(Invoke-SharePointGraphCollection -Uri "https://graph.microsoft.com/v1.0/sites?search=*")
$manifestSites = @()

if ($null -ne $manifestSet -and $null -ne $manifestSet.Manifests) {
    $manifestSites = @(
        foreach ($manifest in @($manifestSet.Manifests)) {
            foreach ($siteEntry in @($manifest.sites)) {
                if ($siteEntry.metadata) {
                    $siteEntry.metadata
                }
            }
        }
    )
}

if ($manifestSites.Count -gt $graphSites.Count) {
    $allSites = @($manifestSites)
    Set-PrintAndLog -message "Graph site search returned $($graphSites.Count) site(s); manifest contains $($manifestSites.Count). Using manifest site list for source selection." -Color Yellow
} else {
    $siteByKey = [ordered]@{}
    foreach ($site in @($graphSites + $manifestSites)) {
        $siteKey = Get-SharePointSourceOptionSiteKey -Site $site
        if (-not $siteByKey.Contains($siteKey)) {
            $siteByKey[$siteKey] = $site
        }
    }

    $allSites = @($siteByKey.Values)
    Set-PrintAndLog -message "Loaded $($allSites.Count) SharePoint site(s) for source selection." -Color Cyan
}

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
        if ($null -eq $selectedSite) { break }
        [void]$userSelectedSites.Add($selectedSite)
    }
} else {
    if ($null -eq $allSites -or $allSites.count -lt 1) {
        Write-Error "No SharePoint sites were loaded. Please check Graph permissions and manifest generation."
        return
    }

    [void]$userSelectedSites.AddRange($allSites)
    }
    if ($userSelectedSites.count -lt 1){
        Write-Error "No Sites Selected. Please try again." -ForegroundColor Red
}
