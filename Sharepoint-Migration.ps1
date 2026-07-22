$workdir = $PSScriptRoot

##### Step 1, Initialize
##
#
# 1.1 Hudu Set-up
$HUDU_MAX_DOCSIZE= 196000
$HuduBaseUrl= $HuduBaseURL ?? $(read-host "enter hudu URL")
$HuduApiKey= $HuduApiKey ?? $(read-host "enter api key")

# 1.2 Sharepoint Set-up- Add these values here if you set up appregistration manually.
$tenantId = $tenantId ?? $null
$clientId = $clientId ?? $null
$scopes =  "Sites.Read.All Files.Read.All User.Read offline_access"

# 1.3 Init and vars
$userSelectedSites = [System.Collections.ArrayList]@()
$AllDiscoveredFiles = [System.Collections.ArrayList]@()
$AllDiscoveredFolders = [System.Collections.ArrayList]@()
$AllProcessedDiscoveredFiles = [System.Collections.ArrayList]@()
$AllProcessedDiscoveredFolders = [System.Collections.ArrayList]@()
$IndexOnlyFiles = [System.Collections.ArrayList]@()
$IndexOnlyArticles = [System.Collections.ArrayList]@()
$Attribution_Options=[System.Collections.ArrayList]@()
$AllNewLinks = [System.Collections.ArrayList]@()        
$discoveredFiles = [System.Collections.ArrayList]@()
$ImageMap = @{}
$allSites = @()
$AllCompanies = @()
$SingleCompanyChoice=@{}
$StubbedArticles=@()
$ClientAttributionMap=@()
$ClientDesignationMap=$null
$SiteCompanyMap=@()
$SharePointMigrationState=@{}

foreach ($file in $(Get-ChildItem -Path ".\helpers" -Filter "*.ps1" -File | Sort-Object Name)) {
    Write-Host "Importing: $($file.Name)" -ForegroundColor DarkBlue
    . $file.FullName
}
foreach ($module in @("MSAL.PS")) {
    write-host "Installing, Updating, Importing module: $module. Please be patient..."  -ForegroundColor DarkBlue;  Update-Module $module -Force;  Install-Module $module -Scope CurrentUser -Force -AllowClobber; Import-Module $module;
}
Set-Content -Path $logFile -Value "Starting Sharepoint Migration" 
Set-PrintAndLog -message "Checked Powershell Version... $(Get-PSVersionCompatible)" -Color DarkBlue
Set-PrintAndLog -message "Imported Hudu Module and authenticated / checked version... $(Set-HuduModuleInitialized -huduBaseurl $HuduBaseURL -huduAPIkey $HuduApiKey)" -Color DarkBlue
$registration = EnsureRegistration -ClientId $clientId -TenantId $tenantId
$clientId = $clientId ?? $registration.clientId
$tenantId = $tenantId ?? $registration.tenantId

# 1.4 Authenticate to Sharepoint
Start-Process "https://microsoft.com/devicelogin"
$tokenResult = $tokenResult ?? $(Get-MsalToken -ClientId $clientId -TenantId $tenantId -DeviceCode -Scopes $scopes)
$accessToken = $accessToken ?? $tokenResult.AccessToken
$SharePointHeaders = Update-SharePointAccessToken

$manifestParams = @{
    ManifestMode = 'Auto'
    ManifestDir  = ".\out\sharepoint-manifests"
    Headers      = $SharePointHeaders
    RefreshHeaders = { Update-SharePointAccessToken }
    FirstSiteOnly = $SharePointManifestFirstSiteOnly ?? $false
}
if ($null -ne $SharePointManifestMaxSites -and $SharePointManifestMaxSites -gt 0) {$manifestParams.MaxSites = $SharePointManifestMaxSites}

$manifestSet = Initialize-SharePointManifestSet @manifestParams

$workItems = @(ConvertFrom-SharePointManifestSet -ManifestSet $manifestSet)

$SharePointMigrationState = if ($RunSummary.SetupInfo.ResumeFromState) {
    Import-SharePointMigrationState -Path $RunSummary.OutputJsonFiles.MigrationState
} else {
    @{}
}
Set-PrintAndLog -message "Loaded SharePoint migration state: $($SharePointMigrationState.Count) completed/skipped/failed state entr$(if ($SharePointMigrationState.Count -eq 1) { 'y' } else { 'ies' }) from $($RunSummary.OutputJsonFiles.MigrationState)" -Color Cyan



##### Step 2 Source and Dest Options
##
#
Set-IncrementedState -newState "Source Data (Sharepoint) and Destination (Hudu) Options"
# 2.1 Select Source Options
. .\jobs\Source-Options.ps1
Set-PrintAndLog -message "$($userSelectedSites.count) Sites selected as source for migration."
Set-PrintAndLog -message "Writing out user-selected sites info to sites.json $($RunSummary.OutputJsonFiles.SelectedSites)...!" -color DarkMagenta
$userSelectedSites | ConvertTo-Json -Depth 45 | Out-File "$($RunSummary.OutputJsonFiles.SelectedSites)"

if ($RunSummary.SetupInfo.FetchSitePages) {
    Set-IncrementedState -newState "Fetch SharePoint Site Pages"
    . .\jobs\Get-SitePages.ps1
}

# 2.2 Select Dest Options
. .\jobs\Dest-Options.ps1

# 2.3 Build optional site-to-company map
. .\jobs\Build-SiteCompanyMap.ps1

# 2.4 Build optional client attribution map
. .\jobs\Build-ClientAttributionMap.ps1

# 2.5 Export configured structured SharePoint lists for later asset import
. .\jobs\Export-StructuredListJson.ps1

if ($RunSummary.SetupInfo.StructuredListJsonOnly) {
    $RunSummary.JobInfo.FinishedAt = Get-Date
    $RunSummary.JobInfo.RunDuration = New-TimeSpan -Start $RunSummary.JobInfo.StartedAt -End $RunSummary.JobInfo.FinishedAt
    Set-PrintAndLog -message "Structured list JSON only mode enabled; stopping before file conversion and article upload." -Color Green
    $RunSummary | ConvertTo-Json -Depth 50 | Out-File -FilePath $RunSummary.OutputJsonFiles.JobSummary -Encoding UTF8
    return
}

##### Step 4, Initialize Libreoffice/Poppler and Convert Files
##
#
Set-IncrementedState -newState "Initialize Libreoffice/Poppler and Convert Files"
Set-PrintAndLog "Checking for Libreoffice and installing if not present. If not presnt, follow the on-screen prompts from the installer with default values and don't select 'Run When Finished' for the last question" -color Green

# Step 4.1 Init Libre / Poppler
$sofficePath=$(if ($true -eq $portableLibreOffice) {$(Get-LibrePortable -tmpfolder $tmpfolder)} else {$(Get-LibreMSI -tmpfolder $tmpfolder)})
Stop-LibreOffice

function Invoke-SharePointMigrationFileBatch {
    param (
        [Parameter(Mandatory)] [array]$Sites,
        [Parameter(Mandatory)] [string]$BatchName,
        [Parameter(Mandatory)] [string]$SofficePath,
        [array]$Drives,
        [array]$RootItems,
        [switch]$CleanupAfterBatch
    )

    $AllDiscoveredFiles = [System.Collections.ArrayList]@()
    $AllDiscoveredFolders = [System.Collections.ArrayList]@()
    $IndexOnlyFiles = [System.Collections.ArrayList]@()
    $IndexOnlyArticles = [System.Collections.ArrayList]@()
    $StubbedArticles = @()
    $successConverted = @()

    $SourceDataSites = [System.Collections.ArrayList]@()
    [void]$SourceDataSites.AddRange(@($Sites))
    $SourceDataDrives = $null
    if ($null -ne $Drives -and $Drives.Count -gt 0) {
        $SourceDataDrives = [System.Collections.ArrayList]@()
        [void]$SourceDataDrives.AddRange(@($Drives))
    }
    $SourceDataRootItems = $null
    if ($null -ne $RootItems -and $RootItems.Count -gt 0) {
        $SourceDataRootItems = [System.Collections.ArrayList]@()
        [void]$SourceDataRootItems.AddRange(@($RootItems))
    }

    Set-IncrementedState -newState "Download From Selection - $BatchName"
    . .\jobs\Get-SourceData.ps1

    Set-IncrementedState -newState "Skip Existing Articles Before Conversion - $BatchName"
    . .\jobs\Skip-ExistingArticlesEarly.ps1

    if ($AllDiscoveredFiles.Count -gt 0) {
        [void]$AllProcessedDiscoveredFiles.AddRange($AllDiscoveredFiles)
    }
    if ($AllDiscoveredFolders.Count -gt 0) {
        [void]$AllProcessedDiscoveredFolders.AddRange($AllDiscoveredFolders)
    }

    Set-PrintAndLog -message "Writing out discovered source file data to $($RunSummary.OutputJsonFiles.SelectedFiles)...!" -color DarkMagenta
    $AllProcessedDiscoveredFiles | ConvertTo-Json -Depth 45 | Out-File "$($RunSummary.OutputJsonFiles.SelectedFiles)"
    $AllProcessedDiscoveredFolders | ConvertTo-Json -Depth 45 | Out-File "$($RunSummary.OutputJsonFiles.SelectedFolders)"

    Set-IncrementedState -newState "Convert Eligible Files - $BatchName"
    $successConverted = @(ConvertDownloadedFiles -downloadedFiles $AllDiscoveredFiles -sofficePath $SofficePath)

    Set-IncrementedState -newState "Read Now-Converted File Contents - $BatchName"
    . .\jobs\Read-ConvertedContents.ps1

    Set-IncrementedState -newState "Create index-only file articles - $BatchName"
    . .\jobs\Make-IndexOnlyArticles.ps1

    Set-IncrementedState -newState "Determine Company Designations and Folder Structure - $BatchName"
    . .\jobs\Make-ArticleStubs.ps1

    Set-IncrementedState -newState "Populate initial data into articles - $BatchName"
    . .\jobs\Populate-Articles.ps1

    Set-IncrementedState -newState "Upload extracted/embedded images / attachments to Hudu - $BatchName"
    . .\jobs\Upload-Images.ps1

    Set-IncrementedState -newState "Relink Articles - $BatchName"
    . .\jobs\Relink-Articles.ps1

    if ($CleanupAfterBatch) {
        foreach ($site in @($Sites)) {
            $safeSiteName = ($site.name -replace '[^\w\-]', '_')
            $siteRootPath = Join-Path $allSitesfolder $safeSiteName
            Clear-SharePointBatchWorkingFiles -Files @(@($AllDiscoveredFiles) + @($successConverted)) -SiteRootPath $siteRootPath -TempPath $tmpfolder
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

if ($RunSummary.SetupInfo.LowDiskMode) {
    Set-PrintAndLog -message "Low-disk mode enabled. Processing and cleaning one top-level SharePoint drive item at a time." -Color Yellow
    $siteIndex = 0
    foreach ($site in $userSelectedSites) {
        $siteIndex++
        try {
            $drives = @(Get-GraphSiteDrives -siteId $site.id)
        } catch {
            Set-PrintAndLog -message "Failed to enumerate drives for site $($site.name): $($_.Exception.Message)" -Color Red
            $RunSummary.Errors.Add(@{
                Site  = $site.name
                Error = $_.Exception.Message
                Step  = "Enumerate site drives"
            })
            continue
        }

        $driveIndex = 0

        foreach ($drive in $drives) {
            $driveIndex++
            try {
                $rootItems = @(Get-GraphDriveChildItems -siteId $site.id -driveId $drive.id -folderId 'root')
            } catch {
                Set-PrintAndLog -message "Failed to enumerate root items for drive $($drive.name) in site $($site.name): $($_.Exception.Message)" -Color Red
                $RunSummary.Errors.Add(@{
                    Site  = $site.name
                    Drive = $drive.name
                    Error = $_.Exception.Message
                    Step  = "Enumerate drive root items"
                })
                continue
            }

            if ($rootItems.Count -eq 0) {
                Set-PrintAndLog -message "Skipping empty drive $($drive.name) for site $($site.name)." -Color DarkGray
                continue
            }

            $rootItemIndex = 0
            foreach ($rootItem in $rootItems) {
                $rootItemIndex++
                Invoke-SharePointMigrationFileBatch `
                    -Sites @($site) `
                    -Drives @($drive) `
                    -RootItems @($rootItem) `
                    -BatchName "site $siteIndex/$($userSelectedSites.Count): $($site.name); drive $driveIndex/$($drives.Count): $($drive.name); item $rootItemIndex/$($rootItems.Count): $($rootItem.name)" `
                    -SofficePath $sofficePath `
                    -CleanupAfterBatch
            }
        }
    }
} else {
    Invoke-SharePointMigrationFileBatch `
        -Sites @($userSelectedSites) `
        -BatchName "all selected sites" `
        -SofficePath $sofficePath
}

##### Step 6, clean up vars, folders, appregistration and generate summary
##
# All set, clean up, and spit the facts, as the kids say.
Set-IncrementedState -newState "Clean Up - AppRegistration"
if ($(Select-ObjectFromList -objects @("yes","no") -message "Would you like to remove the app registration used for this migration?") -eq "yes"){
    Set-PrintAndLog -message "Removing App Registration and Service Principal... $(Remove-AppRegistrationAndSP -AppId $AppId)" -color Magenta
}
Set-IncrementedState -newState "Clean Up - vars"
foreach ($varname in @("tenantId","clientId","scopes","HuduBaseUrl","HuduApiKey","SharePointHeaders","accessToken","tokenResult")) {
    Set-PrintAndLog -message "Removing var $varname... $(remove-variable -name varname -Force -ErrorAction SilentlyContinue)"
}
Set-IncrementedState -newState "Clean Up - files"
foreach ($folder in @($downloadsFolder, $tmpfolder, $allSitesfolder)) {
    Set-PrintAndLog -message "Clearing $folder..." -Color Magenta

    try {
        Get-ChildItem -Path $folder -File -Recurse -Force | Remove-Item -Force -ErrorAction Stop
    } catch {
        Set-PrintAndLog -message "Failed to clear $folder $($_.Exception.Message)" -Color Red
        $RunSummary.Errors += @{
            Folder = $folder
            Error  = $_.Exception.Message
        }
    }
}

Set-IncrementedState -newState "Complete"
Read-Host "Press Enter to Finish and Print Summary (available in )"
$SummaryJson = $RunSummary | ConvertTo-Json -Depth 20
$SummaryJson -split "`n" | ForEach-Object {
    $_ -replace '[\{\[]', '⤵' `
       -replace '[\}\]]', '' `
       -replace '",', '"' `
       -replace '^', '  '
}
$SummaryJson | ConvertTo-Json -Depth 15 | Out-File "$($RunSummary.OutputJsonFiles.SummaryPath)"
Write-Host "$($RunSummary.CompletedStates.Count): $($RunSummary.State) in $($RunSummary.SetupInfo.RunDuration) with $($RunSummary.Errors.Count) errors and $($RunSummary.Warnings.Count) warnings" -ForegroundColor Magenta
