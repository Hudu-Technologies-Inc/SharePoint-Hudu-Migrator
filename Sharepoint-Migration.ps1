$workdir = $PSScriptRoot

##### Step 1, Initialize
##
#
# 1.1 Hudu Set-up
$HUDU_MAX_DOCSIZE=$HUDU_MAX_DOCSIZE ?? 8500
$HuduBaseUrl= $HuduBaseURL ?? $(read-host "enter hudu URL")
$HuduApiKey= $HuduApiKey ?? $(read-host "enter api key")

# 1.2 Sharepoint Set-up
$tenantId = $tenantId ?? $null
$clientId = $clientId ?? $null

$scopes =  "Sites.Read.All Files.Read.All User.Read offline_access"

# 1.3 Init and vars
$userSelectedSites = [System.Collections.ArrayList]@()
$AllDiscoveredFiles = [System.Collections.ArrayList]@()
$AllDiscoveredFolders = [System.Collections.ArrayList]@()
$Attribution_Options=[System.Collections.ArrayList]@()
$AllNewLinks = [System.Collections.ArrayList]@()        
$discoveredFiles = [System.Collections.ArrayList]@()
$ImageMap = @{}
$allSites = @()
$AllCompanies = @()
$SingleCompanyChoice=@{}
$StubbedArticles=@()

foreach ($file in $(Get-ChildItem -Path ".\helpers" -Filter "*.ps1" -File | Sort-Object Name)) {
    Write-Host "Importing: $($file.Name)" -ForegroundColor DarkBlue
    . $file.FullName
}
foreach ($module in @("MSAL.PS")) {
    write-host "Installing, Updating, Importing module: $module. Please be patient..."  -ForegroundColor DarkBlue; Install-Module $module -Scope CurrentUser -Force -AllowClobber; Update-Module $module -Force; Import-Module $module
}
Set-Content -Path $logFile -Value "Starting Sharepoint Migration" 
Set-PrintAndLog -message "Checked Powershell Version... $(Get-PSVersionCompatible)" -Color DarkBlue
Set-PrintAndLog -message "Imported Hudu Module... $(Get-HuduModule)" -Color DarkBlue
Set-PrintAndLog -message "Checked Hudu Credentials... $(Set-HuduInstance)" -Color DarkBlue
Set-PrintAndLog -message "Checked Hudu Version... $(Get-HuduVersionCompatible)" -Color DarkBlue
$registration = EnsureRegistration -ClientId $clientId -TenantId $tenantId
$clientId = $clientId ?? $registration.clientId
$tenantId = $tenantId ?? $registration.tenantId

clear-host

# 1.4 Authenticate to Sharepoint
Start-Process "https://microsoft.com/devicelogin"
$tokenResult = $tokenResult ?? $(Get-MsalToken -ClientId $clientId -TenantId $tenantId -DeviceCode -Scopes $scopes)
$accessToken = $accessToken ?? $tokenResult.AccessToken
$SharePointHeaders = @{ Authorization = "Bearer $accessToken" }


##### Step 2 Source and Dest Options
##
#
Set-IncrementedState -newState "Source Data (Sharepoint) and Destination (Hudu) Options"
# 2.1 Select Source Options
. .\jobs\Source-Options.ps1
Set-PrintAndLog -message "$($userSelectedSites.count) Sites selected as source for migration."
Set-PrintAndLog -message "Writing out user-selected sites info to sites.json $($RunSummary.OutputJsonFiles.SelectedSites)...!" -color DarkMagenta
$userSelectedSites | ConvertTo-Json -Depth 45 | Out-File "$($RunSummary.OutputJsonFiles.SelectedSites)"

# 2.2 Select Dest Options
. .\jobs\Dest-Options.ps1

##### Step 3, Get Source Data from Selection
##
#
Set-IncrementedState -newState "Download From Selection"
. .\jobs\Get-SourceData.ps1
Set-PrintAndLog -message "Writing out discovered source file data to $($RunSummary.OutputJsonFiles.SelectedFiles)...!" -color DarkMagenta
$AllDiscoveredFiles | ConvertTo-Json -Depth 45 | Out-File "$($RunSummary.OutputJsonFiles.SelectedFiles)"
$AllDiscoveredFolders | ConvertTo-Json -Depth 45 | Out-File "$($RunSummary.OutputJsonFiles.SelectedFolders)"

##### Step 4, Initialize Libreoffice/Poppler and Convert Files
##
#
Set-IncrementedState -newState "Initialize Libreoffice/Poppler and Convert Files"
Set-PrintAndLog "Checking for Libreoffice and installing if not present. If not presnt, follow the on-screen prompts from the installer with default values and don't select 'Run When Finished' for the last question" -color Green

# Step 4.1 Init Libre / Poppler
$sofficePath=$(if ($true -eq $portableLibreOffice) {$(Get-LibrePortable -tmpfolder $tmpfolder)} else {$(Get-LibreMSI -tmpfolder $tmpfolder)})
Stop-LibreOffice

# Step 4.2 Convert Files
Set-IncrementedState -newState "Convert Eligible Files"
$successConverted=$(ConvertDownloadedFiles -downloadedFiles $AllDiscoveredFiles -sofficePath $sofficePath)

Set-IncrementedState -newState "Read Now-Converted File Contents"
. .\jobs\Read-ConvertedContents.ps1

##### Step 5, create articles, uploads, folders, then relink articles
##
#
Set-IncrementedState -newState "Determine Company Designations and Folder Structure"
. .\jobs\Make-ArticleStubs.ps1

Set-IncrementedState -newState "Populate initial data into articles"
. .\jobs\Populate-Articles.ps1

Set-IncrementedState -newState "Upload extracted/embedded images / attachments to Hudu"
. .\jobs\Upload-Images.ps1

Set-IncrementedState -newState "Relink Articles"
. .\jobs\Relink-Articles.ps1

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
if ($(Select-ObjectFromList -objects @("yes","no") -message "Would you like to clean up temp files? (not including logs)") -eq "yes"){
    foreach ($folder in @($downloadsFolder, $tmpfolder, $allSitesfolder)) {Set-PrintAndLog -message "Clearing $folder... $(Get-ChildItem -Path "$folder" -File -Recurse -Force | Remove-Item -Force)" -color Magenta}
}

Set-IncrementedState -newState "Complete"
Read-Host "Press Enter to Finish and Print Summary (available in )"
$SummaryJson = $RunSummary | ConvertTo-Json -Depth 20
$SummaryJson -split "`n" | ForEach-Object {
    $_ -replace '[\{\[]', '⤵' `
       -replace '[\}\]]', '' `
       -replace '",', '"' `
       -replace '^', '  '
}$SummaryJson | ConvertTo-Json -Depth 15 | Out-File "$($RunSummary.OutputJsonFiles.SummaryPath)"
Write-Host "$($RunSummary.CompletedStates.Count): $($RunSummary.State) in $($RunSummary.SetupInfo.RunDuration) with $($RunSummary.Errors.Count) errors and $($RunSummary.Warnings.Count) warnings" -ForegroundColor Magenta