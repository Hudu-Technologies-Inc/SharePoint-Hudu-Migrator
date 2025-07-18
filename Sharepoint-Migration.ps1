$workdir = $PSScriptRoot

##### Step 1, Initialize
##
#
# 1.1 Hudu Set-up
$HUDU_MAX_DOCSIZE=$HUDU_MAX_DOCSIZE ?? 8500
$HuduBaseUrl= $HuduBaseURL ?? $(read-host "enter hudu URL")
$HuduApiKey= $HuduApiKey ?? $(read-host "enter api key")

# 1.2 Sharepoint Set-up
$tenantId =  ?? $(read-host "enter Microsoft Tenant ID")
$clientId = ?? $(read-host "enter Microsoft App Registration Client ID")

## todo: Optional auto app registration (and deregistration later) step
$scopes = "Sites.Read.All"

# 1.3 Init and vars
$userSelectedSites = [System.Collections.ArrayList]@()
$AllDiscoveredFiles = [System.Collections.ArrayList]@()
$AllDiscoveredFolders = [System.Collections.ArrayList]@()
$Attribution_Options=[System.Collections.ArrayList]@()
$TrackedAttachments = [System.Collections.ArrayList]@()
$AllReplacedLinks =  [System.Collections.ArrayList]@()
$AllFoundLinks =[System.Collections.ArrayList]@()
$AllNewLinks = [System.Collections.ArrayList]@()        
$discoveredFolders = [System.Collections.Generic.HashSet[string]]::new()
$AllFolders = [System.Collections.Generic.HashSet[string]]::new()
$discoveredFiles = [System.Collections.ArrayList]@()


$ImageMap = @{}
$allSites = @()
$AllCompanies = @()
$SingleCompanyChoice=@{}
$FolderMap = @{}
$SharePointToHuduUrlMap = @{}
$Article_Relinking=@{}

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
clear-host

# 1.4 Authenticate to Sharepoint
$tokenResult = $tokenResult ?? $(Get-MsalToken -ClientId $clientId -TenantId $tenantId -DeviceCode -Scopes $scopes)
$accessToken = $accessToken ?? $tokenResult.AccessToken
$SharePointHeaders = @{ Authorization = "Bearer $accessToken" }







##### Step 2 Source and Dest Options
##
#
$RunSummary.CompletedStates += "$($RunSummary.State) finished At $($($(Get-Date) - $RunSummary.SetupInfo.StartedAt).ToString())"
$RunSummary.State="Source Data (Sharepoint) and Destination (Hudu) Options"
if ($NonInteractive) {write-host "Noninteractive-Mode enabled. Proceeding to $($RunSummary.State)" -ForegroundColor Green} else {read-host "Prese Enter to proceed to $($RunSummary.State)"}
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
$RunSummary.CompletedStates += "$($RunSummary.State) finished At $($($(Get-Date) - $RunSummary.SetupInfo.StartedAt).ToString())"
$RunSummary.State="Download From Selection"
if ($NonInteractive) {Set-PrintAndLog "Noninteractive-Mode enabled. Proceeding to $($RunSummary.State)" -ForegroundColor Green} else {read-host "Prese Enter to proceed to $($RunSummary.State)"}
. .\jobs\Get-SourceData.ps1
Set-PrintAndLog -message "Writing out discovered source file data to $($RunSummary.OutputJsonFiles.SelectedFiles)...!" -color DarkMagenta
$AllDiscoveredFiles | ConvertTo-Json -Depth 45 | Out-File "$($RunSummary.OutputJsonFiles.SelectedFiles)"
$AllDiscoveredFolders | ConvertTo-Json -Depth 45 | Out-File "$($RunSummary.OutputJsonFiles.SelectedFolders)"






##### Step 4, Initialize Libreoffice/Poppler and Convert Files
##
#
$RunSummary.CompletedStates += "$($RunSummary.State) finished At $($($(Get-Date) - $RunSummary.SetupInfo.StartedAt).ToString())"
$RunSummary.State="Initialize Libreoffice/Poppler and Convert Files"
if ($NonInteractive) {write-host "Noninteractive-Mode enabled. Proceeding to $($RunSummary.State)" -ForegroundColor Green} else {read-host "Prese Enter to proceed to $($RunSummary.State)"}

Set-PrintAndLog "Checking for Libreoffice and installing if not present. If not presnt, follow the on-screen prompts from the installer with default values and don't select 'Run When Finished' for the last question" -color Green
$sofficePath=$(if ($true -eq $portableLibreOffice) {$(Get-LibrePortable -tmpfolder $tmpfolder)} else {$(Get-LibreMSI -tmpfolder $tmpfolder)})
Stop-LibreOffice
$successConverted=$(ConvertDownloadedFiles -downloadedFiles $AllDiscoveredFiles -sofficePath $sofficePath)
. .\jobs\Read-ConvertedContents.ps1

$RunSummary.CompletedStates += "$($RunSummary.State) finished At $($($(Get-Date) - $RunSummary.SetupInfo.StartedAt).ToString())"
$RunSummary.State="Create Folder Sructure and stub articles"
if ($NonInteractive) {write-host "Noninteractive-Mode enabled. Proceeding to $($RunSummary.State)" -ForegroundColor Green} else {read-host "Prese Enter to proceed to $($RunSummary.State)"}

$StubbedArticles=@()
$docIDX=0
# Stub articles

. .\jobs\Make-ArticleStubs.ps1


$RunSummary.CompletedStates += "$($RunSummary.State) finished At $($($(Get-Date) - $RunSummary.SetupInfo.StartedAt).ToString())"
$RunSummary.State="Images and attachments to Hudu"
if ($NonInteractive) {write-host "Noninteractive-Mode enabled. Proceeding to $($RunSummary.State)" -ForegroundColor Green} else {read-host "Prese Enter to proceed to $($RunSummary.State)"}


. .\jobs\Populate-Articles.ps1

$RunSummary.CompletedStates += "$($RunSummary.State) finished At $($($(Get-Date) - $RunSummary.SetupInfo.StartedAt).ToString())"
$RunSummary.State="Relink Uploaded Images to Hudu"
if ($NonInteractive) {write-host "Noninteractive-Mode enabled. Proceeding to $($RunSummary.State)" -ForegroundColor Green} else {read-host "Prese Enter to proceed to $($RunSummary.State)"}



. .\jobs\Upload-Images.ps1

$RunSummary.CompletedStates += "$($RunSummary.State) finished At $($($(Get-Date) - $RunSummary.SetupInfo.StartedAt).ToString())"
$RunSummary.State="Populate Articles In Hudu"
if ($NonInteractive) {write-host "Noninteractive-Mode enabled. Proceeding to $($RunSummary.State)" -ForegroundColor Green} else {read-host "Prese Enter to proceed to $($RunSummary.State)"}



. .\jobs\Relink-Articles.ps1