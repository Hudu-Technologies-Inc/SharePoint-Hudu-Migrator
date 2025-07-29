# hudu vars
$HUDU_MAX_DOCSIZE=$HUDU_MAX_DOCSIZE ?? 8500
$USE_HUDUFORK = $true

# General Vars
$NonInteractive=$false
$AvailableB64MimeTypes = 'image|application|text'
# $AvailableB64MimeTypes="image|application|audio|video|text"

# Libre Set-Up
$portableLibreOffice=$false
$LibreFullInstall="https://www.libreoffice.org/donate/dl/win-x86_64/25.2.4/en-US/LibreOffice_25.2.4_Win_x86-64.msi"
$LibrePortaInstall="https://download.documentfoundation.org/libreoffice/portable/25.2.3/LibreOfficePortable_25.2.3_MultilingualStandard.paf.exe"

# Poppler Setup
$includeHiddenText=$true
$includeComplexLayouts=$true
$PopplerBins=$(join-path $workdir "tools\poppler")
$PDFToHTML=$(join-path $PopplerBins "pdftohtml.exe")


# Define and set up some paths
$logsFolder=$(join-path "$workdir" "logs")
$logFile=$(join-path "$logsFolder" "SharepointLog")
$downloadsFolder=$(join-path "$workdir" "downloads")
$allSitesfolder=$(join-path "$workdir" "sites")
$tmpfolder=$(join-path "$workdir" "tmp")
$ErroredItemsFolder=$(join-path "$logsFolder" "errored")
Write-Host "Hudu Max Docsize: $HUDU_MAX_DOCSIZE"

# Sharepoint-Specific Vars
# columns that are ignored in sharepoint lists
$BlockedSPInternalColumns=@(
    "Folder Child Count","Item Child Count","Comment count",
    "Check In Comment","Retention label","Compliance Asset Id","Label applied by",
    "Like count","Source Version (Converted Document)","Source Version","Modified By",
    "Label setting","Source Name (Converted Document)","Source Name","Copy Source",
    "Item is a Record","App Modified By","App Created By"
)
# lists that are ignored in sharepoint sites
$BlockedSPInternalLists = @(
  "AppPages", "Channel Settings", "ContentTypeAppLog", "ContentTypeSyncLog",
  "CSPViolationReportList", "EnterpriseContentTypesUsage", "Hub Settings", "Web Template Extensions",
  "PackageList", "PackagesMetaInfoList", "Shared Documents", "pImg", "pPg", "pSet", "pSiteList", "pVid"
)

# Migration-related vars
# all SP sites for user to choose to migrate from 
$allSites = @()

# All Files and Folders discovered on all sites 'documents' section
$AllDiscoveredFiles = [System.Collections.ArrayList]@()
$AllDiscoveredFolders = [System.Collections.ArrayList]@()

# company ID, Global KB or use-selector when creating asset layouts, procedures, articles
$Attribution_Options=[System.Collections.ArrayList]@()

# Sites that user selected from all available
$userSelectedSites = [System.Collections.ArrayList]@()

# Sharepoint Lists that we found in user-selected sites
$DiscoveredLists = [System.Collections.ArrayList]@()
$ListsCreated = [System.Collections.ArrayList]@()
$AssetsCreated = [System.Collections.ArrayList]@()

# flattened list of files/folders we found in lists that were in user-selected Sites
$discoveredFiles = [System.Collections.ArrayList]@()

# Links to objects in Hudu that we created
$AllNewLinks = [System.Collections.ArrayList]@()

# Map of images/uploads filename before/after for relinking
$ImageMap = @{}

# asset layouts we created from parsed sharepoint lists found on user-selected sites
$LayoutsCreated = @()

# procedures we created from parsed sharepoint lists found on user-selected sites
$ProceduresCreated = @()

# Article/Upload/Photo -> Asset (SP list-entry) relations to resolve after article-stubbing and uploads processing
$RelationsToResolve = @()

# all companies in hudu to migrate lists and documents to
$AllCompanies = @()

# single company as target company if chosen by user
$SingleCompanyChoice=@{}

# articles that were processed, converted, determined to be fit, given a folder, then stubbed
$StubbedArticles=@()

# Base Asset Fields when creating asset layout from list
$BaseSPLayoutFields = @(@{
        label        = 'Imported from SharePoint'
        field_type   = 'Text'
        show_in_list = 'false'
        position     = 500
    },
    @{
        label        = 'SharePoint URL'
        field_type   = 'Text'
        show_in_list = 'false'
        position     = 501
    },
    @{
        label        = 'Sharepoint ID'
        field_type   = 'Text'
        show_in_list = 'false'
        position     = 502})

$EmbeddableImageExtensions = @(
    ".jpg", ".jpeg",  # JPEG
    ".png",           # Portable Network Graphics
    ".gif",           # GIF (including animated)
    ".bmp",           # Bitmap (support varies by browser)
    ".webp",          # WebP (modern, compressed)
    ".svg",           # Scalable Vector Graphics
    ".apng",          # Animated PNG (limited support)
    ".avif",          # AV1 Image File Format (modern)
    ".ico",           # Icon files (used in favicons)
    ".jfif",          # JPEG File Interchange Format
    ".pjpeg",         # Progressive JPEG
    ".pjp"            # Alternative JPEG extension
)

foreach ($folder in @($logsFolder, $downloadsFolder, $tmpfolder, $allSitesfolder, $ErroredItemsFolder)) {
    if (!(Test-Path -Path "$folder")) { Set-PrintAndLog -message  "Making dir... $(New-Item "$folder" -ItemType Directory)" -Color DarkCyan }
        Set-PrintAndLog -message "Clearing $folder... $(Get-ChildItem -Path "$folder" -File -Recurse -Force | Remove-Item -Force)" -Color DarkCyan
}

# Set up logging object
$RunSummary=@{
    State="Set-Up"
    CompletedStates=@()
    OutputJsonFiles = @{
        SelectedSites    =   "$(join-path $logsFolder -ChildPath "sites.json")"
        SelectedFiles    =   "$(join-path $logsFolder -ChildPath "files.json")"
        SelectedFolders  =   "$(join-path $logsFolder -ChildPath "folders.json")"
        ConvertedFiles   =   "$(join-path $logsFolder -ChildPath "converted.json")"
        SummaryPath      =   "$(join-path $logsFolder -ChildPath "job-summary.json")"
        ListsPath        =   "$(join-path $logsFolder -ChildPath "lists.json")"
    }
    SetupInfo=@{
        HuduDestination     = $HuduBaseUrl
        HuduMaxContentLength= 4500
        SharepointSource    = $SharepointBaseUrl
        HuduVersion         = [version]$HuduAppInfo.version
        PowershellVersion   = [version]$PowershellVersion
        project_workdir     = $project_workdir
        includeSPLists      = $false
        SPListsAsLayouts    = $false
        StartedAt           = $(get-date)
        FinishedAt          = $null
        RunDuration         = $null
        PreviewLength       = 2500
        DisallowedForConvert = [System.Collections.ArrayList]@(
            ".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a",
            ".dll", ".so", ".lib", ".bin", ".class", ".pyc", ".pyo", ".o", ".obj",
            ".exe", ".msi", ".bat", ".cmd", ".sh", ".jar", ".app", ".apk", ".dmg", ".iso", ".img",
            ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".lz",
            ".mp4", ".avi", ".mov", ".wmv", ".mkv", ".webm", ".flv",
            ".psd", ".ai", ".eps", ".indd", ".sketch", ".fig", ".xd", ".blend",
            ".ds_store", ".thumbs", ".lnk", ".heic"
        )
        LinkSourceArticles  = $true
        SourceFilesAsAttachments = $true
    }
    JobInfo=@{
        MigrationSource     = [PSCustomObject]@{}
        MigrationDest       = [PSCustomObject]@{}
        sites               = [System.Collections.ArrayList]@()
        pages               = [System.Collections.ArrayList]@()
        sitescount          = 0
        pagescount          = 0
        LinksCreated        = 0
        LinksFound          = 0
        LinksReplaced       = 0
        ArticlesCreated     = 0
        ArticlesSkipped     = 0
        ArticlesErrored     = 0
        AttachmentsFound    = 0
        UploadsCreated      = 0
        UploadsErrored      = 0
    }
    Errors                  = [System.Collections.ArrayList]@()
    Warnings                = [System.Collections.ArrayList]@()
}

function Set-IncrementedState {
    param (
        [string]$newstate,
        [bool]$pausable=$false
    )
    $RunSummary.CompletedStates += "$($RunSummary.State) finished At $($($(Get-Date) - $RunSummary.SetupInfo.StartedAt).ToString())"
    $RunSummary.State="$newstate"
    if ($pausable){
        if ($NonInteractive) {Set-PrintAndLog -message  "Noninteractive-Mode enabled. Proceeding to $($RunSummary.State)" -Color Green} else {Set-PrintandLog -message "Next Step: $($RunSummary.State)"; read-host "Prese Enter to proceed."}
    }
}

function Set-HuduInstance {
    $HuduBaseURL = $HuduBaseURL ?? 
        $((Read-Host -Prompt 'Set the base domain of your Hudu instance (e.g https://myinstance.huducloud.com)') -replace '[\\/]+$', '') -replace '^(?!https://)', 'https://'
    $HuduAPIKey = $HuduAPIKey ?? "$(read-host "Please Enter Hudu API Key")"
    while ($HuduAPIKey.Length -ne 24) {
        $HuduAPIKey = (Read-Host -Prompt "Get a Hudu API Key from $($settings.HuduBaseDomain)/admin/api_keys").Trim()
        if ($HuduAPIKey.Length -ne 24) {
            Write-Host "This doesn't seem to be a valid Hudu API key. It is $($HuduAPIKey.Length) characters long, but should be 24." -ForegroundColor Red
        }
    }
    New-HuduAPIKey $HuduAPIKey
    New-HuduBaseURL $HuduBaseURL
}

function Get-HuduModule {
    param (
        [string]$HAPImodulePath = "C:\Users\$env:USERNAME\Documents\GitHub\HuduAPI\HuduAPI\HuduAPI.psm1",
        [bool]$use_hudu_fork = $true
        )

    if ($true -eq $use_hudu_fork) {
        if (-not $(Test-Path $HAPImodulePath)) {
            $dst = Split-Path -Path (Split-Path -Path $HAPImodulePath -Parent) -Parent
            Write-Host "Using Lastest Master Branch of Hudu Fork for HuduAPI"
            $zip = "$env:TEMP\huduapi.zip"
            Invoke-WebRequest -Uri "https://github.com/Hudu-Technologies-Inc/HuduAPI/archive/refs/heads/master.zip" -OutFile $zip
            Expand-Archive -Path $zip -DestinationPath $env:TEMP -Force 
            $extracted = Join-Path $env:TEMP "HuduAPI-master" 
            if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
            Move-Item -Path $extracted -Destination $dst 
            Remove-Item $zip -Force
        }
    } else {
        Write-Host "Assuming PSGallery Module if not already locally cloned at $HAPImodulePath"
    }

    if (Test-Path $HAPImodulePath) {
        Import-Module $HAPImodulePath -Force
        Write-Host "Module imported from $HAPImodulePath"
    } elseif ((Get-Module -ListAvailable -Name HuduAPI).Version -ge [version]'2.4.4') {
        Import-Module HuduAPI
        Write-Host "Module 'HuduAPI' imported from global/module path"
    } else {
        Install-Module HuduAPI -MinimumVersion 2.4.5 -Scope CurrentUser -Force
        Import-Module HuduAPI
        Write-Host "Installed and imported HuduAPI from PSGallery"
    }
}
function Get-HuduVersionCompatible {
    param (
        [version]$RequiredHuduVersion = [version]"2.37.1",
        $DisallowedVersions = @([version]"2.37.0")
    )

    Write-Host "Required Hudu version: $RequiredHuduVersion" -ForegroundColor Blue
    try {
        $HuduAppInfo = Get-HuduAppInfo
        $CurrentHuduVersion = [version]$HuduAppInfo.version

        if ($DisallowedVersions -contains $CurrentHuduVersion) {
            Write-Host "Hudu version $CurrentHuduVersion is explicitly disallowed." -ForegroundColor Red
            return $false
        }

        if ($CurrentHuduVersion -lt $RequiredHuduVersion) {
            Write-Host "This script requires at least version $RequiredHuduVersion and cannot run with version $CurrentHuduVersion. Please update your version of Hudu." -ForegroundColor Red
            return $false
        }

        Write-Host "Hudu Version $CurrentHuduVersion is compatible" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Error encountered when checking Hudu version for $(Get-HuduBaseURL): $_" -ForegroundColor Yellow
        return $false
    }
}

function Get-PSVersionCompatible {
    param (
        [version]$RequiredPSversion = [version]"7.5.1"
    )

    $currentPSVersion = (Get-Host).Version
    Write-Host "Required PowerShell version: $RequiredPSversion" -ForegroundColor Blue

    if ($currentPSVersion -lt $RequiredPSversion) {
        Write-Host "PowerShell $RequiredPSversion or higher is required. You have $currentPSVersion." -ForegroundColor Red
        exit 1
    } else {
        Write-Host "PowerShell version $currentPSVersion is compatible." -ForegroundColor Green
    }
}

