function Get-LatestLibreURI {
    [CmdletBinding()]
    param([string]$BaseUri = 'https://mirror.usi.edu/pub/tdf/libreoffice/stable/')

    $index = Invoke-WebRequest -Uri $BaseUri -UseBasicParsing
    $versions = $index.Links | Where-Object { $_.href -match '^\d+\.\d+\.\d+\/$' } |
        ForEach-Object {
            $v = $_.href.TrimEnd('/')
            [pscustomobject]@{
                Version = [version]$v
                Href    = $_.href
            }} | Sort-Object Version

    if (-not $versions) {
        throw "No LibreOffice versions found at $BaseUri"
    }

    $latest = $versions[-1]
    $latestUri = "$BaseUri$($latest.Href)"

    Write-Verbose "Latest version directory: $latestUri"
    $archUri = "$latestUri/win/x86_64/"
    $archIndex = Invoke-WebRequest -Uri $archUri -UseBasicParsing
    $pattern = '^LibreOffice_.*_Win_x86-64\.msi$'
    $msi = $archIndex.Links | Where-Object { $_.href -match $pattern } | Select-Object -First 1
    
    if (-not $msi) {throw "Could not locate any LibreOffice MSI matching $pattern at $archUri"}

    return "$archUri$($msi.href)"
}


function Stop-LibreOffice {
    Get-Process | Where-Object { $_.Name -like "soffice*" } | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Get-LibreMSI {
    param ([string]$tmpfolder)
    if ([string]::IsNullOrEmpty($tmpfolder)) {
        $tmpfolder = [System.IO.Path]::GetTempPath()
    }
    if (Test-Path "C:\Program Files\LibreOffice\program\soffice.exe") {
        return "C:\Program Files\LibreOffice\program\soffice.exe"
    }
    $downloadUrl = $(Get-LatestLibreURI) ?? "https://mirror.usi.edu/pub/tdf/libreoffice/stable/25.8.3/win/x86_64/LibreOffice_25.8.3_Win_x86-64.msi"
    $downloadPath = Join-Path $tmpfolder "LibreOffice.msi"

    Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath

    # Attempt to install
    Start-Process msiexec.exe -ArgumentList "/i `"$downloadPath`"" -Wait

    # Look for default install path
    $sofficePath = "C:\Program Files\LibreOffice\program\soffice.exe"
    if (Test-Path $sofficePath) {
        return $sofficePath
    } else {
        $sofficePath=$(read-host "Sorry, but we couldnt find libreoffice install. What we need is soffice.exe, usually at '$sofficePath'. Please enter the path for this manually now.")
    }
    return $sofficePath
}
function Get-LibrePortable {
    param (
        [string]$tmpfolder
    )

    $downloadUrl = "$LibrePortaInstall"
    $downloadPath = Join-Path $tmpfolder "LibreOfficePortable.paf.exe"
    $extractPath = Join-Path $tmpfolder "LibreOfficePortable"

    if (!(Test-Path $extractPath)) {
        New-Item -ItemType Directory -Path $extractPath | Out-Null
    }

    Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath

    Start-Process -FilePath $downloadPath -ArgumentList "/SILENT", "/NORESTART", "/SUPPRESSMSGBOXES", "/DIR=`"$extractPath`"" -Wait

    $sofficePath = Join-Path $extractPath "App\libreoffice\program\soffice.exe"
    if (Test-Path $sofficePath) {
        return $sofficePath
    } else {
        $sofficePath=$(read-host "Sorry, but we couldnt find your poratable libreoffice install. What we need is soffice.exe, usually at $sofficePath")
        $env:PATH = "$(Split-Path $sofficePath);$env:PATH"
    }
    return $sofficePath
}
