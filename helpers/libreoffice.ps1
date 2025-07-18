
function Stop-LibreOffice {
    Get-Process | Where-Object { $_.Name -like "soffice*" } | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Get-LibreMSI {
    param ([string]$tmpfolder)
    if (Test-Path "C:\Program Files\LibreOffice\program\soffice.exe") {
        return "C:\Program Files\LibreOffice\program\soffice.exe"
    }
    $downloadUrl = "$LibreFullInstall"
    $downloadPath = Join-Path $tmpfolder "LibreOffice.msi"

    Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath

    # Attempt to install
    Start-Process msiexec.exe -ArgumentList "/i `"$downloadPath`" /qn" -Wait

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
