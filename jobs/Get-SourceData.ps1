##### Step 3, Get Source Data from Selection

foreach ($site in $userSelectedSites) {
    Write-Host "`nProcessing site: $($site.name)" -ForegroundColor Yellow

    try {
        $drive = Invoke-RestMethod -Headers $SharePointHeaders -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive" -Method Get
        $safeSiteName = ($site.name -replace '[^\w\-]', '_')
        $localBasePath = Join-Path $allSitesfolder $safeSiteName

        if (!(Test-Path $localBasePath)) {
            New-Item -Path $localBasePath -ItemType Directory | Out-Null
        }

        $siteFiles = Download-GraphDriveItemsRecursively `
            -siteName $site.name `
            -siteId $site.id `
            -driveId $drive.id `
            -localPath $localBasePath

        foreach ($f in $siteFiles) {
            $f | Add-Member -NotePropertyName SiteName -NotePropertyValue $site.name -Force
            $f | Add-Member -NotePropertyName SiteId -NotePropertyValue $site.id -Force
            [void]$AllDiscoveredFiles.Add($f)
        }

    } catch {
        Write-Warning "Failed for site $($site.name): $_"
    }
}

$AllDiscoveredFolders = Get-ChildItem -Path $allSitesfolder -Directory -Recurse | ForEach-Object {
    [PSCustomObject]@{
        Name         = $_.Name
        FullPath     = $_.FullName
        RelativePath = $_.FullName.Substring($allSitesfolder.Length).TrimStart('\')
    }
}