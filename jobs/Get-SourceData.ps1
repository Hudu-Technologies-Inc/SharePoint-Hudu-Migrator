##### Step 3, Get Source Data from Selection
$sourceSites = if ($null -ne $SourceDataSites) { $SourceDataSites } else { $userSelectedSites }
foreach ($site in $sourceSites) {
    Write-Host "`nProcessing site: $($site.name)" -ForegroundColor Yellow

    try {
        $drives = if ($null -ne $SourceDataDrives) {
            @($SourceDataDrives)
        } else {
            @(Invoke-RestMethod -Headers (Update-SharePointAccessToken) -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive" -Method Get)
        }

        $safeSiteName = ($site.name -replace '[^\w\-]', '_')
        $localBasePath = Join-Path $allSitesfolder $safeSiteName

        if (!(Test-Path $localBasePath)) {
            New-Item -Path $localBasePath -ItemType Directory | Out-Null
        }

        foreach ($drive in $drives) {
            $safeDriveName = Get-SharePointSafePathName -Name ($drive.name ?? $drive.id)
            $driveBasePath = if ($null -ne $SourceDataDrives) {
                Join-Path $localBasePath $safeDriveName
            } else {
                $localBasePath
            }

            if (!(Test-Path $driveBasePath)) {
                New-Item -Path $driveBasePath -ItemType Directory | Out-Null
            }

            $siteFiles = [System.Collections.ArrayList]@()
            if ($null -ne $SourceDataRootItems) {
                foreach ($rootItem in @($SourceDataRootItems)) {
                    $rootItemFiles = Download-GraphDriveItemRecursively `
                        -item $rootItem `
                        -siteName $site.name `
                        -siteId $site.id `
                        -driveId $drive.id `
                        -driveName $drive.name `
                        -localPath $driveBasePath

                    if ($null -ne $rootItemFiles -and $rootItemFiles.Count -gt 0) {
                        [void]$siteFiles.AddRange(@($rootItemFiles))
                    }
                }
            } else {
                $siteFiles = Download-GraphDriveItemsRecursively `
                    -siteName $site.name `
                    -siteId $site.id `
                    -driveId $drive.id `
                    -driveName $drive.name `
                    -localPath $driveBasePath
            }

            foreach ($f in $siteFiles) {
                $f | Add-Member -NotePropertyName SiteName -NotePropertyValue $site.name -Force
                $f | Add-Member -NotePropertyName SiteId -NotePropertyValue $site.id -Force
                [void]$AllDiscoveredFiles.Add($f)
            }
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
