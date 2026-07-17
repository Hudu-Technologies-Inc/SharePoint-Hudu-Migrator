if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'sharepoint\manifests.ps1')) {
    . (Join-Path $PSScriptRoot 'sharepoint\manifests.ps1')
}

function Download-GraphDriveItemsRecursively {
    param (
        [string]$siteId,
        [string]$siteName,
        [string]$driveId,
        [string]$folderId = 'root',
        [string]$localPath
    )

    $uri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/$folderId/children"
    $items = Invoke-RestMethod -Headers $SharePointHeaders -Uri $uri -Method Get

    $discoveredFiles = [System.Collections.ArrayList]@()

    foreach ($item in $items.value) {
        $itemPath = Join-Path $localPath $item.name

        if ($item.folder) {
            if (!(Test-Path $itemPath)) { New-Item -Path $itemPath -ItemType Directory | Out-Null }

            # Recurse and add all returned files to our list
            $childFiles = Download-GraphDriveItemsRecursively -siteId $siteId -siteName $siteName -driveId $driveId -folderId $item.id -localPath $itemPath
            if ($null -eq $childFiles -or $childFiles.count -lt 1) { continue } 
            $discoveredFiles.AddRange($childFiles)
        }
        elseif ($item.file) {
            $downloadUrl = $item."@microsoft.graph.downloadUrl"
            $itemSize = [int64]($item.size ?? 0)
            $fileTooLarge = $itemSize -ge 100MB
            $downloadSkipped = $false

            if ($fileTooLarge) {
                $downloadSkipped = $true
                Set-PrintAndLog -message "Skipping download for 100 MB or larger file; will link back to SharePoint: $itemPath" -Color Yellow
            } else {
                Invoke-WebRequest -Uri $downloadUrl -OutFile $itemPath -UseBasicParsing
                $itemSize = (Get-Item $itemPath).Length
            }

            $relativePath = Split-Path -Path $itemPath -Parent
            $relativePath = $relativePath.Substring($allSitesfolder.Length).TrimStart('\')
            $originalLinks = @(
                $item.webUrl,
                $item.webDavUrl,
                "$($item.sharepointIds.siteUrl)/_layouts/15/Doc.aspx?sourcedoc={$($item.sharepointIds.listId)}&file=$($item.name)&action=default",
                "$($item.sharepointIds.siteUrl)/_layouts/15/download.aspx?UniqueId={$($item.id)}"
            ) | Where-Object { $_ -and ($_ -notmatch 'null')  -and ($_ -notmatch '') }
            $discoveredFiles.Add([PSCustomObject]@{
                Name                = $item.name
                LocalPath           = $itemPath
                SiteId              = $siteId
                SiteName            = $siteName
                DriveId             = $driveId
                FolderId            = $folderId
                DownloadUrl         = $item."@microsoft.graph.downloadUrl"
                DownloadSkipped     = $downloadSkipped
                webViewUrl          = $item.webUrl
                webDAVUrl           = $item.webDavUrl
                CreatedDateTime     = $item.createdDateTime
                LastModifiedDateTime= $item.lastModifiedDateTime
                sharepointSiteUrl   = $item.sharepointIds.siteUrl
                sharepointListId    = $item.sharepointIds.listId
                sharepointItemId    = $item.sharepointIds.listItemId
                parentDrivePath     = $item.parentReference.path                
                HuduFolder          = $null
                HuduFolderId        = $null
                HuduArticle         = $null
                HuduFolderUUID      = $([guid]::NewGuid().ToString())
                companyID           = $null
                RawContent          = $null
                OriginalFilename    = $item.name
                ReplacedContent     = $null
                OriginalLinks       = $originalLinks
                Stub                = $null
                ReplacedLinks       = $null
                Links               = $null
                UploadedFiles       = [System.Collections.ArrayList]@()
                ContentPreview      = ""
                UsingGeneratedHTML  = $false
                CharsTrimmed        = 0                
                title               = $(Get-SafeTitle -name $item.name)
                Id                  = $item.id
                RelativePath        = $relativePath 
                Filesize            = $itemSize
                FileTooLarge        = $fileTooLarge
            })
            if (-not $downloadSkipped) {
                Set-PrintAndLog -message "Downloaded: $itemPath" -Color DarkMagenta
            }
        }
    }
    return $discoveredFiles
}
