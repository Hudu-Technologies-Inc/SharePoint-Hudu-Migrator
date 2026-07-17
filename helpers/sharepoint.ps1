if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'sharepoint\manifests.ps1')) {
    . (Join-Path $PSScriptRoot 'sharepoint\manifests.ps1')
}

function Invoke-SharePointGraphCollection {
    param (
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $items = [System.Collections.ArrayList]@()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-RestMethod -Headers $SharePointHeaders -Uri $nextUri -Method Get
        if ($response.value) {
            [void]$items.AddRange(@($response.value))
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return @($items)
}

function Get-SharePointSafePathName {
    param (
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return "unnamed"
    }

    return (($Name -replace '[\\/:*?"<>|]', '_') -replace '\s{2,}', ' ').Trim()
}

function Get-GraphSiteDrives {
    param (
        [Parameter(Mandatory)] [string]$siteId
    )

    Invoke-SharePointGraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
}

function Get-GraphDriveChildItems {
    param (
        [Parameter(Mandatory)] [string]$siteId,
        [Parameter(Mandatory)] [string]$driveId,
        [string]$folderId = 'root'
    )

    Invoke-SharePointGraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives/$driveId/items/$folderId/children"
}

function ConvertTo-SharePointDiscoveredFile {
    param (
        [Parameter(Mandatory)] $Item,
        [Parameter(Mandatory)] [string]$SiteId,
        [Parameter(Mandatory)] [string]$SiteName,
        [Parameter(Mandatory)] [string]$DriveId,
        [string]$DriveName,
        [Parameter(Mandatory)] [string]$FolderId,
        [Parameter(Mandatory)] [string]$ItemPath,
        [int64]$ItemSize,
        [bool]$DownloadSkipped
    )

    $relativePath = Split-Path -Path $ItemPath -Parent
    $relativePath = $relativePath.Substring($allSitesfolder.Length).TrimStart('\')
    $originalLinks = @(
        $Item.webUrl,
        $Item.webDavUrl,
        "$($Item.sharepointIds.siteUrl)/_layouts/15/Doc.aspx?sourcedoc={$($Item.sharepointIds.listId)}&file=$($Item.name)&action=default",
        "$($Item.sharepointIds.siteUrl)/_layouts/15/download.aspx?UniqueId={$($Item.id)}"
    ) | Where-Object { $_ -and ($_ -notmatch 'null') -and ($_ -notmatch '') }

    [PSCustomObject]@{
        Name                = $Item.name
        LocalPath           = $ItemPath
        SiteId              = $SiteId
        SiteName            = $SiteName
        DriveId             = $DriveId
        DriveName           = $DriveName
        FolderId            = $FolderId
        DownloadUrl         = $Item."@microsoft.graph.downloadUrl"
        DownloadSkipped     = $DownloadSkipped
        webViewUrl          = $Item.webUrl
        webDAVUrl           = $Item.webDavUrl
        CreatedDateTime     = $Item.createdDateTime
        LastModifiedDateTime= $Item.lastModifiedDateTime
        sharepointSiteUrl   = $Item.sharepointIds.siteUrl
        sharepointListId    = $Item.sharepointIds.listId
        sharepointItemId    = $Item.sharepointIds.listItemId
        parentDrivePath     = $Item.parentReference.path
        HuduFolder          = $null
        HuduFolderId        = $null
        HuduArticle         = $null
        HuduFolderUUID      = $([guid]::NewGuid().ToString())
        companyID           = $null
        RawContent          = $null
        OriginalFilename    = $Item.name
        ReplacedContent     = $null
        OriginalLinks       = $originalLinks
        Stub                = $null
        ReplacedLinks       = $null
        Links               = $null
        UploadedFiles       = [System.Collections.ArrayList]@()
        ContentPreview      = ""
        UsingGeneratedHTML  = $false
        CharsTrimmed        = 0
        title               = $(Get-SafeTitle -name $Item.name)
        Id                  = $Item.id
        RelativePath        = $relativePath
        Filesize            = $ItemSize
        FileTooLarge        = ($ItemSize -ge 100MB)
    }
}

function Download-GraphDriveItemRecursively {
    param (
        [Parameter(Mandatory)] $item,
        [string]$siteId,
        [string]$siteName,
        [string]$driveId,
        [string]$driveName,
        [string]$localPath
    )

    $discoveredFiles = [System.Collections.ArrayList]@()
    $safeItemName = Get-SharePointSafePathName -Name $item.name
    $itemPath = Join-Path $localPath $safeItemName

    if ($item.folder) {
        if (!(Test-Path $itemPath)) { New-Item -Path $itemPath -ItemType Directory | Out-Null }

        $childItems = Get-GraphDriveChildItems -siteId $siteId -driveId $driveId -folderId $item.id
        foreach ($childItem in $childItems) {
            $childFiles = Download-GraphDriveItemRecursively `
                -item $childItem `
                -siteId $siteId `
                -siteName $siteName `
                -driveId $driveId `
                -driveName $driveName `
                -localPath $itemPath

            if ($null -ne $childFiles -and $childFiles.Count -gt 0) {
                [void]$discoveredFiles.AddRange(@($childFiles))
            }
        }
    }
    elseif ($item.file) {
        $downloadUrl = $item."@microsoft.graph.downloadUrl"
        $itemSize = [int64]($item.size ?? 0)
        $downloadSkipped = $false

        if ($itemSize -ge 100MB) {
            $downloadSkipped = $true
            Set-PrintAndLog -message "Skipping download for 100 MB or larger file; will link back to SharePoint: $itemPath" -Color Yellow
        } else {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $itemPath -UseBasicParsing
            $itemSize = (Get-Item $itemPath).Length
            Set-PrintAndLog -message "Downloaded: $itemPath" -Color DarkMagenta
        }

        [void]$discoveredFiles.Add(
            (ConvertTo-SharePointDiscoveredFile `
                -Item $item `
                -SiteId $siteId `
                -SiteName $siteName `
                -DriveId $driveId `
                -DriveName $driveName `
                -FolderId ($item.parentReference.id ?? 'root') `
                -ItemPath $itemPath `
                -ItemSize $itemSize `
                -DownloadSkipped $downloadSkipped)
        )
    }

    return $discoveredFiles
}

function Download-GraphDriveItemsRecursively {
    param (
        [string]$siteId,
        [string]$siteName,
        [string]$driveId,
        [string]$driveName,
        [string]$folderId = 'root',
        [string]$localPath
    )

    $discoveredFiles = [System.Collections.ArrayList]@()
    $items = Get-GraphDriveChildItems -siteId $siteId -driveId $driveId -folderId $folderId

    foreach ($item in $items) {
        $itemFiles = Download-GraphDriveItemRecursively `
            -item $item `
            -siteId $siteId `
            -siteName $siteName `
            -driveId $driveId `
            -driveName $driveName `
            -localPath $localPath

        if ($null -ne $itemFiles -and $itemFiles.Count -gt 0) {
            [void]$discoveredFiles.AddRange(@($itemFiles))
        }
    }

    return $discoveredFiles
}
