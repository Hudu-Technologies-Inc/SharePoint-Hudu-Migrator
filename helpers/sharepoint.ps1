function Get-SPColumnType {
    param ([pscustomobject]$col)

    foreach ($type in 'text','note','number','choice','multichoice','boolean','dateTime','currency','url','lookup','user','calculated','taxonomy') {
        if ($col.PSObject.Properties.Name -contains $type -and $col.$type) {
            return $type
        }
    }

    return 'text'  # Default fallback
}
function Get-SPListItemTypeToHuduALType {
param (
    [string]$SPListItemType,
    [string]$FieldName,
    [array]$SampleValues
)
    $SharePointToHuduMap = @{
        "text"         = "text"
        "note"         = "richtext"
        "number"       = "number"
        "currency"     = "text"
        "boolean"      = "checkbox"
        "datetime"     = "date"
        "choice"       = "dropdown"
        "multichoice"  = "checkboxes"
        "user"         = "text"
        "lookup"       = "text"
        "url"          = "link"
        "picture"      = "image"
        "calculated"   = "text"
        "attachments"  = "upload"
        "taxonomy"     = "text"
    }
    $HuduAssetLayoutFieldType = $SharePointToHuduMap["$($SPListItemType.Trim().ToLowerInvariant())"]

    # Sample just the values of this field across items
    $fieldValues = $SampleItems | ForEach-Object { $_.fields.$FieldName }

    if ($HuduAssetLayoutFieldType -eq "image") {
        return if (Test-IsImageField $fieldValues) { "image" } else { "upload" }
    }

    if ($HuduAssetLayoutFieldType -eq "number") {
        return if (Test-IsIntegerField $fieldValues) { "number" } else { "text" }
    }

    return $HuduAssetLayoutFieldType
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
            $discoveredFiles.AddRange($childFiles)
        }
        elseif ($item.file) {
            $downloadUrl = $item."@microsoft.graph.downloadUrl"
            Invoke-WebRequest -Uri $downloadUrl -OutFile $itemPath -UseBasicParsing
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
                webViewUrl          = $item.webUrl
                webDAVUrl           = $item.webDavUrl
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
                Filesize            = (Get-Item $itemPath).Length
                FileTooLarge        = ((Get-Item $itemPath).Length -gt 100MB)
            })
            Set-PrintAndLog -message "Downloaded: $itemPath" -Color DarkMagenta
        }
    }
    return $discoveredFiles
}
