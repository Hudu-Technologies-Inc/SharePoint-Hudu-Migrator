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


function Get-SPColumnType {
    param ([pscustomobject]$col)

    foreach ($type in 'text','note','number','choice','multichoice','boolean','dateTime','currency','url','lookup','user','calculated','taxonomy') {
        if ($col.PSObject.Properties.Name -contains $type -and $col.$type) {
            return $type
        }
    }

    return 'text'  # Default fallback
}
function Get-SPColumnChoices {
    param ([pscustomobject]$col)

    $fieldType = Get-SPColumnType $col
    if ($fieldType -eq 'choice' -or $fieldType -eq 'multichoice') {
        $options = $col.$fieldType.choices
        Write-Host "Field '$($col.displayName)' has options:"
        foreach ($opt in $options) {
            Write-Host "  - $opt"
        }
        return $options
    }

    return @()  # Return empty array if not a choice field
}


function Get-SPListItemTypeToHuduALType {
param (
    [string]$SPListItemType,
    [string]$FieldName,
    [array]$SampleValues
)
    $SharePointToHuduMap = @{
        "text"         = "Text"
        "note"         = "RichText"
        "number"       = "Number"
        "currency"     = "Text"
        "boolean"      = "CheckBox"
        "datetime"     = "Date"
        "choice"       = "ListSelect"
        "multichoice"  = "ListSelect"
        "user"         = "RichText"
        "lookup"       = "RichText"
        "url"          = "Website"
        "picture"      = "RichText"
        "calculated"   = "Text"
        "attachments"  = "RichText"
        "taxonomy"     = "RichText"
    }
    $HuduAssetLayoutFieldType = $SharePointToHuduMap["$($SPListItemType.Trim().ToLowerInvariant())"]

    $fieldValues = $SampleItems | ForEach-Object { $_.fields.$FieldName }

    if ($HuduAssetLayoutFieldType -eq "Number") {
        if (Test-IsIntegerField $fieldValues) {
            return "Number"
        } else {
            return "Text"
        }
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
