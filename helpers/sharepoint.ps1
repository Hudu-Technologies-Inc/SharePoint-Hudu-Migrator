if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'sharepoint\manifests.ps1')) {
    . (Join-Path $PSScriptRoot 'sharepoint\manifests.ps1')
}

function Update-SharePointAccessToken {
    param (
        [switch]$Force,
        [int]$RefreshWindowMinutes = 10
    )

    if ($null -eq $global:SharePointAuthState) {
        $global:SharePointAuthState = @{}
    }

    $authState = $global:SharePointAuthState

    foreach ($name in @('tokenResult', 'clientId', 'tenantId', 'scopes', 'accessToken', 'SharePointHeaders')) {
        if (-not $authState.ContainsKey($name) -or $null -eq $authState[$name]) {
            $visibleVariable = Get-Variable -Name $name -ErrorAction SilentlyContinue
            if ($visibleVariable) {
                $authState[$name] = $visibleVariable.Value
            }
        }
    }

    $nowUtc = (Get-Date).ToUniversalTime()
    $expiresOnUtc = $null

    if ($null -ne $authState.tokenResult -and $null -ne $authState.tokenResult.ExpiresOn) {
        $expiresOnUtc = $authState.tokenResult.ExpiresOn.UtcDateTime
    }

    $shouldRefresh = (
        $Force -or
        $null -eq $authState.tokenResult -or
        [string]::IsNullOrWhiteSpace([string]$authState.tokenResult.AccessToken) -or
        $null -eq $expiresOnUtc -or
        $expiresOnUtc -le $nowUtc.AddMinutes($RefreshWindowMinutes)
    )

    if ($shouldRefresh) {
        foreach ($requiredName in @('clientId', 'tenantId', 'scopes')) {
            if ([string]::IsNullOrWhiteSpace([string]$authState[$requiredName])) {
                throw "Cannot refresh SharePoint Graph access token because `$$requiredName is not available."
            }
        }

        $previousExpiry = if ($expiresOnUtc) { $expiresOnUtc.ToString('u') } else { 'unknown' }
        Set-PrintAndLog -message "Refreshing SharePoint Graph access token. Previous expiry: $previousExpiry" -Color DarkCyan

        try {
            $authState.tokenResult = Get-MsalToken `
                -ClientId $authState.clientId `
                -TenantId $authState.tenantId `
                -Scopes $authState.scopes `
                -Silent `
                -ForceRefresh:$Force `
                -ErrorAction Stop
        } catch {
            Set-PrintAndLog -message "Silent SharePoint token refresh failed; falling back to device code authentication. $($_.Exception.Message)" -Color Yellow
            $authState.tokenResult = Get-MsalToken `
                -ClientId $authState.clientId `
                -TenantId $authState.tenantId `
                -DeviceCode `
                -Scopes $authState.scopes `
                -ErrorAction Stop
        }

        $authState.accessToken = $authState.tokenResult.AccessToken
        $authState.SharePointHeaders = @{ Authorization = "Bearer $($authState.accessToken)" }

        Set-Variable -Name tokenResult -Value $authState.tokenResult -Scope Global -Force
        Set-Variable -Name accessToken -Value $authState.accessToken -Scope Global -Force
        Set-Variable -Name SharePointHeaders -Value $authState.SharePointHeaders -Scope Global -Force

        if ($authState.tokenResult.ExpiresOn) {
            Set-PrintAndLog -message "SharePoint Graph access token refreshed. New expiry: $($authState.tokenResult.ExpiresOn.UtcDateTime.ToString('u'))" -Color DarkCyan
        }
    }
    elseif (
        ($null -eq $authState.SharePointHeaders -or -not $authState.SharePointHeaders.ContainsKey('Authorization')) -and
        $authState.tokenResult -and
        -not [string]::IsNullOrWhiteSpace([string]$authState.tokenResult.AccessToken)
    ) {
        $authState.accessToken = $authState.tokenResult.AccessToken
        $authState.SharePointHeaders = @{ Authorization = "Bearer $($authState.accessToken)" }
        Set-Variable -Name accessToken -Value $authState.accessToken -Scope Global -Force
        Set-Variable -Name SharePointHeaders -Value $authState.SharePointHeaders -Scope Global -Force
    }

    if ($null -eq $authState.SharePointHeaders -or -not $authState.SharePointHeaders.ContainsKey('Authorization')) {
        throw "SharePoint Graph authorization headers are not available."
    }

    return $authState.SharePointHeaders
}

function Invoke-SharePointGraphCollection {
    param (
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $items = [System.Collections.ArrayList]@()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-RestMethod -Headers (Update-SharePointAccessToken) -Uri $nextUri -Method Get
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

function Get-SharePointDriveItemSourceKey {
    param ($Item)

    if ($Item.id) {
        return 'sharepoint:driveItem:{0}' -f $Item.id
    }

    if ($Item.webUrl) {
        return 'sharepoint:driveItemUrl:{0}' -f $Item.webUrl
    }

    return $null
}

function Get-SharePointDriveItemSourceETag {
    param ($Item)

    return ($Item.eTag ?? $Item.cTag)
}

function Test-SharePointDriveItemAlreadyCompleted {
    param ($Item)

    if (-not $RunSummary.SetupInfo.ResumeFromState) { return $false }
    if ($null -eq $SharePointMigrationState) { return $false }

    $resumeItem = [PSCustomObject]@{
        SourceKey  = Get-SharePointDriveItemSourceKey -Item $Item
        SourceETag = Get-SharePointDriveItemSourceETag -Item $Item
    }

    Test-SharePointItemAlreadyMigrated `
        -Item $resumeItem `
        -State $SharePointMigrationState `
        -IgnoreETag:$RunSummary.SetupInfo.ResumeIgnoreETag
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
        SourceKey           = Get-SharePointDriveItemSourceKey -Item $Item
        SourceETag          = Get-SharePointDriveItemSourceETag -Item $Item
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
        if (Test-SharePointDriveItemAlreadyCompleted -Item $item) {
            Set-PrintAndLog -message "Skipping already completed SharePoint item: $itemPath" -Color DarkGray
            $RunSummary.JobInfo.ArticlesSkipped++
            return $discoveredFiles
        }

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
