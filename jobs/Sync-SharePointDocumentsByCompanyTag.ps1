<#
.SYNOPSIS
Dry-run first SharePoint document sync that uses trailing company tags to place Hudu articles.

.DESCRIPTION
For each file in a specific SharePoint site drive, this job:
- Uses the filename without extension as the Hudu article title.
- Reads the trailing parenthetical tag from that title, for example "VPN Notes (ACME)" => "ACME".
- Matches that tag to the trailing parenthetical tag on Hudu company names.
- Preserves the SharePoint folder path under the matched company KB.
- Moves an exact-title article if one exists elsewhere, or creates it when missing.

No Hudu changes are made unless -Apply is supplied.

.EXAMPLE
.\jobs\Sync-SharePointDocumentsByCompanyTag.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Clients" -MaxItems 25

.EXAMPLE
.\jobs\Sync-SharePointDocumentsByCompanyTag.ps1 -SiteId "contoso.sharepoint.com,siteCollectionId,webId" -DriveNames "Documents" -DestinationRootFolderName "SharePoint" -Apply
#>

param(
    [hashtable]$Headers = $SharePointHeaders,

    # Provide either SiteId or SiteUrl. SiteId is the Graph site id.
    [string]$SiteId,
    [string]$SiteUrl,

    [string[]]$DriveIds = @(),
    [string[]]$DriveNames = @(),

    [string]$DestinationRootFolderName = "",

    # Use when exporting multiple drives and you want "Documents\Folder\File" in Hudu.
    [switch]$IncludeDriveNameInFolderPath,

    [string]$WorkingDirectory = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'downloads') 'tagged-document-sync'),
    [string]$ReportPath = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'logs') 'sharepoint-tagged-document-sync.csv'),

    [bool]$MoveExistingArticles = $true,
    [bool]$CreateMissingArticles = $true,
    [bool]$RefreshExistingContent = $false,
    [bool]$UploadSourceFile = $false,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxItems = 0,

    # Mutations only happen when -Apply is supplied.
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

function Write-TaggedDocumentSyncLog {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [string]$Color = 'White'
    )

    if (Get-Command Set-PrintAndLog -ErrorAction SilentlyContinue) {
        try {
            Set-PrintAndLog -message $Message -Color $Color
            return
        } catch {}
    }

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message" -ForegroundColor $Color
}

function ConvertTo-TaggedDocumentSyncKey {
    param($Value)

    if ($null -eq $Value) { return "" }
    $text = ([string]$Value).Normalize([Text.NormalizationForm]::FormD).ToLowerInvariant()
    $text = $text -replace '\p{Mn}', ''
    try {
        $text = [System.Web.HttpUtility]::HtmlDecode($text)
    } catch {
        $text = [System.Net.WebUtility]::HtmlDecode($text)
    }
    $text = $text -replace '&', ' and '
    $text = $text -replace '[^a-z0-9]+', ' '
    return ($text -replace '\s+', ' ').Trim()
}

function Get-TaggedDocumentSyncTrailingTag {
    param($Value)

    if ($null -eq $Value) { return $null }
    $match = [regex]::Match([string]$Value, '\((?<tag>[^()]*)\)\s*$')
    if (-not $match.Success) { return $null }

    $tag = $match.Groups['tag'].Value.Trim()
    if ([string]::IsNullOrWhiteSpace($tag)) { return $null }
    return $tag
}

function Get-TaggedDocumentSyncSafePathName {
    param(
        [string]$Name,
        [string]$Fallback = 'unnamed'
    )

    $value = if ([string]::IsNullOrWhiteSpace($Name)) { $Fallback } else { $Name }
    $safe = (($value -replace '[\\/:*?"<>|]', '_') -replace '\s{2,}', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { return $Fallback }
    return $safe
}

function Get-TaggedDocumentSyncArticleTitle {
    param([Parameter(Mandatory)] $DriveItem)

    $name = [string]($DriveItem.name ?? $DriveItem.Name ?? 'Untitled')
    $title = [System.IO.Path]::GetFileNameWithoutExtension($name)
    if ([string]::IsNullOrWhiteSpace($title)) { return $name }
    return $title
}

function Get-TaggedDocumentSyncHeaders {
    if (Get-Command Update-SharePointAccessToken -ErrorAction SilentlyContinue) {
        return Update-SharePointAccessToken
    }

    if ($null -eq $script:Headers -or -not $script:Headers.ContainsKey('Authorization')) {
        throw "SharePoint Graph headers are not available. Run your auth setup first, or pass -Headers."
    }

    return $script:Headers
}

function Invoke-TaggedDocumentSyncGraphRequest {
    param(
        [Parameter(Mandatory)] [string]$Uri,
        [ValidateRange(0, 20)] [int]$MaxRetries = 6
    )

    $attempt = 0
    while ($true) {
        try {
            return Invoke-RestMethod `
                -Method Get `
                -Uri $Uri `
                -Headers (Get-TaggedDocumentSyncHeaders) `
                -ErrorAction Stop
        } catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
            $isTransient = $statusCode -in @(429, 502, 503, 504)

            if (-not $isTransient -or $attempt -ge $MaxRetries) {
                throw
            }

            $delaySeconds = [math]::Min(60, [math]::Pow(2, $attempt + 1))
            Write-TaggedDocumentSyncLog -Message "Graph returned HTTP $statusCode. Retrying in $delaySeconds second(s): $Uri" -Color Yellow
            Start-Sleep -Seconds $delaySeconds
            $attempt++
        }
    }
}

function Invoke-TaggedDocumentSyncGraphCollection {
    param([Parameter(Mandatory)] [string]$Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-TaggedDocumentSyncGraphRequest -Uri $nextUri
        foreach ($item in @($response.value)) {
            $items.Add($item)
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return @($items)
}

function Resolve-TaggedDocumentSyncSite {
    param(
        [string]$GraphSiteId,
        [string]$SharePointSiteUrl
    )

    if (-not [string]::IsNullOrWhiteSpace($GraphSiteId)) {
        return Invoke-TaggedDocumentSyncGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$GraphSiteId"
    }

    if ([string]::IsNullOrWhiteSpace($SharePointSiteUrl)) {
        throw "Provide -SiteId or -SiteUrl."
    }

    $parsedUrl = [uri]$SharePointSiteUrl
    $hostname = $parsedUrl.Host
    $serverRelativePath = $parsedUrl.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($serverRelativePath)) {
        $serverRelativePath = '/'
    }

    Invoke-TaggedDocumentSyncGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/${hostname}:$serverRelativePath"
}

function Get-TaggedDocumentSyncDriveItems {
    param(
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] $Drive,
        [string]$FolderId = 'root',
        [string[]]$FolderPath = @()
    )

    $itemsUri = "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($Drive.id)/items/$FolderId/children"
    $items = @(Invoke-TaggedDocumentSyncGraphCollection -Uri $itemsUri)

    foreach ($item in $items) {
        if ($item.folder) {
            $nextPath = @($FolderPath) + @([string]$item.name)
            foreach ($child in @(Get-TaggedDocumentSyncDriveItems -Site $Site -Drive $Drive -FolderId $item.id -FolderPath $nextPath)) {
                $child
            }
            continue
        }

        if (-not $item.file) {
            continue
        }

        [PSCustomObject]@{
            Item       = $item
            FolderPath = @($FolderPath)
            Drive      = $Drive
            Site       = $Site
        }
    }
}

function New-TaggedDocumentSyncCompanyTagIndex {
    param($Companies)

    $index = @{}
    foreach ($company in @($Companies)) {
        $companyId = $company.id ?? $company.Id
        $companyName = [string]($company.name ?? $company.Name)
        $tag = Get-TaggedDocumentSyncTrailingTag -Value $companyName
        $tagKey = ConvertTo-TaggedDocumentSyncKey -Value $tag

        if (-not $companyId -or [string]::IsNullOrWhiteSpace($tagKey)) {
            continue
        }

        if (-not $index.ContainsKey($tagKey)) {
            $index[$tagKey] = [System.Collections.Generic.List[object]]::new()
        }

        $index[$tagKey].Add([PSCustomObject]@{
            Id     = [int]$companyId
            Name   = $companyName
            Tag    = $tag
            TagKey = $tagKey
            Object = $company
        })
    }

    return $index
}

function Get-TaggedDocumentSyncArticleObject {
    param($Article)

    return ($Article.article ?? $Article.Article ?? $Article)
}

function Expand-TaggedDocumentSyncArticles {
    param($InputObject)

    $expanded = [System.Collections.Generic.List[object]]::new()
    foreach ($item in @($InputObject)) {
        if ($null -eq $item) { continue }

        $articleSet = $item.articles ?? $item.Articles
        if ($articleSet) {
            foreach ($article in @($articleSet)) {
                $expanded.Add((Get-TaggedDocumentSyncArticleObject -Article $article))
            }
            continue
        }

        $expanded.Add((Get-TaggedDocumentSyncArticleObject -Article $item))
    }

    return @($expanded)
}

function Get-TaggedDocumentSyncArticleName {
    param($Article)

    $article = Get-TaggedDocumentSyncArticleObject -Article $Article
    return [string]($article.name ?? $article.Name ?? $article.title ?? $article.Title)
}

function Get-TaggedDocumentSyncArticleId {
    param($Article)

    $article = Get-TaggedDocumentSyncArticleObject -Article $Article
    return ($article.id ?? $article.Id ?? $article.article_id ?? $article.ArticleId)
}

function Get-TaggedDocumentSyncArticleCompanyId {
    param($Article)

    $article = Get-TaggedDocumentSyncArticleObject -Article $Article
    return ($article.company_id ?? $article.companyId ?? $article.CompanyId ?? $article.company.id ?? $article.Company.Id)
}

function Find-TaggedDocumentSyncExistingArticle {
    param(
        [Parameter(Mandatory)] [string]$Title,
        [Parameter(Mandatory)] [int]$TargetCompanyId
    )

    if (-not (Get-Command Get-HuduArticles -ErrorAction SilentlyContinue)) {
        throw "Get-HuduArticles is not available. Load the Hudu API module/auth before running this job."
    }

    $articles = @(Expand-TaggedDocumentSyncArticles -InputObject (Get-HuduArticles -Name $Title))
    $titleKey = $Title.Trim().ToLowerInvariant()
    $exactMatches = @($articles | Where-Object {
        (Get-TaggedDocumentSyncArticleName -Article $_).Trim().ToLowerInvariant() -eq $titleKey
    })

    if ($exactMatches.Count -lt 1) {
        return [PSCustomObject]@{
            Status  = 'NotFound'
            Article = $null
            Count   = 0
        }
    }

    $targetCompanyMatch = @($exactMatches | Where-Object {
        [string](Get-TaggedDocumentSyncArticleCompanyId -Article $_) -eq [string]$TargetCompanyId
    } | Select-Object -First 1)

    if ($targetCompanyMatch.Count -gt 0) {
        return [PSCustomObject]@{
            Status  = 'FoundInTargetCompany'
            Article = $targetCompanyMatch[0]
            Count   = $exactMatches.Count
        }
    }

    if ($exactMatches.Count -eq 1) {
        return [PSCustomObject]@{
            Status  = 'FoundElsewhere'
            Article = $exactMatches[0]
            Count   = 1
        }
    }

    return [PSCustomObject]@{
        Status  = 'DuplicateExactTitle'
        Article = $null
        Count   = $exactMatches.Count
    }
}

function Invoke-TaggedDocumentSyncHuduRequest {
    param(
        [Parameter(Mandatory)] [ValidateSet('Get', 'Post', 'Put')] [string]$Method,
        [Parameter(Mandatory)] [string]$Resource,
        [string]$Body
    )

    if (Get-Command Invoke-HuduRequest -ErrorAction SilentlyContinue) {
        if ($Body) {
            return Invoke-HuduRequest -Method $Method.ToLowerInvariant() -Resource $Resource -Body $Body
        }
        return Invoke-HuduRequest -Method $Method.ToLowerInvariant() -Resource $Resource
    }

    $baseUrl = [string](Get-HuduBaseURL)
    $apiKey = (New-Object PSCredential 'user', (Get-HuduApiKey)).GetNetworkCredential().Password
    if ([string]::IsNullOrWhiteSpace($baseUrl) -or [string]::IsNullOrWhiteSpace($apiKey)) {
        throw "Hudu base URL/API key are not initialized."
    }

    $params = @{
        Method      = $Method
        Uri         = "{0}{1}" -f $baseUrl.TrimEnd('/'), $Resource
        Headers     = @{ 'x-api-key' = $apiKey }
        ErrorAction = 'Stop'
    }
    if ($Body) {
        $params.Body = $Body
        $params.ContentType = 'application/json'
    }

    Invoke-RestMethod @params
}

function New-TaggedDocumentSyncArticleBody {
    param(
        [Parameter(Mandatory)] $DriveItem,
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] $Drive,
        [string[]]$FolderPath = @()
    )

    $encodedName = [System.Net.WebUtility]::HtmlEncode([string]$DriveItem.name)
    $encodedSite = [System.Net.WebUtility]::HtmlEncode([string]($Site.displayName ?? $Site.name))
    $encodedDrive = [System.Net.WebUtility]::HtmlEncode([string]$Drive.name)
    $encodedPath = [System.Net.WebUtility]::HtmlEncode((@($FolderPath) -join '\'))
    $encodedUrl = [System.Net.WebUtility]::HtmlEncode([string]$DriveItem.webUrl)
    $encodedModified = [System.Net.WebUtility]::HtmlEncode([string]$DriveItem.lastModifiedDateTime)

@"
<p><strong>Source SharePoint document:</strong> <a href="$encodedUrl" target="_blank">$encodedName</a></p>
<table>
  <tbody>
    <tr><th>SharePoint site</th><td>$encodedSite</td></tr>
    <tr><th>Drive</th><td>$encodedDrive</td></tr>
    <tr><th>Folder path</th><td>$encodedPath</td></tr>
    <tr><th>Last modified</th><td>$encodedModified</td></tr>
  </tbody>
</table>
"@
}

function New-TaggedDocumentSyncArticle {
    param(
        [Parameter(Mandatory)] [string]$Title,
        [Parameter(Mandatory)] [string]$Content,
        [Parameter(Mandatory)] [int]$CompanyId,
        $FolderId
    )

    if (Get-Command New-HuduArticle -ErrorAction SilentlyContinue) {
        $params = @{
            Name      = $Title
            Content   = $Content
            CompanyId = $CompanyId
        }
        if ($FolderId) { $params.FolderId = $FolderId }

        $created = New-HuduArticle @params
        return ($created.article ?? $created.Article ?? $created)
    }

    $article = @{
        name       = $Title
        content    = $Content
        company_id = $CompanyId
    }
    if ($FolderId) { $article.folder_id = $FolderId }

    $createdResponse = Invoke-TaggedDocumentSyncHuduRequest `
        -Method Post `
        -Resource '/api/v1/articles' `
        -Body (@{ article = $article } | ConvertTo-Json -Depth 20)

    return ($createdResponse.article ?? $createdResponse)
}

function Set-TaggedDocumentSyncArticle {
    param(
        [Parameter(Mandatory)] [int]$ArticleId,
        [Parameter(Mandatory)] [int]$CompanyId,
        $FolderId,
        [string]$Name,
        [string]$Content,
        [bool]$UpdateContent
    )

    $object = Invoke-TaggedDocumentSyncHuduRequest -Method Get -Resource "/api/v1/articles/$ArticleId"
    $article = $object.article ?? $object
    if (-not $article) { throw "Hudu article $ArticleId was not returned." }

    $article.company_id = $CompanyId
    $article.folder_id = if ($FolderId) { $FolderId } else { $null }
    if (-not [string]::IsNullOrWhiteSpace($Name)) {
        $article.name = $Name
    }
    if ($UpdateContent) {
        $article.content = $Content
    }

    $updated = Invoke-TaggedDocumentSyncHuduRequest `
        -Method Put `
        -Resource "/api/v1/articles/$ArticleId" `
        -Body (@{ article = $article } | ConvertTo-Json -Depth 30)

    return ($updated.article ?? $updated)
}

function Get-TaggedDocumentSyncFolderParentId {
    param($Folder)

    return ($Folder.parent_folder_id ?? $Folder.ParentFolderId ?? $Folder.parentFolderId)
}

function Get-TaggedDocumentSyncFolderName {
    param($Folder)

    return [string]($Folder.name ?? $Folder.Name)
}

function New-TaggedDocumentSyncFolderIndex {
    param($Folders)

    $folderById = @{}
    $childrenByParent = @{}

    foreach ($folder in @($Folders)) {
        if ($null -eq $folder) { continue }

        $id = [string]($folder.id ?? $folder.Id)
        if (-not [string]::IsNullOrWhiteSpace($id)) {
            $folderById[$id] = $folder
        }

        $parentId = [string](Get-TaggedDocumentSyncFolderParentId -Folder $folder)
        if ([string]::IsNullOrWhiteSpace($parentId)) { $parentId = "" }
        if (-not $childrenByParent.ContainsKey($parentId)) {
            $childrenByParent[$parentId] = [System.Collections.Generic.List[object]]::new()
        }
        $childrenByParent[$parentId].Add($folder)
    }

    [PSCustomObject]@{
        FolderById       = $folderById
        ChildrenByParent = $childrenByParent
    }
}

function Find-TaggedDocumentSyncChildFolder {
    param(
        [Parameter(Mandatory)] [hashtable]$ChildrenByParent,
        $ParentId,
        [Parameter(Mandatory)] [string]$Name
    )

    $parentKey = if ($ParentId) { [string]$ParentId } else { "" }
    if (-not $ChildrenByParent.ContainsKey($parentKey)) { return $null }

    $nameKey = ConvertTo-TaggedDocumentSyncKey -Value $Name
    return @($ChildrenByParent[$parentKey] | Where-Object {
        (ConvertTo-TaggedDocumentSyncKey -Value (Get-TaggedDocumentSyncFolderName -Folder $_)) -eq $nameKey
    } | Select-Object -First 1)
}

function Add-TaggedDocumentSyncFolderToIndex {
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)] [hashtable]$FolderById,
        [Parameter(Mandatory)] [hashtable]$ChildrenByParent
    )

    $id = [string]($Folder.id ?? $Folder.Id)
    if (-not [string]::IsNullOrWhiteSpace($id)) {
        $FolderById[$id] = $Folder
    }

    $parentId = [string](Get-TaggedDocumentSyncFolderParentId -Folder $Folder)
    if ([string]::IsNullOrWhiteSpace($parentId)) { $parentId = "" }
    if (-not $ChildrenByParent.ContainsKey($parentId)) {
        $ChildrenByParent[$parentId] = [System.Collections.Generic.List[object]]::new()
    }
    $ChildrenByParent[$parentId].Add($Folder)
}

function New-TaggedDocumentSyncFolder {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [int]$CompanyId,
        $ParentFolderId
    )

    if (Get-Command New-HuduFolder -ErrorAction SilentlyContinue) {
        $params = @{
            Name      = $Name
            CompanyId = $CompanyId
        }
        if ($ParentFolderId) { $params.ParentFolderId = $ParentFolderId }

        $created = New-HuduFolder @params
        return ($created.folder ?? $created.Folder ?? $created)
    }

    $folder = @{
        name       = $Name
        company_id = $CompanyId
    }
    if ($ParentFolderId) { $folder.parent_folder_id = $ParentFolderId }

    $createdResponse = Invoke-TaggedDocumentSyncHuduRequest `
        -Method Post `
        -Resource '/api/v1/folders' `
        -Body (@{ folder = $folder } | ConvertTo-Json -Depth 20)

    return ($createdResponse.folder ?? $createdResponse)
}

function Get-TaggedDocumentSyncFolderIndexForCompany {
    param(
        [Parameter(Mandatory)] [int]$CompanyId,
        [Parameter(Mandatory)] [hashtable]$FolderIndexByCompany
    )

    $companyKey = [string]$CompanyId
    if ($FolderIndexByCompany.ContainsKey($companyKey)) {
        return $FolderIndexByCompany[$companyKey]
    }

    if (-not (Get-Command Get-HuduFolders -ErrorAction SilentlyContinue)) {
        throw "Get-HuduFolders is not available. Load the Hudu API module/auth before running this job."
    }

    $folders = @(Get-HuduFolders -CompanyId $CompanyId)
    $index = New-TaggedDocumentSyncFolderIndex -Folders $folders
    $FolderIndexByCompany[$companyKey] = $index
    return $index
}

function Ensure-TaggedDocumentSyncFolderPath {
    param(
        [Parameter(Mandatory)] [string[]]$Path,
        [Parameter(Mandatory)] [int]$CompanyId,
        [Parameter(Mandatory)] [hashtable]$FolderById,
        [Parameter(Mandatory)] [hashtable]$ChildrenByParent,
        [switch]$DryRun
    )

    if (@($Path).Count -lt 1) { return $null }

    $parentId = $null
    $lastFolder = $null
    $pathSoFar = [System.Collections.Generic.List[string]]::new()

    foreach ($folderName in @($Path)) {
        if ([string]::IsNullOrWhiteSpace($folderName)) { continue }

        $pathSoFar.Add($folderName)
        $existing = Find-TaggedDocumentSyncChildFolder -ChildrenByParent $ChildrenByParent -ParentId $parentId -Name $folderName
        if ($existing) {
            $lastFolder = $existing
            $parentId = $existing.id ?? $existing.Id
            continue
        }

        if ($DryRun) {
            $createdFolder = [PSCustomObject]@{
                id               = "dryrun:${CompanyId}:$(@($pathSoFar) -join '/')"
                name             = $folderName
                parent_folder_id = $parentId
                company_id       = $CompanyId
            }
        } else {
            $createdFolder = New-TaggedDocumentSyncFolder -Name $folderName -CompanyId $CompanyId -ParentFolderId $parentId
        }

        Add-TaggedDocumentSyncFolderToIndex -Folder $createdFolder -FolderById $FolderById -ChildrenByParent $ChildrenByParent
        $lastFolder = $createdFolder
        $parentId = $createdFolder.id ?? $createdFolder.Id
    }

    return $lastFolder
}

function Save-TaggedDocumentSyncDriveItem {
    param(
        [Parameter(Mandatory)] $DriveItem,
        [Parameter(Mandatory)] [string]$TargetDirectory
    )

    if ($DriveItem.size -ge 100MB) {
        return $null
    }

    $downloadUrl = $DriveItem.'@microsoft.graph.downloadUrl'
    if ([string]::IsNullOrWhiteSpace([string]$downloadUrl)) {
        return $null
    }

    if (-not (Test-Path -LiteralPath $TargetDirectory -PathType Container)) {
        $null = New-Item -ItemType Directory -Path $TargetDirectory -Force
    }

    $safeName = Get-TaggedDocumentSyncSafePathName -Name $DriveItem.name -Fallback $DriveItem.id
    $targetPath = Join-Path $TargetDirectory $safeName
    Invoke-WebRequest -Uri $downloadUrl -OutFile $targetPath -UseBasicParsing
    return $targetPath
}

function Add-TaggedDocumentSyncSourceUpload {
    param(
        [Parameter(Mandatory)] [string]$FilePath,
        [Parameter(Mandatory)] [int]$ArticleId
    )

    if (-not (Get-Command New-HuduUpload -ErrorAction SilentlyContinue)) {
        throw "New-HuduUpload is not available. Load the Hudu API module before using -UploadSourceFile."
    }

    $upload = New-HuduUpload -FilePath $FilePath -record_id $ArticleId -record_type 'Article'
    return ($upload.upload ?? $upload)
}

$dryRun = -not $Apply

$resolvedWorkingDirectory = if ([System.IO.Path]::IsPathRooted($WorkingDirectory)) {
    [System.IO.Path]::GetFullPath($WorkingDirectory)
} else {
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $WorkingDirectory))
}

$resolvedReportPath = if ([System.IO.Path]::IsPathRooted($ReportPath)) {
    [System.IO.Path]::GetFullPath($ReportPath)
} else {
    [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $ReportPath))
}

$reportDirectory = Split-Path -Parent $resolvedReportPath
if (-not (Test-Path -LiteralPath $reportDirectory -PathType Container)) {
    $null = New-Item -ItemType Directory -Path $reportDirectory -Force
}

if ($UploadSourceFile -and -not (Test-Path -LiteralPath $resolvedWorkingDirectory -PathType Container)) {
    $null = New-Item -ItemType Directory -Path $resolvedWorkingDirectory -Force
}

if (-not (Get-Command Get-HuduCompanies -ErrorAction SilentlyContinue)) {
    throw "Get-HuduCompanies is not available. Load the Hudu API module/auth before running this job."
}

$site = Resolve-TaggedDocumentSyncSite -GraphSiteId $SiteId -SharePointSiteUrl $SiteUrl
$siteLabel = [string]($site.displayName ?? $site.name ?? $site.id)

$drives = @(Invoke-TaggedDocumentSyncGraphCollection -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives")
if ($DriveIds.Count -gt 0) {
    $driveIdSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($driveId in @($DriveIds)) { [void]$driveIdSet.Add([string]$driveId) }
    $drives = @($drives | Where-Object { $driveIdSet.Contains([string]$_.id) })
}
if ($DriveNames.Count -gt 0) {
    $driveNameKeys = @($DriveNames | ForEach-Object { ConvertTo-TaggedDocumentSyncKey -Value $_ })
    $drives = @($drives | Where-Object {
        $driveNameKey = ConvertTo-TaggedDocumentSyncKey -Value $_.name
        $driveNameKeys -contains $driveNameKey
    })
}

Write-TaggedDocumentSyncLog -Message "SharePoint tagged document sync for '$siteLabel'. Drives=$($drives.Count); DryRun=$dryRun; MoveExisting=$MoveExistingArticles; CreateMissing=$CreateMissingArticles; RefreshExisting=$RefreshExistingContent." -Color Cyan

$companyTagIndex = New-TaggedDocumentSyncCompanyTagIndex -Companies @(Get-HuduCompanies)
$duplicateCompanyTags = @($companyTagIndex.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 })
if ($duplicateCompanyTags.Count -gt 0) {
    Write-TaggedDocumentSyncLog -Message "Found $($duplicateCompanyTags.Count) duplicate company tag(s). Documents with those tags will be skipped." -Color Yellow
}

$folderIndexByCompany = @{}
$reportRows = [System.Collections.Generic.List[object]]::new()
$processed = 0
$created = 0
$moved = 0
$updated = 0
$skipped = 0
$failed = 0
$uploaded = 0

foreach ($drive in @($drives)) {
    $driveLabel = [string]($drive.name ?? $drive.id)
    Write-TaggedDocumentSyncLog -Message "Scanning drive '$driveLabel'." -Color DarkCyan

    foreach ($entry in @(Get-TaggedDocumentSyncDriveItems -Site $site -Drive $drive)) {
        if ($MaxItems -gt 0 -and $processed -ge $MaxItems) {
            break
        }

        $processed++
        $item = $entry.Item
        $articleTitle = Get-TaggedDocumentSyncArticleTitle -DriveItem $item
        $documentTag = Get-TaggedDocumentSyncTrailingTag -Value $articleTitle
        $documentTagKey = ConvertTo-TaggedDocumentSyncKey -Value $documentTag
        $destinationCompanyId = $null
        $destinationCompanyName = $null
        $destinationFolderId = $null
        $existingArticleId = $null
        $existingArticleCompanyId = $null
        $action = $null
        $status = $null
        $errorMessage = $null
        $uploadPath = $null
        $uploadUrl = $null

        $destinationPath = [System.Collections.Generic.List[string]]::new()
        if (-not [string]::IsNullOrWhiteSpace($DestinationRootFolderName)) {
            $destinationPath.Add($DestinationRootFolderName)
        }
        if ($IncludeDriveNameInFolderPath) {
            $destinationPath.Add($driveLabel)
        }
        foreach ($part in @($entry.FolderPath)) {
            $destinationPath.Add([string]$part)
        }

        try {
            if ([string]::IsNullOrWhiteSpace($documentTagKey)) {
                $status = 'SkippedNoDocumentTag'
                $skipped++
                continue
            }

            if (-not $companyTagIndex.ContainsKey($documentTagKey)) {
                $status = 'SkippedNoMatchingCompanyTag'
                $skipped++
                continue
            }

            $companyMatches = @($companyTagIndex[$documentTagKey])
            if ($companyMatches.Count -gt 1) {
                $status = 'SkippedDuplicateCompanyTag'
                $skipped++
                continue
            }

            $matchedCompany = $companyMatches[0]
            $destinationCompanyId = [int]$matchedCompany.Id
            $destinationCompanyName = [string]$matchedCompany.Name

            if ($destinationPath.Count -gt 0) {
                $folderIndex = Get-TaggedDocumentSyncFolderIndexForCompany -CompanyId $destinationCompanyId -FolderIndexByCompany $folderIndexByCompany
                $folder = Ensure-TaggedDocumentSyncFolderPath `
                    -Path $destinationPath.ToArray() `
                    -CompanyId $destinationCompanyId `
                    -FolderById $folderIndex.FolderById `
                    -ChildrenByParent $folderIndex.ChildrenByParent `
                    -DryRun:$dryRun
                $destinationFolderId = $folder.id ?? $folder.Id
            }

            $content = New-TaggedDocumentSyncArticleBody -DriveItem $item -Site $site -Drive $drive -FolderPath $entry.FolderPath
            $existingResult = Find-TaggedDocumentSyncExistingArticle -Title $articleTitle -TargetCompanyId $destinationCompanyId
            $existingArticle = $existingResult.Article

            if ($existingResult.Status -eq 'DuplicateExactTitle') {
                $status = 'SkippedDuplicateExactArticleTitle'
                $skipped++
                continue
            }

            if ($existingArticle) {
                $existingArticleId = Get-TaggedDocumentSyncArticleId -Article $existingArticle
                $existingArticleCompanyId = Get-TaggedDocumentSyncArticleCompanyId -Article $existingArticle

                if (-not $MoveExistingArticles -and [string]$existingArticleCompanyId -ne [string]$destinationCompanyId) {
                    $status = 'SkippedExistingArticleMoveDisabled'
                    $skipped++
                    continue
                }

                if ($dryRun) {
                    $action = if ([string]$existingArticleCompanyId -eq [string]$destinationCompanyId) { 'WouldUpdateExistingArticleLocation' } else { 'WouldMoveExistingArticle' }
                    if ($RefreshExistingContent) { $action = "$action`AndRefreshContent" }
                    $status = 'DryRun'
                    if ([string]$existingArticleCompanyId -eq [string]$destinationCompanyId) { $updated++ } else { $moved++ }
                    continue
                }

                Set-TaggedDocumentSyncArticle `
                    -ArticleId ([int]$existingArticleId) `
                    -CompanyId $destinationCompanyId `
                    -FolderId $destinationFolderId `
                    -Name $articleTitle `
                    -Content $content `
                    -UpdateContent:$RefreshExistingContent | Out-Null

                if ($UploadSourceFile) {
                    $downloadDirectory = Join-Path $resolvedWorkingDirectory (Get-TaggedDocumentSyncSafePathName -Name $destinationCompanyName -Fallback $destinationCompanyId)
                    foreach ($part in @($destinationPath)) {
                        $downloadDirectory = Join-Path $downloadDirectory (Get-TaggedDocumentSyncSafePathName -Name $part)
                    }
                    $uploadPath = Save-TaggedDocumentSyncDriveItem -DriveItem $item -TargetDirectory $downloadDirectory
                    if ($uploadPath) {
                        $upload = Add-TaggedDocumentSyncSourceUpload -FilePath $uploadPath -ArticleId ([int]$existingArticleId)
                        $uploadUrl = $upload.url
                        $uploaded++
                    }
                }

                $action = if ([string]$existingArticleCompanyId -eq [string]$destinationCompanyId) { 'UpdatedExistingArticleLocation' } else { 'MovedExistingArticle' }
                if ($RefreshExistingContent) { $action = "$action`AndRefreshedContent" }
                $status = 'Completed'
                if ([string]$existingArticleCompanyId -eq [string]$destinationCompanyId) { $updated++ } else { $moved++ }
                Write-TaggedDocumentSyncLog -Message "$action`: '$articleTitle' => '$destinationCompanyName' / '$(@($destinationPath) -join '\')'." -Color Green
                continue
            }

            if (-not $CreateMissingArticles) {
                $status = 'SkippedMissingArticleCreateDisabled'
                $skipped++
                continue
            }

            if ($dryRun) {
                $action = 'WouldCreateArticle'
                $status = 'DryRun'
                $created++
                continue
            }

            $createdArticle = New-TaggedDocumentSyncArticle `
                -Title $articleTitle `
                -Content $content `
                -CompanyId $destinationCompanyId `
                -FolderId $destinationFolderId

            $createdArticleId = Get-TaggedDocumentSyncArticleId -Article $createdArticle
            $existingArticleId = $createdArticleId

            if ($UploadSourceFile -and $createdArticleId) {
                $downloadDirectory = Join-Path $resolvedWorkingDirectory (Get-TaggedDocumentSyncSafePathName -Name $destinationCompanyName -Fallback $destinationCompanyId)
                foreach ($part in @($destinationPath)) {
                    $downloadDirectory = Join-Path $downloadDirectory (Get-TaggedDocumentSyncSafePathName -Name $part)
                }
                $uploadPath = Save-TaggedDocumentSyncDriveItem -DriveItem $item -TargetDirectory $downloadDirectory
                if ($uploadPath) {
                    $upload = Add-TaggedDocumentSyncSourceUpload -FilePath $uploadPath -ArticleId ([int]$createdArticleId)
                    $uploadUrl = $upload.url
                    $uploaded++
                }
            }

            $action = 'CreatedArticle'
            $status = 'Completed'
            $created++
            Write-TaggedDocumentSyncLog -Message "Created article: '$articleTitle' => '$destinationCompanyName' / '$(@($destinationPath) -join '\')'." -Color Green
        } catch {
            $status = 'Failed'
            $errorMessage = $_.Exception.Message
            $failed++
            Write-TaggedDocumentSyncLog -Message "Failed '$articleTitle': $errorMessage" -Color Red
        } finally {
            if (-not $status) {
                $status = 'Skipped'
                $skipped++
            }

            $reportRows.Add([PSCustomObject]@{
                Status                   = $status
                Action                   = $action
                SiteName                 = $siteLabel
                SiteId                   = $site.id
                DriveName                = $driveLabel
                DriveId                  = $drive.id
                SharePointItemId         = $item.id
                SharePointName           = $item.name
                SharePointUrl            = $item.webUrl
                SharePointModified       = $item.lastModifiedDateTime
                ArticleTitle             = $articleTitle
                DocumentTag              = $documentTag
                DestinationCompanyId     = $destinationCompanyId
                DestinationCompanyName   = $destinationCompanyName
                DestinationFolderId      = $destinationFolderId
                DestinationFolderPath    = (@($destinationPath) -join '\')
                ExistingArticleId        = $existingArticleId
                ExistingArticleCompanyId = $existingArticleCompanyId
                SourceUploadPath         = $uploadPath
                SourceUploadUrl          = $uploadUrl
                Error                    = $errorMessage
            })
        }
    }
}

$reportRows | Export-Csv -LiteralPath $resolvedReportPath -NoTypeInformation -Encoding UTF8

$summary = [PSCustomObject]@{
    SiteName      = $siteLabel
    SiteId        = $site.id
    ReportPath    = $resolvedReportPath
    DryRun        = $dryRun
    Processed     = $processed
    Created       = $created
    Moved         = $moved
    Updated       = $updated
    UploadedFiles = $uploaded
    Skipped       = $skipped
    Failed        = $failed
}

Write-TaggedDocumentSyncLog -Message "Tagged document sync complete: processed=$processed, created=$created, moved=$moved, updated=$updated, uploaded=$uploaded, skipped=$skipped, failed=$failed. Report: $resolvedReportPath" -Color Cyan

$summary
