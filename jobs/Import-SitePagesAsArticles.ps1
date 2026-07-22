##### Optional job, Import fetched SharePoint site pages as Hudu articles

function Get-SharePointSitePageImportSafeFileName {
    param ([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return "site-page" }
    return (Get-SharePointSafePathName -Name $Name)
}

function Get-SharePointSitePageImportSelectedSiteIds {
    $ids = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($site in @($userSelectedSites)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$site.id)) {
            [void]$ids.Add([string]$site.id)
        }
    }
    return $ids
}

function Get-SharePointSitePageImportExtensionFromMimeType {
    param ([string]$MimeType)

    switch -Regex ($MimeType) {
        '^image/jpeg$' { return '.jpg' }
        '^image/png$'  { return '.png' }
        '^image/gif$'  { return '.gif' }
        '^image/webp$' { return '.webp' }
        '^image/svg\+xml$' { return '.svg' }
        '^image/bmp$'  { return '.bmp' }
        '^image/x-icon$' { return '.ico' }
        default { return '.bin' }
    }
}

function Convert-SharePointSitePageEmbeddedBase64Images {
    param (
        [Parameter(Mandatory)]
        [string]$Html,

        [Parameter(Mandatory)]
        [string]$AssetDirectory,

        [Parameter(Mandatory)]
        [string]$AssetBaseName
    )

    $attachments = [System.Collections.ArrayList]@()
    if ([string]::IsNullOrWhiteSpace($Html)) {
        return [PSCustomObject]@{
            Html        = $Html
            Attachments = $attachments
        }
    }

    if (-not (Test-Path -LiteralPath $AssetDirectory)) {
        $null = New-Item -ItemType Directory -Path $AssetDirectory -Force
    }

    $updatedHtml = [string]$Html
    $pattern = '(?is)(?<prefix>\b(?:src|href)\s*=\s*["''])(?<data>data:(?<mime>image/[-+.\w]+);base64,(?<payload>[^"'']+))(?<suffix>["''])'
    $matches = [regex]::Matches($Html, $pattern)
    $assetIndex = 0
    $seenDataUris = @{}

    foreach ($match in $matches) {
        $dataUri = [string]$match.Groups['data'].Value
        if ($seenDataUris.ContainsKey($dataUri)) {
            $updatedHtml = $updatedHtml.Replace($dataUri, $seenDataUris[$dataUri])
            continue
        }

        $mimeType = [string]$match.Groups['mime'].Value
        $payload = ([string]$match.Groups['payload'].Value) -replace '\s', ''
        if ([string]::IsNullOrWhiteSpace($payload)) { continue }

        $extension = Get-SharePointSitePageImportExtensionFromMimeType -MimeType $mimeType
        $assetIndex++
        $fileName = '{0}-embedded-{1}{2}' -f $AssetBaseName, $assetIndex, $extension
        $filePath = Join-Path $AssetDirectory $fileName

        try {
            [System.IO.File]::WriteAllBytes($filePath, [Convert]::FromBase64String($payload))
            [void]$attachments.Add($filePath)
            $seenDataUris[$dataUri] = $fileName
            $updatedHtml = $updatedHtml.Replace($dataUri, $fileName)
        } catch {
            Set-PrintAndLog -message "Failed to extract base64 image for SharePoint site page '$AssetBaseName': $($_.Exception.Message)" -Color Yellow
        }
    }

    [PSCustomObject]@{
        Html        = $updatedHtml
        Attachments = $attachments
    }
}

function Get-SharePointSitePageExternalImageSources {
    param ([string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) { return @() }

    @(
        [regex]::Matches($Html, '(?is)\bsrc\s*=\s*["''](?<url>https?://[^"'']+)["'']') |
            ForEach-Object { $_.Groups['url'].Value } |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Sort-Object -Unique
    )
}

function Convert-SharePointSitePageExportToConvertedDoc {
    param (
        [Parameter(Mandatory)]
        $PageExport,

        [Parameter(Mandatory)]
        [string]$AssetRoot
    )

    $htmlPath = [string]$PageExport.HtmlPath
    $html = [string]$PageExport.Html
    if ([string]::IsNullOrWhiteSpace($html) -and (Test-Path -LiteralPath $htmlPath -PathType Leaf)) {
        $html = Get-Content -LiteralPath $htmlPath -Raw
    }

    if ([string]::IsNullOrWhiteSpace($html)) {
        $html = "<p>No content returned for this SharePoint site page.</p>"
    }

    $title = [string]($PageExport.Title ?? $PageExport.Name ?? $PageExport.PageId ?? "Untitled SharePoint page")
    $siteName = [string]($PageExport.SiteName ?? $PageExport.SiteId ?? "SharePoint Site")
    $safeSiteName = Get-SharePointSitePageImportSafeFileName -Name $siteName
    $assetBaseName = Get-SharePointSitePageImportSafeFileName -Name ($PageExport.PageId ?? $title)
    $assetDirectory = Join-Path (Join-Path $AssetRoot $safeSiteName) $assetBaseName
    $assetResult = Convert-SharePointSitePageEmbeddedBase64Images `
        -Html $html `
        -AssetDirectory $assetDirectory `
        -AssetBaseName $assetBaseName

    $updatedHtml = [string]$assetResult.Html
    if (-not (Test-Path -LiteralPath $htmlPath -PathType Leaf)) {
        $htmlPath = Join-Path $assetDirectory "$assetBaseName.html"
    }

    if (-not (Test-Path -LiteralPath (Split-Path -Parent $htmlPath))) {
        $null = New-Item -ItemType Directory -Path (Split-Path -Parent $htmlPath) -Force
    }
    $updatedHtml | Out-File -LiteralPath $htmlPath -Encoding UTF8

    $plainText = if (Get-Command ConvertTo-SharePointSitePagePlainText -ErrorAction SilentlyContinue) {
        ConvertTo-SharePointSitePagePlainText -Html $updatedHtml
    } else {
        [System.Web.HttpUtility]::HtmlDecode(([regex]::Replace($updatedHtml, '<[^>]+>', ' '))) -replace '\s{2,}', ' '
    }

    $previewText = if ($plainText.Length -gt $RunSummary.SetupInfo.PreviewLength) {
        $plainText.Substring(0, $RunSummary.SetupInfo.PreviewLength)
    } else {
        $plainText
    }

    $attachments = [System.Collections.ArrayList]@()
    foreach ($attachment in @($assetResult.Attachments)) {
        if ($attachment) { [void]$attachments.Add($attachment) }
    }

    [PSCustomObject]@{
        Name                 = $PageExport.Name
        SourceKey            = $PageExport.SourceKey
        SourceETag           = $PageExport.SourceETag
        LocalPath            = $htmlPath
        NewPath              = $htmlPath
        SiteId               = $PageExport.SiteId
        SiteName             = $siteName
        DriveId              = $null
        DriveName            = "Site Pages"
        FolderId             = $null
        DownloadUrl          = $null
        DownloadSkipped      = $true
        webViewUrl           = $PageExport.WebUrl
        webDAVUrl            = $null
        CreatedDateTime      = $PageExport.CreatedDateTime
        LastModifiedDateTime = $PageExport.LastModifiedDateTime
        sharepointSiteUrl    = $null
        sharepointListId     = $null
        sharepointItemId     = $PageExport.PageId
        parentDrivePath      = "Site Pages"
        HuduFolder           = $null
        HuduFolderId         = $null
        HuduArticle          = $null
        HuduFolderUUID       = [guid]::NewGuid().ToString()
        companyID            = $null
        RawContent           = $updatedHtml
        OriginalFilename     = $PageExport.Name
        ReplacedContent      = $updatedHtml
        OriginalLinks        = @($PageExport.WebUrl)
        Stub                 = $null
        ReplacedLinks        = $null
        Links                = $null
        UploadedFiles        = [System.Collections.ArrayList]@()
        ContentPreview       = Get-ArticlePreviewBlock -Title $title -docId $PageExport.PageId -Content $previewText -MaxLength $RunSummary.SetupInfo.PreviewLength
        UsingGeneratedHTML   = $false
        SuccessConverted     = $true
        CharsTrimmed         = 0
        title                = $title
        Id                   = $PageExport.PageId
        RelativePath         = "Site Pages\$title"
        RelativeFolderPath   = "$safeSiteName\Site Pages"
        Filesize             = if (Test-Path -LiteralPath $htmlPath -PathType Leaf) { (Get-Item -LiteralPath $htmlPath).Length } else { $updatedHtml.Length }
        FileTooLarge         = $false
        AllAttachments       = $attachments
        ExternalEmbeddedFiles = [System.Collections.ArrayList]@(Get-SharePointSitePageExternalImageSources -Html $updatedHtml)
        Base64EmbeddedImages = [System.Collections.ArrayList]@($attachments)
        SourceType           = "SharePointSitePage"
    }
}

$sitePagesJsonDir = $RunSummary.OutputJsonFiles.SitePagesJsonDir ?? (Join-Path $logsFolder "site-pages-json")
$sitePageAssetsDir = Join-Path $logsFolder "site-pages-assets"

if (-not (Test-Path -LiteralPath $sitePagesJsonDir -PathType Container)) {
    Set-PrintAndLog -message "SharePoint site page JSON directory not found: $sitePagesJsonDir. Run Get-SitePages.ps1 first or enable `$SharePointFetchSitePages." -Color Yellow
    return
}

$selectedSiteIds = Get-SharePointSitePageImportSelectedSiteIds
$pageJsonFiles = @(Get-ChildItem -LiteralPath $sitePagesJsonDir -Filter *.json -File)
if ($pageJsonFiles.Count -lt 1) {
    Set-PrintAndLog -message "No SharePoint site page JSON files found in $sitePagesJsonDir." -Color Yellow
    return
}

$pageDocs = [System.Collections.Generic.List[object]]::new()
foreach ($jsonFile in $pageJsonFiles) {
    try {
        $pageExport = Get-Content -LiteralPath $jsonFile.FullName -Raw | ConvertFrom-Json
    } catch {
        Set-PrintAndLog -message "Failed to read SharePoint site page JSON '$($jsonFile.FullName)': $($_.Exception.Message)" -Color Yellow
        continue
    }

    if ($selectedSiteIds.Count -gt 0 -and -not $selectedSiteIds.Contains([string]$pageExport.SiteId)) {
        continue
    }

    $resumeProbe = [PSCustomObject]@{
        SourceKey  = $pageExport.SourceKey
        SourceETag = $pageExport.SourceETag
    }
    if (
        $RunSummary.SetupInfo.ResumeFromState -and
        (Test-SharePointItemAlreadyMigrated -Item $resumeProbe -State $SharePointMigrationState -IgnoreETag:$RunSummary.SetupInfo.ResumeIgnoreETag)
    ) {
        Set-PrintAndLog -message "Skipping already completed SharePoint site page: $($pageExport.Title)" -Color DarkGray
        $RunSummary.JobInfo.ArticlesSkipped++
        continue
    }

    $pageDocs.Add((Convert-SharePointSitePageExportToConvertedDoc -PageExport $pageExport -AssetRoot $sitePageAssetsDir))
}

if ($pageDocs.Count -lt 1) {
    Set-PrintAndLog -message "No SharePoint site pages are queued for Hudu article import." -Color DarkGray
    return
}

$successConverted = @($pageDocs.ToArray())
$IndexOnlyFiles = [System.Collections.ArrayList]@()
$IndexOnlyArticles = [System.Collections.ArrayList]@()
$StubbedArticles = @()

Set-PrintAndLog -message "Importing $($successConverted.Count) SharePoint site page(s) as pre-converted Hudu article(s)." -Color Cyan

Set-IncrementedState -newState "Determine Company Designations and Folder Structure - SharePoint site pages"
. .\jobs\Make-ArticleStubs.ps1

Set-IncrementedState -newState "Populate initial data into articles - SharePoint site pages"
. .\jobs\Populate-Articles.ps1

Set-IncrementedState -newState "Upload extracted/embedded images / attachments to Hudu - SharePoint site pages"
. .\jobs\Upload-Images.ps1

Set-IncrementedState -newState "Relink Articles - SharePoint site pages"
. .\jobs\Relink-Articles.ps1

Set-PrintAndLog -message "SharePoint site page article import complete: $(@($StubbedArticles).Count) article stub(s) processed." -Color Green
