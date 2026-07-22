##### Optional job, Import fetched SharePoint site pages as Hudu articles

function Get-SharePointSitePageImportSafeFileName {
    param ([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return "site-page" }
    return (Get-SharePointSafePathName -Name $Name)
}

function Import-SharePointSitePageRendererFunctions {
    if (Get-Command ConvertTo-SharePointSitePageHtml -ErrorAction SilentlyContinue) {
        return
    }

    $fetchJobPath = if (-not [string]::IsNullOrWhiteSpace([string]$PSScriptRoot)) {
        Join-Path $PSScriptRoot "Get-SitePages.ps1"
    } else {
        $null
    }
    if ([string]::IsNullOrWhiteSpace([string]$fetchJobPath) -or -not (Test-Path -LiteralPath $fetchJobPath -PathType Leaf)) {
        $repoRoot = if (-not [string]::IsNullOrWhiteSpace([string]$workdir)) {
            [string]$workdir
        } else {
            (Get-Location).Path
        }
        $fetchJobPath = Join-Path $repoRoot "jobs\Get-SitePages.ps1"
    }
    if (-not (Test-Path -LiteralPath $fetchJobPath -PathType Leaf)) {
        return
    }

    $fetchJobScript = Get-Content -LiteralPath $fetchJobPath -Raw
    $executionBoundary = $fetchJobScript.IndexOf('$sitePagesJsonDir')
    if ($executionBoundary -le 0) {
        return
    }

    Invoke-Expression ($fetchJobScript.Substring(0, $executionBoundary))
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

function Get-SharePointSitePageImportObjectValues {
    param (
        $Object,
        [string[]]$PropertyNames
    )

    if ($null -eq $Object -or -not $Object.PSObject.Properties) { return @() }

    foreach ($propertyName in $PropertyNames) {
        $property = $Object.PSObject.Properties[$propertyName]
        if (-not $property -or $null -eq $property.Value) { continue }

        foreach ($value in @($property.Value)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                [string]$value
            }
        }
    }
}

function ConvertTo-SharePointSitePageImportWebPartHtml {
    param ($WebPart)

    if ($null -eq $WebPart) { return "" }

    if (-not [string]::IsNullOrWhiteSpace([string]$WebPart.innerHtml)) {
        return [string]$WebPart.innerHtml
    }

    $parts = [System.Collections.Generic.List[string]]::new()
    $title = $WebPart.data.title ?? $WebPart.data.description ?? $WebPart.webPartType
    if (-not [string]::IsNullOrWhiteSpace([string]$title)) {
        $parts.Add("<h2>$([System.Web.HttpUtility]::HtmlEncode([string]$title))</h2>")
    }

    $serverContent = $WebPart.data.serverProcessedContent
    foreach ($propertyName in @("htmlStrings", "searchablePlainTexts")) {
        foreach ($entry in @($serverContent.$propertyName)) {
            foreach ($value in @(Get-SharePointSitePageImportObjectValues -Object $entry -PropertyNames @("value"))) {
                if ($propertyName -eq "htmlStrings") {
                    $parts.Add($value)
                } else {
                    $parts.Add("<p>$([System.Web.HttpUtility]::HtmlEncode([string]$value))</p>")
                }
            }
        }
    }

    foreach ($entry in @($serverContent.links)) {
        $url = $entry.value
        if ([string]::IsNullOrWhiteSpace([string]$url)) { continue }
        $safeUrl = [System.Web.HttpUtility]::HtmlAttributeEncode([string]$url)
        $safeText = [System.Web.HttpUtility]::HtmlEncode([string]($entry.key ?? $url))
        $parts.Add("<p><a href=""$safeUrl"" target=""_blank"">$safeText</a></p>")
    }

    foreach ($entry in @($serverContent.imageSources)) {
        $url = $entry.value
        if ([string]::IsNullOrWhiteSpace([string]$url)) { continue }
        $safeUrl = [System.Web.HttpUtility]::HtmlAttributeEncode([string]$url)
        $safeAlt = [System.Web.HttpUtility]::HtmlAttributeEncode([string]($entry.key ?? "SharePoint image"))
        $parts.Add("<p><img src=""$safeUrl"" alt=""$safeAlt"" /></p>")
    }

    return ($parts -join "`n")
}

function Get-SharePointSitePageImportWebPartsFromObject {
    param ($Object)

    if ($null -eq $Object) { return @() }

    $found = [System.Collections.Generic.List[object]]::new()
    if ($Object -is [System.Collections.IEnumerable] -and -not ($Object -is [string]) -and -not $Object.PSObject.Properties) {
        foreach ($item in @($Object)) {
            foreach ($webPart in @(Get-SharePointSitePageImportWebPartsFromObject -Object $item)) {
                $found.Add($webPart)
            }
        }
        return @($found)
    }

    if (-not $Object.PSObject.Properties) { return @() }

    if ($Object.PSObject.Properties["innerHtml"] -or $Object.PSObject.Properties["webPartType"] -or $Object.PSObject.Properties["data"]) {
        $found.Add($Object)
    }

    foreach ($property in @($Object.PSObject.Properties)) {
        if ($property.Name -like "@odata*") { continue }
        if ($null -eq $property.Value) { continue }
        if ($property.Value -is [string] -or $property.Value -is [ValueType]) { continue }

        foreach ($webPart in @(Get-SharePointSitePageImportWebPartsFromObject -Object $property.Value)) {
            $found.Add($webPart)
        }
    }

    return @($found)
}

function ConvertTo-SharePointSitePageImportHtml {
    param ($PageExport)

    $pageObject = $PageExport.Page ?? $PageExport
    $webPartProperty = if ($PageExport.PSObject.Properties['WebParts']) {
        $PageExport.PSObject.Properties['WebParts'].Value
    } else {
        $null
    }
    $webParts = if ($null -ne $webPartProperty -and @($webPartProperty).Count -gt 0) {
        @($webPartProperty)
    } elseif ($PageExport.Page -and $PageExport.Page.canvasLayout) {
        @(Get-SharePointSitePageImportWebPartsFromObject -Object $PageExport.Page.canvasLayout)
    } else {
        @()
    }

    if (@($webParts).Count -lt 1) { return $null }

    $title = [string]($PageExport.Title ?? $pageObject.title ?? $pageObject.name ?? "Untitled SharePoint page")
    $siteName = [string]($PageExport.SiteName ?? $PageExport.SiteId ?? "SharePoint Site")
    $safeTitle = [System.Web.HttpUtility]::HtmlEncode($title)
    $safeSite = [System.Web.HttpUtility]::HtmlEncode($siteName)
    $safeUrl = [System.Web.HttpUtility]::HtmlAttributeEncode([string]($PageExport.WebUrl ?? $pageObject.webUrl))
    $safeModified = [System.Web.HttpUtility]::HtmlEncode([string]($PageExport.LastModifiedDateTime ?? $pageObject.lastModifiedDateTime))

    $bodyParts = [System.Collections.Generic.List[string]]::new()
    foreach ($webPart in @($webParts)) {
        $partHtml = ConvertTo-SharePointSitePageImportWebPartHtml -WebPart $webPart
        if (-not [string]::IsNullOrWhiteSpace([string]$partHtml)) {
            $bodyParts.Add($partHtml)
        }
    }

    if ($bodyParts.Count -lt 1) { return $null }

@"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>$safeTitle</title>
</head>
<body>
  <h1>$safeTitle</h1>
  <div class="meta">
    <div><strong>SharePoint site:</strong> $safeSite</div>
    <div><strong>Last modified:</strong> $safeModified</div>
    <div><strong>Source:</strong> <a href="$safeUrl" target="_blank">$safeUrl</a></div>
  </div>
  $($bodyParts -join "`n")
</body>
</html>
"@
}

function Convert-SharePointSitePageExportToConvertedDoc {
    param (
        [Parameter(Mandatory)]
        $PageExport,

        [Parameter(Mandatory)]
        [string]$AssetRoot
    )

    Import-SharePointSitePageRendererFunctions

    $htmlPath = [string]$PageExport.HtmlPath
    $html = $null
    if (Get-Command ConvertTo-SharePointSitePageHtml -ErrorAction SilentlyContinue) {
        $pageObject = $PageExport.Page ?? $PageExport
        $webPartProperty = if ($PageExport.PSObject.Properties['WebParts']) {
            $PageExport.PSObject.Properties['WebParts'].Value
        } else {
            $null
        }
        $webParts = if ($null -ne $webPartProperty -and @($webPartProperty).Count -gt 0) {
            @($webPartProperty)
        } elseif ($PageExport.Page -and $PageExport.Page.canvasLayout -and (Get-Command Get-SharePointSitePageWebPartsFromCanvasObject -ErrorAction SilentlyContinue)) {
            @(Get-SharePointSitePageWebPartsFromCanvasObject -Object $PageExport.Page.canvasLayout)
        } else {
            @()
        }

        if (@($webParts).Count -gt 0) {
            $siteObject = [PSCustomObject]@{
                displayName = $PageExport.SiteName
                name        = $PageExport.SiteName
                id          = $PageExport.SiteId
            }
            $html = ConvertTo-SharePointSitePageHtml -Site $siteObject -Page $pageObject -WebParts $webParts
        }
    }

    if ([string]::IsNullOrWhiteSpace($html)) {
        $html = ConvertTo-SharePointSitePageImportHtml -PageExport $PageExport
    }

    if ([string]::IsNullOrWhiteSpace($html)) {
        $html = [string]$PageExport.Html
    }
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
