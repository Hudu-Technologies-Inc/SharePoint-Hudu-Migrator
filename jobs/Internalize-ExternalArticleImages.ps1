##### Optional maintenance job, download external article images into Hudu and rewrite img src values

$internalizeDryRun = [bool]($HuduInternalizeExternalArticleImagesDryRun ?? $true)
$internalizeUsePublicPhotos = [bool]($HuduInternalizeExternalArticleImagesUsePublicPhotos ?? $true)
$internalizeScrubUnexpectedLocalSources = [bool]($HuduInternalizeExternalArticleImagesScrubUnexpectedLocalSources ?? $false)
$internalizePreferExistingHuduImages = [bool]($HuduInternalizeExternalArticleImagesPreferExistingHuduImages ?? $true)
$internalizeRewriteUnexpectedLocalExisting = [bool]($HuduInternalizeExternalArticleImagesRewriteUnexpectedLocalExisting ?? $false)
$internalizeProbeDownloads = [bool]($HuduInternalizeExternalArticleImagesProbeDownloads ?? $false)
$internalizeMaxArticles = [int]($HuduInternalizeExternalArticleImagesMaxArticles ?? 0)
$internalizeMaxImagesPerArticle = [int]($HuduInternalizeExternalArticleImagesMaxImagesPerArticle ?? 0)
$internalizeArticleIds = @(
    if ($null -ne $HuduInternalizeExternalArticleImagesArticleIds) {
        @($HuduInternalizeExternalArticleImagesArticleIds | ForEach-Object { [string]$_ })
    } else {
        @()
    }
)
$internalizeCompanyIds = @(
    if ($null -ne $HuduInternalizeExternalArticleImagesCompanyIds) {
        @($HuduInternalizeExternalArticleImagesCompanyIds | ForEach-Object { [string]$_ })
    } else {
        @()
    }
)

$internalizeLogRoot = if (-not [string]::IsNullOrWhiteSpace([string]$logsFolder)) {
    [string]$logsFolder
} elseif ($RunSummary -and $RunSummary.OutputJsonFiles -and $RunSummary.OutputJsonFiles.SummaryPath) {
    Split-Path -Parent $RunSummary.OutputJsonFiles.SummaryPath
} else {
    Join-Path (Get-Location).Path "logs"
}

$internalizeOutputDir = Join-Path $internalizeLogRoot "internalized-external-images"
$internalizeDownloadDir = Join-Path $internalizeOutputDir "downloads"
$internalizeReportPath = Join-Path $internalizeOutputDir "internalized-external-images.csv"

foreach ($folder in @($internalizeOutputDir, $internalizeDownloadDir)) {
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        $null = New-Item -ItemType Directory -Path $folder -Force
    }
}

function Get-HuduInternalizeBaseUri {
    $base = [string]($HuduBaseURL ?? $HuduBaseUrl)

    if ([string]::IsNullOrWhiteSpace($base) -and (Get-Command Get-HuduBaseURL -ErrorAction SilentlyContinue)) {
        $base = [string](Get-HuduBaseURL)
    }

    if ([string]::IsNullOrWhiteSpace($base)) {
        $base = [string]$RunSummary.SetupInfo.HuduDestination
    }

    if ([string]::IsNullOrWhiteSpace($base)) {
        throw "Hudu base URL is not available. Set `$HuduBaseURL before running this job."
    }

    [uri]($base.TrimEnd('/'))
}

function Get-HuduInternalizeArticleId {
    param ($Article)

    return (
        $Article.id ??
        $Article.Id ??
        $Article.article_id ??
        $Article.ArticleId ??
        $Article.ArticleID ??
        $Article.article.id ??
        $Article.article.Id
    )
}

function Get-HuduInternalizeArticleName {
    param ($Article)

    return [string](
        $Article.name ??
        $Article.Name ??
        $Article.title ??
        $Article.Title ??
        $Article.article.name ??
        $Article.article.Name ??
        $Article.article.title ??
        $Article.article.Title ??
        "Untitled Article"
    )
}

function Get-HuduInternalizeArticleCompanyId {
    param ($Article)

    return (
        $Article.company_id ??
        $Article.companyId ??
        $Article.CompanyId ??
        $Article.company.id ??
        $Article.article.company_id ??
        $Article.article.companyId ??
        $Article.article.CompanyId ??
        $Article.article.company.id
    )
}

function Get-HuduInternalizeArticleContent {
    param ($Article)

    return [string](
        $Article.content ??
        $Article.Content ??
        $Article.contents ??
        $Article.Contents ??
        $Article.body ??
        $Article.Body ??
        $Article.article.content ??
        $Article.article.Content ??
        $Article.article.contents ??
        $Article.article.Contents ??
        $Article.article.body ??
        $Article.article.Body ??
        ""
    )
}

function Expand-HuduInternalizeArticles {
    param ($InputObject)

    $expanded = [System.Collections.Generic.List[object]]::new()

    foreach ($item in @($InputObject)) {
        $articleSet = $item.articles ?? $item.Articles

        if ($null -ne $articleSet) {
            foreach ($wrappedArticle in @($articleSet)) {
                $expanded.Add(($wrappedArticle.article ?? $wrappedArticle.Article ?? $wrappedArticle))
            }
            continue
        }

        $expanded.Add(($item.article ?? $item.Article ?? $item))
    }

    return @($expanded)
}

function Test-HuduInternalizeExternalImageSource {
    param (
        [Parameter(Mandatory)] [string]$Source,
        [Parameter(Mandatory)] [uri]$HuduBaseUri
    )

    if ([string]::IsNullOrWhiteSpace($Source)) { return $false }
    if ($Source -match '(?i)^(data|cid|blob|mailto|tel):') { return $false }
    if ($Source -notmatch '(?i)^https?://') { return $false }

    try {
        $sourceUri = [uri]$Source
    } catch {
        return $false
    }

    return (-not ([string]::Equals($sourceUri.Host, $HuduBaseUri.Host, [System.StringComparison]::OrdinalIgnoreCase)))
}

function Test-HuduInternalizeExpectedHuduImageSource {
    param (
        [Parameter(Mandatory)] [string]$Source,
        [Parameter(Mandatory)] [uri]$HuduBaseUri
    )

    if ([string]::IsNullOrWhiteSpace($Source)) { return $false }

    $expectedPathPattern = '(?i)^/?(public_photos?|uploads?|files?|file|photos?|photo)(/|\?|$)'
    if ($Source -match $expectedPathPattern) { return $true }

    if ($Source -notmatch '(?i)^https?://') { return $false }

    try {
        $sourceUri = [uri]$Source
    } catch {
        return $false
    }

    if (-not [string]::Equals($sourceUri.Host, $HuduBaseUri.Host, [System.StringComparison]::OrdinalIgnoreCase)) {
        if ($sourceUri.Host -notmatch '(?i)(^|\.)huducloud\.com$') {
            return $false
        }
    }

    return ($sourceUri.AbsolutePath -match $expectedPathPattern)
}

function Get-HuduInternalizeImageSourceRecords {
    param (
        [Parameter(Mandatory)] [string]$Html,
        [Parameter(Mandatory)] [uri]$HuduBaseUri
    )

    $records = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $matches = [regex]::Matches($Html, '(?is)<img\b[^>]*?\bsrc\s*=\s*(?<quote>["''])(?<src>.*?)(\k<quote>)[^>]*>')

    foreach ($match in $matches) {
        $source = [System.Web.HttpUtility]::HtmlDecode([string]$match.Groups['src'].Value)
        if (-not $seen.Add($source)) { continue }

        $classification = if ([string]::IsNullOrWhiteSpace($source)) {
            "Empty"
        } elseif ($source -match '(?i)^(data|cid|blob|mailto|tel):') {
            "IgnoredSpecial"
        } elseif (Test-HuduInternalizeExpectedHuduImageSource -Source $source -HuduBaseUri $HuduBaseUri) {
            "ExpectedHudu"
        } elseif (Test-HuduInternalizeExternalImageSource -Source $source -HuduBaseUri $HuduBaseUri) {
            "External"
        } elseif ($source -notmatch '(?i)^https?://') {
            "UnexpectedLocal"
        } else {
            "UnexpectedAbsolute"
        }

        $records.Add([PSCustomObject]@{
            Source         = $source
            Classification = $classification
        })
    }

    return @($records)
}

function Get-HuduInternalizeExternalImageSources {
    param (
        [Parameter(Mandatory)] [string]$Html,
        [Parameter(Mandatory)] [uri]$HuduBaseUri
    )

    @(
        Get-HuduInternalizeImageSourceRecords -Html $Html -HuduBaseUri $HuduBaseUri |
            Where-Object { $_.Classification -eq "External" } |
            ForEach-Object { $_.Source }
    )
}

function Get-HuduInternalizeExtensionFromContentType {
    param ([string]$ContentType)

    switch -Regex ($ContentType) {
        '^image/jpeg' { return '.jpg' }
        '^image/png'  { return '.png' }
        '^image/gif'  { return '.gif' }
        '^image/webp' { return '.webp' }
        '^image/svg\+xml' { return '.svg' }
        '^image/bmp'  { return '.bmp' }
        '^image/x-icon' { return '.ico' }
        default { return $null }
    }
}

function Get-HuduInternalizeSafeDownloadName {
    param (
        [Parameter(Mandatory)] [string]$Url,
        [Parameter(Mandatory)] [int]$Index,
        [string]$ContentType
    )

    $extension = $null
    try {
        $uri = [uri]$Url
        $fileName = [System.IO.Path]::GetFileName($uri.AbsolutePath)
        $extension = [System.IO.Path]::GetExtension($fileName)
    } catch {
        $fileName = $null
    }

    if ([string]::IsNullOrWhiteSpace($extension)) {
        $extension = Get-HuduInternalizeExtensionFromContentType -ContentType $ContentType
    }
    if ([string]::IsNullOrWhiteSpace($extension)) {
        $extension = ".img"
    }

    $safeStem = if (-not [string]::IsNullOrWhiteSpace($fileName)) {
        [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    } else {
        "external-image-$Index"
    }
    $safeStem = Get-SharePointSafePathName -Name $safeStem
    if ([string]::IsNullOrWhiteSpace($safeStem)) {
        $safeStem = "external-image-$Index"
    }

    "{0}-{1}{2}" -f $safeStem, $Index, $extension
}

function Get-HuduInternalizeSourceFileName {
    param ([string]$Source)

    if ([string]::IsNullOrWhiteSpace($Source)) { return $null }

    try {
        $decoded = [System.Web.HttpUtility]::HtmlDecode($Source)
        if ($decoded -match '(?i)^https?://') {
            $decoded = ([uri]$decoded).AbsolutePath
        } else {
            $decoded = ($decoded -split '[?#]', 2)[0]
        }

        $fileName = [System.IO.Path]::GetFileName($decoded)
        if ([string]::IsNullOrWhiteSpace($fileName)) { return $null }

        return [uri]::UnescapeDataString($fileName)
    } catch {
        return $null
    }
}

function Add-HuduInternalizeImageIndexEntry {
    param (
        [Parameter(Mandatory)] [hashtable]$Index,
        [Parameter(Mandatory)] [string]$FileName,
        [Parameter(Mandatory)] [string]$Url,
        [string]$Kind,
        $Id,
        $ArticleId
    )

    if ([string]::IsNullOrWhiteSpace($FileName) -or [string]::IsNullOrWhiteSpace($Url)) { return }

    $key = $FileName.Trim().ToLowerInvariant()
    if (-not $Index.ContainsKey($key)) {
        $Index[$key] = [System.Collections.Generic.List[object]]::new()
    }

    $Index[$key].Add([PSCustomObject]@{
        FileName  = $FileName
        Url       = $Url
        Kind      = $Kind
        Id        = $Id
        ArticleId = $ArticleId
    })
}

function New-HuduInternalizeExistingImageIndex {
    $index = @{}

    if (Get-Command Get-HuduPublicPhotos -ErrorAction SilentlyContinue) {
        try {
            foreach ($photo in @(Get-HuduPublicPhotos)) {
                Add-HuduInternalizeImageIndexEntry `
                    -Index $index `
                    -FileName ([string]($photo.file_name ?? $photo.FileName ?? $photo.name ?? $photo.Name)) `
                    -Url ([string]($photo.url ?? $photo.Url)) `
                    -Kind "PublicPhoto" `
                    -Id ($photo.id ?? $photo.Id ?? $photo.numeric_id ?? $photo.NumericId) `
                    -ArticleId ($photo.record_id ?? $photo.RecordId)
            }
        } catch {
            Set-PrintAndLog -message "Unable to build public photo reuse index: $($_.Exception.Message)" -Color Yellow
        }
    }

    if (Get-Command Get-HuduUploads -ErrorAction SilentlyContinue) {
        try {
            foreach ($upload in @(Get-HuduUploads)) {
                Add-HuduInternalizeImageIndexEntry `
                    -Index $index `
                    -FileName ([string]($upload.name ?? $upload.Name ?? $upload.file_name ?? $upload.FileName)) `
                    -Url ([string]($upload.url ?? $upload.Url)) `
                    -Kind "Upload" `
                    -Id ($upload.id ?? $upload.Id) `
                    -ArticleId ($upload.uploadable_id ?? $upload.UploadableId ?? $upload.record_id ?? $upload.RecordId)
            }
        } catch {
            Set-PrintAndLog -message "Unable to build upload reuse index: $($_.Exception.Message)" -Color Yellow
        }
    }

    return $index
}

function Find-HuduInternalizeExistingImage {
    param (
        [Parameter(Mandatory)] [hashtable]$Index,
        [Parameter(Mandatory)] [string]$Source,
        $ArticleId
    )

    $fileName = Get-HuduInternalizeSourceFileName -Source $Source
    if ([string]::IsNullOrWhiteSpace($fileName)) { return $null }

    $key = $fileName.Trim().ToLowerInvariant()
    if (-not $Index.ContainsKey($key)) { return $null }

    $matches = @($Index[$key])
    $sameArticleMatch = @(
        $matches |
            Where-Object {
                $ArticleId -and $_.ArticleId -and
                ([string]$_.ArticleId -eq [string]$ArticleId)
            } |
            Select-Object -First 1
    )
    if ($sameArticleMatch.Count -gt 0) { return $sameArticleMatch[0] }

    return ($matches | Select-Object -First 1)
}

function Test-HuduInternalizeExternalImageDownload {
    param ([Parameter(Mandatory)] [string]$Url)

    $client = [System.Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromSeconds(20)
    $client.DefaultRequestHeaders.UserAgent.ParseAdd("SharePoint-Hudu-Migrator/1.0")

    try {
        foreach ($methodName in @("HEAD", "GET")) {
            $request = if ($methodName -eq "HEAD") {
                [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Head, $Url)
            } else {
                $getRequest = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $Url)
                $getRequest.Headers.Range = [System.Net.Http.Headers.RangeHeaderValue]::new(0, 0)
                $getRequest
            }

            try {
                $response = $client.SendAsync($request, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
                try {
                    $statusCode = [int]$response.StatusCode
                    $contentType = [string]$response.Content.Headers.ContentType
                    $contentLength = $response.Content.Headers.ContentLength

                    if ($response.IsSuccessStatusCode -and $contentType -match '^image/') {
                        return [PSCustomObject]@{
                            Reachable     = $true
                            Method        = $methodName
                            StatusCode    = $statusCode
                            ContentType   = $contentType
                            ContentLength = $contentLength
                            Error         = $null
                        }
                    }

                    if ($methodName -eq "GET") {
                        return [PSCustomObject]@{
                            Reachable     = $false
                            Method        = $methodName
                            StatusCode    = $statusCode
                            ContentType   = $contentType
                            ContentLength = $contentLength
                            Error         = if ($response.IsSuccessStatusCode) { "Response content-type was not image/*." } else { "HTTP $statusCode" }
                        }
                    }
                } finally {
                    $response.Dispose()
                }
            } catch {
                if ($methodName -eq "GET") {
                    return [PSCustomObject]@{
                        Reachable     = $false
                        Method        = $methodName
                        StatusCode    = $null
                        ContentType   = $null
                        ContentLength = $null
                        Error         = $_.Exception.Message
                    }
                }
            } finally {
                $request.Dispose()
            }
        }
    } finally {
        $client.Dispose()
    }
}

function Save-HuduInternalizeExternalImage {
    param (
        [Parameter(Mandatory)] [string]$Url,
        [Parameter(Mandatory)] [string]$Directory,
        [Parameter(Mandatory)] [int]$Index
    )

    $client = [System.Net.Http.HttpClient]::new()
    $client.DefaultRequestHeaders.UserAgent.ParseAdd("SharePoint-Hudu-Migrator/1.0")
    $response = $client.GetAsync($Url).GetAwaiter().GetResult()
    $response.EnsureSuccessStatusCode() | Out-Null
    $contentType = [string]$response.Content.Headers.ContentType
    if ($contentType -and $contentType -notmatch '^image/') {
        throw "URL did not return an image content-type: $contentType"
    }

    $fileName = Get-HuduInternalizeSafeDownloadName -Url $Url -Index $Index -ContentType $contentType
    $filePath = Join-Path $Directory $fileName

    $stream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
    try {
        $fileStream = [System.IO.File]::Create($filePath)
        try {
            $stream.CopyTo($fileStream)
        } finally {
            $fileStream.Dispose()
        }
    } finally {
        $stream.Dispose()
        $response.Dispose()
        $client.Dispose()
    }

    [PSCustomObject]@{
        FilePath    = $filePath
        ContentType = $contentType
    }
}

function New-HuduInternalizedArticleImageUpload {
    param (
        [Parameter(Mandatory)] [string]$FilePath,
        [Parameter(Mandatory)] $ArticleId,
        [switch]$UsePublicPhotos
    )

    $extension = [System.IO.Path]::GetExtension($FilePath).TrimStart('.').ToLowerInvariant()

    if ($UsePublicPhotos -and $extension -match '^(jpg|jpeg|png|webp)$' -and (Get-Command New-HuduPublicPhoto -ErrorAction SilentlyContinue)) {
        $upload = New-HuduPublicPhoto -FilePath $FilePath -record_id $ArticleId -record_type 'Article'
        return ($upload.public_photo ?? $upload)
    }

    $upload = New-HuduUpload -FilePath $FilePath -record_id $ArticleId -record_type 'Article'
    return ($upload.upload ?? $upload)
}

function Get-HuduInternalizedImageUrl {
    param (
        $Upload,
        [Parameter(Mandatory)] [uri]$HuduBaseUri,
        [string]$FilePath
    )

    $url = $Upload.url ?? $Upload.Url
    if (-not [string]::IsNullOrWhiteSpace([string]$url)) {
        return [string]$url
    }

    $id = $Upload.id ?? $Upload.Id
    if (-not $id) { return $null }

    $extension = if ($FilePath) { [System.IO.Path]::GetExtension($FilePath).TrimStart('.').ToLowerInvariant() } else { "" }
    if ($extension -match '^(jpg|jpeg|png|webp)$') {
        return "$($HuduBaseUri.ToString().TrimEnd('/'))/public_photo/$id"
    }

    return "$($HuduBaseUri.ToString().TrimEnd('/'))/file/$id"
}

function Replace-HuduArticleImageSource {
    param (
        [Parameter(Mandatory)] [string]$Html,
        [Parameter(Mandatory)] [string]$OldSource,
        [Parameter(Mandatory)] [string]$NewSource
    )

    $pattern = '(?is)(<img\b[^>]*?\bsrc\s*=\s*)(?<quote>["''])' + [regex]::Escape($OldSource) + '(\k<quote>)'
    [regex]::Replace($Html, $pattern, {
        param($Match)
        $quote = $Match.Groups['quote'].Value
        return "$($Match.Groups[1].Value)$quote$NewSource$quote"
    })
}

function Remove-HuduArticleImageTagBySource {
    param (
        [Parameter(Mandatory)] [string]$Html,
        [Parameter(Mandatory)] [string]$Source
    )

    $pattern = '(?is)<img\b(?=[^>]*?\bsrc\s*=\s*(?<quote>["''])' + [regex]::Escape($Source) + '(\k<quote>))[^>]*>'
    [regex]::Replace($Html, $pattern, '')
}

$huduBaseUri = Get-HuduInternalizeBaseUri
$existingHuduImageIndex = if ($internalizePreferExistingHuduImages) {
    Set-PrintAndLog -message "Building existing Hudu image reuse index from uploads/public photos." -Color Cyan
    New-HuduInternalizeExistingImageIndex
} else {
    @{}
}

$allArticlesResponse = Get-HuduArticles
$articles = @(Expand-HuduInternalizeArticles -InputObject $allArticlesResponse)

if ($internalizeArticleIds.Count -gt 0) {
    $articles = @($articles | Where-Object { $internalizeArticleIds -contains [string](Get-HuduInternalizeArticleId $_) })
}
if ($internalizeCompanyIds.Count -gt 0) {
    $articles = @($articles | Where-Object { $internalizeCompanyIds -contains [string](Get-HuduInternalizeArticleCompanyId $_) })
}
if ($internalizeMaxArticles -gt 0) {
    $articles = @($articles | Select-Object -First $internalizeMaxArticles)
}

Set-PrintAndLog -message "Scanning $($articles.Count) Hudu article(s) for external/unexpected image sources. DryRun=$internalizeDryRun; ProbeDownloads=$internalizeProbeDownloads; ScrubUnexpectedLocalSources=$internalizeScrubUnexpectedLocalSources." -Color Cyan

$report = [System.Collections.Generic.List[object]]::new()
$articleIndex = 0
$articlesWithContent = 0
$articlesWithImageTags = 0
$externalSourceCount = 0
$unexpectedSourceCount = 0
$emptyContentCount = 0
$missingArticleIdCount = 0
$existingImageReuseCandidateCount = 0
$existingImageReuseCount = 0
$downloadProbeSuccessCount = 0
$downloadProbeFailureCount = 0

foreach ($article in @($articles)) {
    $articleIndex++
    $articleId = Get-HuduInternalizeArticleId -Article $article
    $articleName = Get-HuduInternalizeArticleName -Article $article
    $companyId = Get-HuduInternalizeArticleCompanyId -Article $article
    $content = Get-HuduInternalizeArticleContent -Article $article

    if (-not $articleId) {
        $missingArticleIdCount++
    }

    if ([string]::IsNullOrWhiteSpace($content)) {
        $emptyContentCount++
        continue
    }

    $articlesWithContent++

    $imageSourceRecords = @(Get-HuduInternalizeImageSourceRecords -Html $content -HuduBaseUri $huduBaseUri)
    if ($imageSourceRecords.Count -gt 0) {
        $articlesWithImageTags++
    }

    $unexpectedLocalSources = @($imageSourceRecords | Where-Object { $_.Classification -in @("UnexpectedLocal", "UnexpectedAbsolute") })
    $sources = @($imageSourceRecords | Where-Object { $_.Classification -eq "External" } | ForEach-Object { $_.Source })
    if ($internalizeMaxImagesPerArticle -gt 0) {
        $sources = @($sources | Select-Object -First $internalizeMaxImagesPerArticle)
    }
    $externalSourceCount += $sources.Count
    $unexpectedSourceCount += $unexpectedLocalSources.Count
    if ($sources.Count -lt 1 -and $unexpectedLocalSources.Count -lt 1) { continue }
    if (-not $articleId -and -not $internalizeDryRun) {
        Set-PrintAndLog -message "Skipping article '$articleName' because it has image sources but no article ID was available." -Color Yellow
        continue
    }

    if ($sources.Count -eq 0 -and $unexpectedSourceCount -eq 0) {
        Set-PrintAndLog -message "Article $articleIndex/$($articles.Count) '$articleName': No external or unexpected image sources found." -Color Gray
    } else {
        Set-PrintAndLog -message "Article $articleIndex/$($articles.Count) '$articleName': $($sources.Count) external image source(s), $($unexpectedLocalSources.Count) unexpected local/absolute source(s)." -Color Yellow
    }

    $updatedContent = [string]$content
    $articleChanged = $false
    $imageIndex = 0

    foreach ($unexpectedSourceRecord in $unexpectedLocalSources) {
        $record = [ordered]@{
            ArticleId      = $articleId
            ArticleName    = $articleName
            CompanyId      = $companyId
            OldSource      = $unexpectedSourceRecord.Source
            LocalPath      = $null
            NewSource      = $null
            UploadId       = $null
            Classification = $unexpectedSourceRecord.Classification
            Status         = if ($internalizeScrubUnexpectedLocalSources -and -not $internalizeDryRun) { "Scrubbed" } else { "Reported" }
            ExistingKind   = $null
            ProbeMethod    = $null
            ProbeStatusCode = $null
            ProbeContentType = $null
            ProbeContentLength = $null
            Error          = $null
        }

        $existingImage = $null
        if ($internalizePreferExistingHuduImages -and $internalizeRewriteUnexpectedLocalExisting) {
            $existingImage = Find-HuduInternalizeExistingImage -Index $existingHuduImageIndex -Source $unexpectedSourceRecord.Source -ArticleId $articleId
        }

        if ($existingImage) {
            $existingImageReuseCandidateCount++
            $record.NewSource = $existingImage.Url
            $record.UploadId = $existingImage.Id
            $record.ExistingKind = $existingImage.Kind
            $record.Status = if ($internalizeDryRun) { "DryRunReuseExisting" } else { "ReusedExisting" }
            if (-not $internalizeDryRun) {
                $updatedContent = Replace-HuduArticleImageSource -Html $updatedContent -OldSource $unexpectedSourceRecord.Source -NewSource $existingImage.Url
                $articleChanged = $true
                $existingImageReuseCount++
            }
        } elseif ($internalizeScrubUnexpectedLocalSources -and -not $internalizeDryRun) {
            $updatedContent = Remove-HuduArticleImageTagBySource -Html $updatedContent -Source $unexpectedSourceRecord.Source
            $articleChanged = $true
        }

        $report.Add([PSCustomObject]$record)
    }

    foreach ($source in $sources) {
        $imageIndex++
        $record = [ordered]@{
            ArticleId    = $articleId
            ArticleName  = $articleName
            CompanyId    = $companyId
            OldSource    = $source
            LocalPath    = $null
            NewSource    = $null
            UploadId     = $null
            Classification = "External"
            Status       = if ($internalizeDryRun) { "DryRun" } else { "Pending" }
            ExistingKind = $null
            ProbeMethod  = $null
            ProbeStatusCode = $null
            ProbeContentType = $null
            ProbeContentLength = $null
            Error        = $null
        }

        $existingImage = $null
        if ($internalizePreferExistingHuduImages) {
            $existingImage = Find-HuduInternalizeExistingImage -Index $existingHuduImageIndex -Source $source -ArticleId $articleId
        }

        if ($existingImage) {
            $existingImageReuseCandidateCount++
            $record.NewSource = $existingImage.Url
            $record.UploadId = $existingImage.Id
            $record.ExistingKind = $existingImage.Kind
            $record.Status = if ($internalizeDryRun) { "DryRunReuseExisting" } else { "ReusedExisting" }
            if (-not $internalizeDryRun) {
                $updatedContent = Replace-HuduArticleImageSource -Html $updatedContent -OldSource $source -NewSource $existingImage.Url
                $articleChanged = $true
                $existingImageReuseCount++
            }
        } elseif ($internalizeDryRun -and $internalizeProbeDownloads) {
            $probe = Test-HuduInternalizeExternalImageDownload -Url $source
            $record.ProbeMethod = $probe.Method
            $record.ProbeStatusCode = $probe.StatusCode
            $record.ProbeContentType = $probe.ContentType
            $record.ProbeContentLength = $probe.ContentLength

            if ($probe.Reachable) {
                $record.Status = "DryRunDownloadable"
                $downloadProbeSuccessCount++
            } else {
                $record.Status = "DryRunDownloadFailed"
                $record.Error = $probe.Error
                $downloadProbeFailureCount++
            }
        } elseif (-not $internalizeDryRun) {
            try {
                $download = Save-HuduInternalizeExternalImage -Url $source -Directory $internalizeDownloadDir -Index $imageIndex
                $upload = New-HuduInternalizedArticleImageUpload `
                    -FilePath $download.FilePath `
                    -ArticleId $articleId `
                    -UsePublicPhotos:$internalizeUsePublicPhotos
                $newSource = Get-HuduInternalizedImageUrl -Upload $upload -HuduBaseUri $huduBaseUri -FilePath $download.FilePath

                if ([string]::IsNullOrWhiteSpace([string]$newSource)) {
                    throw "Hudu upload did not return a usable URL."
                }

                $updatedContent = Replace-HuduArticleImageSource -Html $updatedContent -OldSource $source -NewSource $newSource
                $articleChanged = $true

                $record.LocalPath = $download.FilePath
                $record.NewSource = $newSource
                $record.UploadId = $upload.id ?? $upload.Id
                $record.Status = "Uploaded"
            } catch {
                $record.Status = "Failed"
                $record.Error = $_.Exception.Message
                Set-PrintAndLog -message "Failed to internalize image for article '$articleName': $source :: $($_.Exception.Message)" -Color Red
            }
        }

        $report.Add([PSCustomObject]$record)
    }

    if (-not $internalizeDryRun -and $articleChanged) {
        try {
            if ($companyId -and [int]$companyId -gt 0) {
                Set-HuduArticle -ArticleId $articleId -Content $updatedContent -name $articleName -CompanyId $companyId | Out-Null
            } else {
                Set-HuduArticle -ArticleId $articleId -Content $updatedContent -name $articleName | Out-Null
            }
            Set-PrintAndLog -message "Updated Hudu article '$articleName' ($articleId) with internalized image source(s)." -Color Green
        } catch {
            Set-PrintAndLog -message "Failed to update Hudu article '$articleName' ($articleId): $($_.Exception.Message)" -Color Red
            $report.Add([PSCustomObject]@{
                ArticleId   = $articleId
                ArticleName = $articleName
                CompanyId   = $companyId
                OldSource   = $null
                LocalPath   = $null
                NewSource   = $null
                UploadId    = $null
                Classification = "Article"
                Status      = "ArticleUpdateFailed"
                ExistingKind = $null
                ProbeMethod = $null
                ProbeStatusCode = $null
                ProbeContentType = $null
                ProbeContentLength = $null
                Error       = $_.Exception.Message
            })
        }
    }
}

$report |
    Export-Csv -LiteralPath $internalizeReportPath -NoTypeInformation -Encoding UTF8

Set-PrintAndLog -message "External image internalization scan complete: $($report.Count) image reference(s). Report: $internalizeReportPath" -Color Green
Set-PrintAndLog -message "Scan summary: $articlesWithContent article(s) had content; $emptyContentCount article(s) had empty content; $missingArticleIdCount article(s) had no detected ID; $articlesWithImageTags article(s) had img tag(s); $externalSourceCount external source(s); $unexpectedSourceCount unexpected local/absolute source(s); $existingImageReuseCandidateCount existing Hudu image reuse candidate(s); $existingImageReuseCount existing Hudu image source(s) reused; $downloadProbeSuccessCount dry-run download probe(s) passed; $downloadProbeFailureCount failed." -Color Cyan
