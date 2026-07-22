##### Optional job, Fetch SharePoint site pages as HTML snapshots

function ConvertTo-SharePointSitePageHtmlText {
    param ($Value)

    if ($null -eq $Value) { return "" }
    return [System.Web.HttpUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-SharePointSitePageHtmlAttribute {
    param ($Value)

    if ($null -eq $Value) { return "" }
    return [System.Web.HttpUtility]::HtmlAttributeEncode([string]$Value)
}

function Get-SharePointSitePageSafeFileBaseName {
    param (
        $Site,
        $Page
    )

    $siteName = Get-SharePointSafePathName -Name ($Site.displayName ?? $Site.name ?? $Site.id ?? "site")
    $pageName = Get-SharePointSafePathName -Name ($Page.title ?? $Page.name ?? $Page.id ?? "page")
    $pageId = Get-SharePointSafePathName -Name ($Page.id ?? ([guid]::NewGuid().ToString()))
    return "$siteName-$pageName-$pageId"
}

function Invoke-SharePointSitePageRequest {
    param (
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $headers = @{} + (Update-SharePointAccessToken)
    $headers["Accept"] = "application/json;odata.metadata=none"

    Invoke-RestMethod `
        -Method Get `
        -Uri $Uri `
        -Headers $headers `
        -ErrorAction Stop
}

function Get-SharePointSitePageCollection {
    param (
        [Parameter(Mandatory)]
        [string]$SiteId
    )

    $pages = [System.Collections.ArrayList]@()
    $nextUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/pages/microsoft.graph.sitePage?`$select=id,name,title,webUrl,eTag,createdDateTime,lastModifiedDateTime,description,pageLayout,promotionKind"

    while ($nextUri) {
        $response = Invoke-SharePointSitePageRequest -Uri $nextUri
        if ($response.value) {
            [void]$pages.AddRange(@($response.value))
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return @($pages)
}

function Get-SharePointSitePageWithCanvas {
    param (
        [Parameter(Mandatory)]
        [string]$SiteId,

        [Parameter(Mandatory)]
        [string]$PageId
    )

    $uri = "https://graph.microsoft.com/v1.0/sites/$SiteId/pages/$PageId/microsoft.graph.sitePage?`$expand=canvasLayout"
    Invoke-SharePointSitePageRequest -Uri $uri
}

function Get-SharePointSitePageWebParts {
    param (
        [Parameter(Mandatory)]
        [string]$SiteId,

        [Parameter(Mandatory)]
        [string]$PageId
    )

    $webParts = [System.Collections.ArrayList]@()
    $nextUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/pages/$PageId/microsoft.graph.sitePage/webparts"

    while ($nextUri) {
        $response = Invoke-SharePointSitePageRequest -Uri $nextUri
        if ($response.value) {
            [void]$webParts.AddRange(@($response.value))
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return @($webParts)
}

function Get-SharePointSitePageObjectValues {
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

function ConvertTo-SharePointSitePageWebPartHtml {
    param ($WebPart)

    if ($null -eq $WebPart) { return "" }

    $odataType = [string]$WebPart.'@odata.type'
    if ($odataType -like "*textWebPart" -and -not [string]::IsNullOrWhiteSpace([string]$WebPart.innerHtml)) {
        return [string]$WebPart.innerHtml
    }

    $parts = [System.Collections.Generic.List[string]]::new()
    $title = $WebPart.data.title ?? $WebPart.data.description ?? $WebPart.webPartType
    if (-not [string]::IsNullOrWhiteSpace([string]$title)) {
        $parts.Add("<h2>$(ConvertTo-SharePointSitePageHtmlText $title)</h2>")
    }

    $serverContent = $WebPart.data.serverProcessedContent
    foreach ($propertyName in @("htmlStrings", "searchablePlainTexts")) {
        foreach ($entry in @($serverContent.$propertyName)) {
            foreach ($value in @(Get-SharePointSitePageObjectValues -Object $entry -PropertyNames @("value"))) {
                if ($propertyName -eq "htmlStrings") {
                    $parts.Add($value)
                } else {
                    $parts.Add("<p>$(ConvertTo-SharePointSitePageHtmlText $value)</p>")
                }
            }
        }
    }

    foreach ($entry in @($serverContent.links)) {
        $url = $entry.value
        if ([string]::IsNullOrWhiteSpace([string]$url)) { continue }
        $safeUrl = ConvertTo-SharePointSitePageHtmlAttribute $url
        $safeText = ConvertTo-SharePointSitePageHtmlText ($entry.key ?? $url)
        $parts.Add("<p><a href=""$safeUrl"" target=""_blank"">$safeText</a></p>")
    }

    foreach ($entry in @($serverContent.imageSources)) {
        $url = $entry.value
        if ([string]::IsNullOrWhiteSpace([string]$url)) { continue }
        $safeUrl = ConvertTo-SharePointSitePageHtmlAttribute $url
        $safeAlt = ConvertTo-SharePointSitePageHtmlAttribute ($entry.key ?? "SharePoint image")
        $parts.Add("<p><img src=""$safeUrl"" alt=""$safeAlt"" /></p>")
    }

    if ($parts.Count -lt 1) {
        $description = $WebPart.data.description ?? $WebPart.data.title ?? $odataType
        if (-not [string]::IsNullOrWhiteSpace([string]$description)) {
            $parts.Add("<p>$(ConvertTo-SharePointSitePageHtmlText $description)</p>")
        }
    }

    return ($parts -join "`n")
}

function Get-SharePointSitePageWebPartsFromCanvasObject {
    param ($Object)

    if ($null -eq $Object) { return @() }

    $found = [System.Collections.Generic.List[object]]::new()

    if ($Object -is [System.Collections.IEnumerable] -and -not ($Object -is [string])) {
        foreach ($item in @($Object)) {
            foreach ($webPart in @(Get-SharePointSitePageWebPartsFromCanvasObject -Object $item)) {
                $found.Add($webPart)
            }
        }
        return @($found)
    }

    if (-not $Object.PSObject.Properties) { return @() }

    $odataType = [string]$Object.'@odata.type'
    if ($odataType -like "*WebPart" -or $Object.PSObject.Properties["innerHtml"] -or $Object.PSObject.Properties["webPartType"]) {
        $found.Add($Object)
    }

    foreach ($property in @($Object.PSObject.Properties)) {
        if ($property.Name -like "@odata*") { continue }
        if ($null -eq $property.Value) { continue }
        if ($property.Value -is [string] -or $property.Value -is [ValueType]) { continue }

        foreach ($webPart in @(Get-SharePointSitePageWebPartsFromCanvasObject -Object $property.Value)) {
            $found.Add($webPart)
        }
    }

    return @($found)
}

function ConvertTo-SharePointSitePageHtml {
    param (
        $Site,
        $Page,
        [array]$WebParts = @()
    )

    $title = $Page.title ?? $Page.name ?? "Untitled SharePoint page"
    $safeTitle = ConvertTo-SharePointSitePageHtmlText $title
    $safeSite = ConvertTo-SharePointSitePageHtmlText ($Site.displayName ?? $Site.name ?? $Site.id)
    $safeUrl = ConvertTo-SharePointSitePageHtmlAttribute $Page.webUrl
    $safeDescription = ConvertTo-SharePointSitePageHtmlText $Page.description
    $safeModified = ConvertTo-SharePointSitePageHtmlText $Page.lastModifiedDateTime

    $bodyParts = [System.Collections.Generic.List[string]]::new()
    foreach ($webPart in @($WebParts)) {
        $html = ConvertTo-SharePointSitePageWebPartHtml -WebPart $webPart
        if (-not [string]::IsNullOrWhiteSpace([string]$html)) {
            $bodyParts.Add($html)
        }
    }

    if ($bodyParts.Count -lt 1 -and -not [string]::IsNullOrWhiteSpace([string]$Page.description)) {
        $bodyParts.Add("<p>$safeDescription</p>")
    }

    if ($bodyParts.Count -lt 1) {
        $bodyParts.Add("<p>No HTML content was returned for this SharePoint page.</p>")
    }

@"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>$safeTitle</title>
  <style>
    body { font-family: sans-serif; font-size: 14px; line-height: 1.55; color: #242424; padding: 2em; max-width: 960px; margin: 0 auto; }
    img { max-width: 100%; height: auto; }
    .meta { color: #666; font-size: 0.9em; margin-bottom: 2em; }
    .meta a { color: #1a5fb4; }
  </style>
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

function ConvertTo-SharePointSitePagePlainText {
    param ([string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) { return "" }
    $withoutStyle = [regex]::Replace($Html, '<(script|style)\b[^>]*>.*?</\1>', ' ', [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $withoutTags = [regex]::Replace($withoutStyle, '<[^>]+>', ' ')
    $decoded = [System.Web.HttpUtility]::HtmlDecode($withoutTags)
    return (($decoded -replace '\s{2,}', ' ').Trim())
}

$sitePagesJsonDir = $RunSummary.OutputJsonFiles.SitePagesJsonDir ?? (Join-Path $logsFolder "site-pages-json")
$sitePagesHtmlDir = $RunSummary.OutputJsonFiles.SitePagesHtmlDir ?? (Join-Path $logsFolder "site-pages-html")
$sitePagesIndexPath = $RunSummary.OutputJsonFiles.SitePagesIndex ?? (Join-Path $logsFolder "site-pages-index.csv")

foreach ($folder in @($sitePagesJsonDir, $sitePagesHtmlDir)) {
    if (-not (Test-Path -LiteralPath $folder)) {
        $null = New-Item -ItemType Directory -Path $folder -Force
    }
}

$sourceSites = if ($null -ne $SourceDataSites) { @($SourceDataSites) } else { @($userSelectedSites) }
if ($sourceSites.Count -lt 1) {
    Set-PrintAndLog -message "No selected SharePoint sites are available for site page fetching." -Color Yellow
    return
}

$sitePageIndex = [System.Collections.Generic.List[object]]::new()
$sitePageDocuments = [System.Collections.ArrayList]@()

foreach ($site in $sourceSites) {
    $siteId = $site.id
    $siteName = $site.displayName ?? $site.name ?? $site.id
    if ([string]::IsNullOrWhiteSpace([string]$siteId)) {
        Set-PrintAndLog -message "Skipping SharePoint site page fetch for site without id: $siteName" -Color Yellow
        continue
    }

    Set-PrintAndLog -message "Fetching SharePoint site pages for '$siteName'." -Color Cyan

    try {
        $pages = @(Get-SharePointSitePageCollection -SiteId $siteId)
    } catch {
        Set-PrintAndLog -message "Failed to list SharePoint site pages for '$siteName': $($_.Exception.Message)" -Color Red
        $RunSummary.Errors.Add(@{
            Site  = $siteName
            Error = $_.Exception.Message
            Step  = "List SharePoint site pages"
        })
        continue
    }

    $pageNumber = 0
    foreach ($page in $pages) {
        $pageNumber++
        $pageId = $page.id
        $title = $page.title ?? $page.name ?? $page.id

        if ([string]::IsNullOrWhiteSpace([string]$pageId)) {
            Set-PrintAndLog -message "Skipping SharePoint page without id in '$siteName': $title" -Color Yellow
            continue
        }

        try {
            $pageWithCanvas = Get-SharePointSitePageWithCanvas -SiteId $siteId -PageId $pageId
        } catch {
            Set-PrintAndLog -message "Could not fetch canvas layout for '$title'; using list metadata. $($_.Exception.Message)" -Color Yellow
            $pageWithCanvas = $page
        }

        $webParts = @()
        if ($pageWithCanvas.canvasLayout) {
            $webParts = @(Get-SharePointSitePageWebPartsFromCanvasObject -Object $pageWithCanvas.canvasLayout)
        }

        if ($webParts.Count -lt 1) {
            try {
                $webParts = @(Get-SharePointSitePageWebParts -SiteId $siteId -PageId $pageId)
            } catch {
                Set-PrintAndLog -message "Could not fetch webparts for '$title'. $($_.Exception.Message)" -Color Yellow
                $webParts = @()
            }
        }

        $html = ConvertTo-SharePointSitePageHtml -Site $site -Page $pageWithCanvas -WebParts $webParts
        $plainText = ConvertTo-SharePointSitePagePlainText -Html $html
        $fileBaseName = Get-SharePointSitePageSafeFileBaseName -Site $site -Page $pageWithCanvas
        $htmlPath = Join-Path $sitePagesHtmlDir "$fileBaseName.html"
        $jsonPath = Join-Path $sitePagesJsonDir "$fileBaseName.json"

        $html | Out-File -LiteralPath $htmlPath -Encoding UTF8

        $pageExport = [PSCustomObject]@{
            SourceKey            = ('sharepoint:sitePage:{0}:{1}' -f $siteId, $pageId)
            SourceETag           = $pageWithCanvas.eTag
            SiteId               = $siteId
            SiteName             = $siteName
            PageId               = $pageId
            Name                 = $pageWithCanvas.name
            Title                = $title
            WebUrl               = $pageWithCanvas.webUrl
            Description          = $pageWithCanvas.description
            PageLayout           = $pageWithCanvas.pageLayout
            PromotionKind        = $pageWithCanvas.promotionKind
            CreatedDateTime      = $pageWithCanvas.createdDateTime
            LastModifiedDateTime = $pageWithCanvas.lastModifiedDateTime
            HtmlPath             = $htmlPath
            JsonPath             = $jsonPath
            WebPartCount         = @($webParts).Count
            ContentPreview       = if ($plainText.Length -gt 500) { $plainText.Substring(0, 500) } else { $plainText }
            Html                 = $html
            Page                 = $pageWithCanvas
            WebParts             = @($webParts)
        }

        $pageExport | ConvertTo-Json -Depth 80 | Out-File -LiteralPath $jsonPath -Encoding UTF8
        $sitePageDocuments.Add($pageExport) | Out-Null
        $sitePageIndex.Add([PSCustomObject]@{
            SiteName             = $siteName
            SiteId               = $siteId
            Title                = $title
            Name                 = $pageWithCanvas.name
            PageId               = $pageId
            WebUrl               = $pageWithCanvas.webUrl
            HtmlPath             = $htmlPath
            JsonPath             = $jsonPath
            WebPartCount         = @($webParts).Count
            LastModifiedDateTime = $pageWithCanvas.lastModifiedDateTime
        })

        Set-PrintAndLog -message "Fetched SharePoint site page $pageNumber/$($pages.Count) from '$siteName': $title ($(@($webParts).Count) webpart(s))." -Color DarkCyan
    }
}

if ($sitePageIndex.Count -gt 0) {
    $sitePageIndex |
        Sort-Object SiteName, Title |
        Export-Csv -LiteralPath $sitePagesIndexPath -NoTypeInformation -Encoding UTF8
}

if ($null -ne $RunSummary.JobInfo.pages) {
    foreach ($page in @($sitePageDocuments)) {
        [void]$RunSummary.JobInfo.pages.Add($page)
    }
    $RunSummary.JobInfo.pagescount = $RunSummary.JobInfo.pages.Count
}

Set-PrintAndLog -message "SharePoint site page fetch complete: $($sitePageDocuments.Count) page(s). HTML: $sitePagesHtmlDir; JSON: $sitePagesJsonDir; index: $sitePagesIndexPath" -Color Green
