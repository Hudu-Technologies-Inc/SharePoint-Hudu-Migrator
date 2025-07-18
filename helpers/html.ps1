$SHAREPOINT_URL_DELIMITER = "<SHAREPOINT_WEBVIEW_DELIMITER>"
$HUDU_LOCALATTACHMENT_DELIMITER = "<HUDU_LOCALATTACHMENT_DELIMITER>"

function Make-HiddenHTMLParaGraph {
  param (
    [string]$paragraphContent
  )
  return @"
<p class="hidden-paragraph">$paragraphContent</p>

<style>
.hidden-paragraph {
  display: none;
}
</style>
"@
}
function Get-AttachmentLink {
    param (
                [PSCustomObject]$sourceFile
    )

return $(if ($RunSummary.SetupInfo.SourceFilesAsAttachments) {
@"
<br><a href='$([System.IO.Path]::GetFileName($sourceFile.LocalPath))'>Attached Original File: $($sourceFile.title)</a>
"@
} else {
    ""
})
}
function Get-GeneratedHTMLForImageFile {
    param (
        [Parameter(Mandatory)]
        [PSCustomObject]$sourceFile,

        [Parameter(Mandatory)]
        [string]$outputFile
    )

    $filename = [System.IO.Path]::GetFileName($sourceFile.LocalPath)
    $title    = [System.Web.HttpUtility]::HtmlEncode($sourceFile.title)
    $site     = [System.Web.HttpUtility]::HtmlEncode($sourceFile.SiteName)

    $html = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Embedded Image - $title</title>
  <style>
    body { font-family: sans-serif; text-align: center; background: #f9f9f9; padding: 2em; }
    img { max-width: 95%; height: auto; border: 1px solid #ccc; background: #fff; padding: 10px; }
    .meta { font-size: 0.9em; margin-top: 1em; }
  </style>
</head>
<body>
  <h1>$title</h1>
  <img src="$filename" alt="$title" />
  <div class="meta">
    <p><strong>From:</strong> $site</p>
    $SHAREPOINT_URL_DELIMITER
    <hr>
  </div>
  $HUDU_LOCALATTACHMENT_DELIMITER
</body>
</html>
"@

    Set-Content -Path $outputFile -Value $html -Encoding UTF8
    return $outputFile1
}
function Get-GeneratedAttachmentLinkLargeDocs {
    param (
        [Parameter(Mandatory)]
        [PSCustomObject]$sourceFile,

        [Parameter(Mandatory)]
        [string]$outputFile, 

        [string]$link,
        [string]$note
    )

    $filename     = [System.IO.Path]::GetFileName($sourceFile.LocalPath)
    $title        = [System.Web.HttpUtility]::HtmlEncode($sourceFile.title)
    $site         = [System.Web.HttpUtility]::HtmlEncode($sourceFile.SiteName)
    $fileSizeKB = if (Test-Path $sourceFile.LocalPath) {
        [math]::Round($sourceFile.Filesize / 1KB, 2)
    } else {
        "N/A"
    }

    $html = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>$title</title>
  <style>
    body { font-family: sans-serif; padding: 2em; background: #f4f4f4; }
    h1 { color: #333; }
    .info { font-size: 0.95em; line-height: 1.6em; }
  </style>
</head>
<body>
  <h1>$title</h1>
  <p class="info">
    This document was too large to convert into a readable HTML format.<br />
    <strong>Site:</strong> $site<br />
    <strong>File:</strong> $filename ($fileSizeKB KB)<br />
    $link
    
  </p>
</body>
</html>
"@

    Set-Content -Path $outputFile -Value $html -Encoding UTF8
    return $outputFile
}


function Get-DisallowedExtensionGeneratedHTML {
    param (
        [Parameter(Mandatory)]
        [PSCustomObject]$sourceFile,

        [Parameter(Mandatory)]
        [string]$outputFile
    )

    $baseName     = [System.IO.Path]::GetFileNameWithoutExtension($sourceFile.LocalPath)
    $extension    = [System.IO.Path]::GetExtension($sourceFile.LocalPath).ToLowerInvariant()
    $title        = [System.Web.HttpUtility]::HtmlEncode($sourceFile.title)
    $site         = [System.Web.HttpUtility]::HtmlEncode($sourceFile.SiteName)
    $drivePath    = [System.Web.HttpUtility]::HtmlEncode($sourceFile.parentDrivePath)
    $fileSizeKB = if (Test-Path $sourceFile.LocalPath) {
    [math]::Round($sourceFile.Filesize / 1KB, 2)
    } else {
    "N/A"
    }



    $html = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Disallowed FileType - $baseName</title>
  <style>
    body { font-family: sans-serif; font-size: 13px; line-height: 1.5; background-color: #f9f9f9; color: #333; padding: 2em; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 2em; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: left; vertical-align: top; }
    th { background-color: #f0f0f0; }
    tr:nth-child(even) { background-color: #fafafa; }
    a { color: #1a0dab; text-decoration: none; }
    a:hover { text-decoration: underline; }
  </style>
</head>
<body>
  <h1>Disallowed File Reference</h1>
  <p>This file was skipped during conversion due to its disallowed filetype: <strong>$extension</strong>.</p>
  <p>$HUDU_LOCALATTACHMENT_DELIMITER</p>
  <table>
    <thead>
      <tr>
        <th>Title</th>
        <th>Site</th>
        <th>SharePoint Link</th>
        <th>Drive Path</th>
        <th>Download URL</th>
        <th>Size (KB)</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td>$title</td>
        <td>$site</td>
        $SHAREPOINT_URL_DELIMITER
        <td>$drivePath</td>
        <td>$fileSizeKB</td>
      </tr>
    </tbody>
  </table>
</body>
</html>
"@

    Set-Content -Path $outputFile -Value $html -Encoding UTF8
    return $outputFile
}


function Compress-Html {
    param (
        [Parameter(Mandatory = $true)][string]$Html
    )
    # Remove b64 data without a useful mime type
    $mimePattern = "data:(?!" + $AvailableB64MimeTypes + ")[^;]*;base64,[^`"'\s>]+"
    $Html = $Html -replace $mimePattern, ''
    $Html = $Html -replace '<img[^>]*src="data:[^:;"]*;base64,[^"]*"[^>]*>', ''
    $Html = $Html -replace '<[^>]*(src|href)="data:[^:;"]*;base64,[^"]*"[^>]*>', ''


    # Remove <meta>, <link>, <script> blocks
    $Html = $Html -replace '(?i)<(meta|link|script)[^>]*>.*?</\1>', ''
    $Html = $Html -replace '(?i)<table[^>]*>\s*<tr[^>]*>\s*<td[^>]*>\s*', ''
    $Html = $Html -replace '(?i)\s*</td>\s*</tr>\s*</table>', ''

    # Strip inline styles completely
    $Html = $Html -replace '\s*style="[^"]*"', ''

    # Strip unnecessary attributes
    $Html = $Html -replace '\s+xmlns(:\w+)?="[^"]*"', ''
    $Html = $Html -replace '\s+lang="[^"]*"', ''
    $Html = $Html -replace '\s+dir="[^"]*"', ''
    $Html = $Html -replace '\s+xml:lang="[^"]*"', ''

    # Collapse whitespace
    $Html = $Html -replace '\s{2,}', ' '
    $Html = $Html -replace '\r?\n', ''

    # Strip font-related inline styles (if they slipped through)
    $Html = $Html -replace 'font-family:[^;"]*;', ''
    $Html = $Html -replace 'font-size:[^;"]*;', ''
    # Remove Excel auto-generated classes
    $Html = $Html -replace '\sclass="xl\d+"', ''

    # Remove empty rows/cells
    $Html = $Html -replace '<tr[^>]*>\s*</tr>', ''
    $Html = $Html -replace '<td[^>]*>\s*</td>', ''

    # Remove empty tags
    $Html = $Html -replace '<(p|div)>\s*(<br\s*/?>)?\s*</\1>', ''
    $Html = $Html -replace '<span>\s*</span>', ''
    $Html = $Html -replace '<strong>\s*</strong>', ''
    $Html = $Html -replace '<em>\s*</em>', ''

    # Optional: Replace double <br><br> with paragraph
    $Html = $Html -replace '(<br\s*/?>\s*){2,}', '</p><p>'

    # Remove Word/LibreOffice class clutter
    $Html = $Html -replace '\sclass="Mso[^"]*"', ''
    $Html = $Html -replace '\sclass="TableGrid[^"]*"', ''
    $Html = $Html -replace '\sclass="[^"]*?"', ''

    # Remove MS VML and Office-specific tags
    $Html = $Html -replace '<v:[^>]+>.*?</v:[^>]+>', ''
    $Html = $Html -replace '<o:[^>]+>.*?</o:[^>]+>', ''

    # Remove excessive div wrappers with position:abs
    $Html = $Html -replace '<div[^>]*style="[^"]*position:absolute;[^"]*"[^>]*>', ''

    # Remove Word-generated comments and namespaces
    $Html = $Html -replace '<!--\[if.*?endif\]-->', ''
    $Html = $Html -replace 'xmlns(:\w+)?="[^"]*"', ''

    # Collapse empty paragraphs and spans
    $Html = $Html -replace '<p[^>]*>\s*</p>', ''
    $Html = $Html -replace '<span[^>]*>\s*</span>', ''

    # Remove ridiculous <style> blocks
    $Html = $Html -replace '<style[^>]*>.*?</style>', ''

    # Remove empty <div>, <span>
    $Html = $Html -replace '<(div|span)[^>]*>\s*</\1>', ''

    # Remove <!-- comments -->
    $Html = $Html -replace '<!--.*?-->', ''
    $Html = $Html -replace '\s{2,}', ' '     # Collapse spaces
    $Html = $Html -replace '\n|\r', ''       # Remove newlines
    $Html = $Html.Trim()

    return $Html

}
function Get-LinksFromHTML {
    param (
        [string]$htmlContent,
        [string]$title,
        [bool]$includeImages = $true,
        [bool]$suppressOutput = $false

    )

    $allLinks = @()

    # Match all href attributes inside anchor tags
    $hrefPattern = '<a\s[^>]*?href=["'']([^"'']+)["'']'
    $hrefMatches = [regex]::Matches($htmlContent, $hrefPattern, 'IgnoreCase')
    foreach ($match in $hrefMatches) { 
        $allLinks += $match.Groups[1].Value
    }

    if ($includeImages) {
        # Match all src attributes inside img tags
        $srcPattern = '<img\s[^>]*?src=["'']([^"'']+)["'']'
        $srcMatches = [regex]::Matches($htmlContent, $srcPattern, 'IgnoreCase')
        foreach ($match in $srcMatches) {
            $allLinks += $match.Groups[1].Value
        }
    }
    if ($false -eq $suppressOutput){
        $linkidx=0
        foreach ($link in $allLinks) {
            $linkidx=$linkidx+1
            Set-PrintAndLog -message "link $linkidx of $($allLinks.count) total found for $title - $link" -Color Blue
        }
    }

    return $allLinks | Sort-Object -Unique
}
function Replace-SharePointAttachmentTags {
    param(
        [string]$Html,
        [hashtable]$AttachmentMap,
        [string]$HuduBaseUrl
    )

    foreach ($filename in $AttachmentMap.Keys) {
        $upload = $AttachmentMap[$filename]
        if (-not $upload) { continue }

        $ext = if ($upload.ext) { $upload.ext.ToLowerInvariant() } else {
            [System.IO.Path]::GetExtension($upload.OriginalFilename).Trim('.').ToLowerInvariant()
        }
        $id = $upload.id
        $safeFilename = [regex]::Escape($filename)

        $url = "$HuduBaseUrl/file/$id"
        $imgUrl = "$HuduBaseUrl/public_photo/$id"

        if ($ext -match '^(jpg|jpeg|png|webp)$') {
            $replacement = "<a href='$imgUrl' target='_blank'><img src='$imgUrl' alt='$filename' /></a>"
        } elseif ($ext -match '^(gif|bmp|svg)$') {
            $replacement = "<a href='$url' target='_blank'><img src='$url' alt='$filename' /></a>"
        } else {
            $replacement = "<a href='$url'>$filename</a>"
        }

        # Replace anywhere in the HTML that matches this filename
        $Html = [regex]::Replace($Html, $safeFilename, [regex]::Escape($replacement))
    }

    return $Html
}

function Replace-SharePointLinkBlock {
    param (
        [string]$html,
        [string]$webViewUrl
    )

    $replacement = if ($webViewUrl) {
        "<a href='$webViewUrl' target='_blank'>View in SharePoint</a>"
    } else {
        ""
    }
    if ($RunSummary.SetupInfo.LinkSourceArticles -eq $false){
      $replacement="$(Make-HiddenHTMLParaGraph -paragraphContent $replacement)"
    }

    return $html -replace [regex]::Escape($SHAREPOINT_URL_DELIMITER), $replacement
}

function Replace-HuduAttachmentLinkBlock {
    param (
        [string]$html,
        [PSCustomObject]$sourceFile
    )

    $replacement = if ($sourceFile) {
        Get-AttachmentLink -sourceFile $sourceFile
    } else {
        ""
    }
    if ($RunSummary.SetupInfo.SourceFilesAsAttachments -eq $false) {
      $replacement="$(Make-HiddenHTMLParaGraph -paragraphContent $replacement)"
    }


    return $html -replace [regex]::Escape($HUDU_LOCALATTACHMENT_DELIMITER), $replacement
}