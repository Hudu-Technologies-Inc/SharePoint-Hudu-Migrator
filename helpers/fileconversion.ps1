
function Convert-WithLibreOffice {
    param (
        [string]$inputFile,
        [string]$outputDir,
        [string]$sofficePath
    )

    try {
        $extension = [System.IO.Path]::GetExtension($inputFile).ToLowerInvariant()
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)

        switch ($extension.ToLowerInvariant()) {
            # Word processors
            ".doc"      { $intermediateExt = "odt" }
            ".docx"     { $intermediateExt = "odt" }
            ".docm"     { $intermediateExt = "odt" }
            ".rtf"      { $intermediateExt = "odt" }
            ".txt"      { $intermediateExt = "odt" }
            ".md"       { $intermediateExt = "odt" }
            ".wpd"      { $intermediateExt = "odt" }

            # Spreadsheets
            ".xls"      { $intermediateExt = "ods" }
            ".xlsx"     { $intermediateExt = "ods" }
            ".csv"      { $intermediateExt = "ods" }

            # Presentations
            ".ppt"      { $intermediateExt = "odp" }
            ".pptx"     { $intermediateExt = "odp" }
            ".pptm"     { $intermediateExt = "odp" }

            # Already OpenDocument
            ".odt"      { $intermediateExt = $null }
            ".ods"      { $intermediateExt = $null }
            ".odp"      { $intermediateExt = $null }

            default { $intermediateExt = $null }
        }
        if ($intermediateExt) {
            $intermediatePath = Join-Path $outputDir "$baseName.$intermediateExt"
            Set-PrintAndLog -message "Step 1: Converting to .$intermediateExt..." -Color DarkCyan

            Start-Process -FilePath "$sofficePath" `
                -ArgumentList "--headless", "--convert-to", $intermediateExt, "--outdir", "`"$outputDir`"", "`"$inputFile`"" `
                -Wait -NoNewWindow

            if (-not (Test-Path $intermediatePath)) {
                throw "$intermediateExt conversion failed for $inputFile"
            }
        } else {
            # No conversion needed
            $intermediatePath = $inputFile
        }

        Set-PrintAndLog -message  "Step $(if ($intermediateExt) {'2'} else {'1'}): Converting .$intermediateExt to XHTML..." -Color DarkCyan

        Start-Process -FilePath "$sofficePath" `
            -ArgumentList "--headless", "--convert-to", "xhtml", "--outdir", "`"$outputDir`"", "`"$intermediatePath`"" `
            -Wait -NoNewWindow

        $htmlPath = Join-Path $outputDir "$baseName.xhtml"

        if (-not (Test-Path $htmlPath)) {
            throw "XHTML conversion failed for $intermediatePath"
        }

        return $htmlPath
    }
    catch {
        Write-ErrorObjectsToFile -ErrorObject @{
            fileconversionError = @{
                error      = $_
                file       = $inputFile
                officepath = $sofficePath
                outdir     = $outputDir
            }
        }
        return $null
    }
}

function Get-EmbeddedFilesFromHtml {
    param (
        [string]$htmlPath,
        [int32]$resolution=5
    )

    if (-not (Test-Path $htmlPath)) {
        Write-Warning "HTML file not found: $htmlPath"
        return @{}
    }

    $htmlContent = Get-Content $htmlPath -Raw
    $baseDir = Split-Path -Path $htmlPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($htmlPath)
    $trimmedBaseName = if ($baseName.Length -gt $resolution) {
        $baseName.Substring(0, $baseName.Length - $resolution).ToLower()
    } else {
        $baseName.ToLower()
    }
    $results = @{
        ExternalFiles        = @()
        Base64Images         = @()
        Base64ImagesWritten  = @()
        UpdatedHTMLContent   = $null
    }

    $guid = [guid]::NewGuid().ToString()
    $uuidSuffix = ($guid -split '-')[0]

    $counter = 0
    $htmlContent = [regex]::Replace($htmlContent, '(?i)<img([^>]+?)src\s*=\s*["'']data:image/(?<type>[a-z]+);base64,(?<b64data>[^"'']+)["'']', {
        param($match)

        $type = $match.Groups["type"].Value
        $b64  = $match.Groups["b64data"].Value

        $ext = switch ($type) {
            'png'  { 'png' }
            'jpeg' { 'jpg' }
            'jpg'  { 'jpg' }
            'gif'  { 'gif' }
            'svg'  { 'svg' }
            'bmp'  { 'bmp' }
            default { 'bin' }
        }

        $counter++
        $filename = "${baseName}_embedded_${uuidSuffix}_$counter.$ext"
        $filepath = Join-Path $baseDir $filename

        try {
            [IO.File]::WriteAllBytes($filepath, [Convert]::FromBase64String($b64))
            $results.ExternalFiles += $(join-path $baseDir $filename)
            $results.Base64Images  += "data:image/$type;base64,..."
            $results.Base64ImagesWritten += $(join-path $baseDir $filename)

            return "<img$($match.Groups[1].Value)src='$filename'"
        } catch {
            Write-Warning "Failed to decode embedded image: $($_.Exception.Message)"
            return "<img$($match.Groups[1].Value)src='$filename'"
        }
    })
    $skipExts = @(
        ".doc", ".docx", ".docm", ".rtf", ".txt", ".md", ".wpd",
        ".xls", ".xlsx", ".csv", ".ppt", ".pptx", ".pptm",
        ".odt", ".ods", ".odp", ".xhtml", ".xml", ".html", ".json", ".htm"
    )

    $allFiles = Get-ChildItem -Path $baseDir -File
    foreach ($file in $allFiles) {
        $fullFilePath = [IO.Path]::GetFullPath($file.FullName).ToLowerInvariant()
        $htmlPathNormalized = [IO.Path]::GetFullPath($htmlPath).ToLowerInvariant()

        if ($fullFilePath -eq $htmlPathNormalized) {
            continue
        }

        if ($file.Extension.ToLowerInvariant() -in $skipExts) {
            continue
        }

        $otherBaseName = $file.BaseName.ToLower()
        if ($otherBaseName.StartsWith($trimmedBaseName)) {
            $results.ExternalFiles += "$fullFilePath"
        }
    }
        
        
    $results.UpdatedHTMLContent = $htmlContent
    return $results
}

# TODO: DRY this up later.
function ConvertDownloadedFiles {
    param (
        $downloadedFiles,
        $sofficePath
    )

    $convertedbatch = [System.Collections.ArrayList]@()

    $convertIDX=0
    foreach ($file in $downloadedFiles) {
        $convertIDX=$convertIDX+1
        $completionPercentage = Get-PercentDone -Current $convertIDX -Total $downloadedFiles.count

        if (-not $file.LocalPath) {
            Set-PrintAndLog -message "Skipping site-level reference object..." -Color Yellow
            continue
        }

        $file | Add-Member -NotePropertyName SuccessConverted -NotePropertyValue $false -Force
        $file | Add-Member -NotePropertyName NewPath -NotePropertyValue $null -Force
        $file | Add-Member -NotePropertyName ConversionError -NotePropertyValue $null -Force

        Set-PrintAndLog -message "processing $($file.LocalPath)" -Color Green

        try {
            $extension = [System.IO.Path]::GetExtension($file.LocalPath).ToLowerInvariant()
            Set-PrintAndLog -Message "Checking extension '$extension'" -Color Green

            $outputDir = Split-Path $file.LocalPath
            $htmlPath = $null
            # images as sharepoint file download
            if ($EmbeddableImageExtensions -contains $extension){
                Set-PrintAndLog -message "Image extension: $extension — generating user-friendly HTML." -Color Yellow
                $file.NewPath = Join-Path $outputDir "$([System.IO.Path]::GetFileNameWithoutExtension($file.localpath))-gen-image.html"
                Get-GeneratedHTMLForImageFile -sourceFile $file -outputFile $file.newpath
                $file.RawContent = Get-Content $file.NewPath -Raw
                $file.ReplacedContent = $file.RawContent
                $file.SuccessConverted = $false
                $file.UsingGeneratedHTML = $true
                $file | Add-Member -NotePropertyName ExternalEmbeddedFiles -NotePropertyValue ([System.Collections.ArrayList]@()) -Force
                $file | Add-Member -NotePropertyName Base64EmbeddedImages  -NotePropertyValue ([System.Collections.ArrayList]@()) -Force
                $file | Add-Member -NotePropertyName AllAttachments -NotePropertyValue $(if ($RunSummary.SetupInfo.SourceFilesAsAttachments) {@($file.LocalPath)} else {[System.Collections.ArrayList]@()}) -Force
                $convertedbatch.Add($file) | Out-Null
                continue
            }
            # disallowed for conversion as sharepoint file download
            if ($RunSummary.SetupInfo.DisallowedForConvert -contains $extension){
                Set-PrintAndLog -message "extension: $extension is disallowed for converting— skipping conversion." -Color Yellow
                $file.NewPath = Join-Path $outputDir "$([System.IO.Path]::GetFileNameWithoutExtension($file.localpath))-generated.html"
                Get-DisallowedExtensionGeneratedHTML -sourceFile $file -outputFile $file.NewPath
                $file.RawContent = Get-Content $file.NewPath -Raw
                $file.ReplacedContent = $file.RawContent
                $file.SuccessConverted = $false
                $file.UsingGeneratedHTML = $true
                $file | Add-Member -NotePropertyName ExternalEmbeddedFiles -NotePropertyValue ([System.Collections.ArrayList]@()) -Force
                $file | Add-Member -NotePropertyName Base64EmbeddedImages  -NotePropertyValue ([System.Collections.ArrayList]@()) -Force
                $file | Add-Member -NotePropertyName AllAttachments -NotePropertyValue $(if ($RunSummary.SetupInfo.SourceFilesAsAttachments) {@($file.LocalPath)} else {[System.Collections.ArrayList]@()}) -Force
                $convertedbatch.Add($file) | Out-Null
                continue    
            }
            switch ($extension) {
                ".pdf" {
                    $htmlPath = Convert-PdfToSlimHtml -InputPdfPath $file.localpath -PdfToHtmlPath $PDFToHTML
                    # $htmlPath = Convert-PdfToHtml -inputPath $file.LocalPath `
                    #                               -pdftohtmlPath $PDFToHTML `
                    #                               -includeHiddenText $includeHiddenText `
                    #                               -complexLayoutMode $includeComplexLayouts
                }
                default {
                    $htmlPath = Convert-WithLibreOffice -inputFile $file.LocalPath `
                                                  -outputDir $outputDir `
                                                  -sofficePath $sofficePath
                }
            }

            if ($htmlPath -and (Test-Path $htmlPath)) {
                $file.NewPath = $htmlPath
                $file.RawContent = Get-Content $file.NewPath -Raw

                $file.SuccessConverted = $true
                Set-PrintAndLog -message "Converted: $($file.LocalPath) => $htmlPath" -Color Green
                Set-PrintAndLog -message "Discovering Embedded Files for $htmlPath" -Color DarkGreen

                $foundfiles = Get-EmbeddedFilesFromHtml -htmlPath $htmlPath
                $totalFound = [int]$foundfiles.Base64Images.Count + [int]$foundfiles.ExternalFiles.Count
                Set-PrintAndLog -message "Found $totalFound ($($foundfiles.ExternalFiles.count) extracted / $($foundfiles.Base64Images.count) embedded) inside converted $htmlpath" -Color DarkMagenta
                if ($foundfiles.UpdatedHTMLContent) {
                    $file.ReplacedContent = "$($foundfiles.UpdatedHTMLContent)<br>$($SHAREPOINT_URL_DELIMITER)<br>$($HUDU_LOCALATTACHMENT_DELIMITER)"
                }
                $file | Add-Member -NotePropertyName ExternalFiles -NotePropertyValue $foundfiles.ExternalFiles -Force
                $file | Add-Member -NotePropertyName Base64ImagesWritten  -NotePropertyValue $foundfiles.Base64ImagesWritten  -Force
                $allfiles=@() 
                if ($RunSummary.SetupInfo.SourceFilesAsAttachments) {
                    $allFiles = @(@($file.ExternalFiles) + @($file.Base64ImagesWritten) + @($file.LocalPath)) | Sort-Object -Unique
                } else {
                    $allFiles = @(@($file.ExternalFiles) + @($file.Base64ImagesWritten)) | Sort-Object -Unique
                }
                $file | Add-Member -NotePropertyName AllAttachments -NotePropertyValue $allFiles -Force

            }
        } catch {
            $file.ConversionError = $_.Exception.Message
            $file.SuccessConverted = $false
            Write-ErrorObjectsToFile -ErrorObject @{
                FileObject = $file
                FoundEmbeds = $foundFiles
                HtmlPath   = $htmlPath
            } -Name "converterror-$(Get-SafeTitle $($file.LocalPath))"
            continue
        }
        $convertedbatch.Add($file) | Out-Null
        Write-Progress -Activity "converting $($file.title)" -Status "$completionPercentage%" -PercentComplete $completionPercentage

    }

    return $convertedbatch
}
function Convert-PdfToSlimHtml {
    param (
        [Parameter(Mandatory)][string]$InputPdfPath,
        [string]$OutputDir = (Split-Path -Path $InputPdfPath),
        [string]$PdfToHtmlPath = "C:\tools\poppler\bin\pdftohtml.exe"
    )

    if (-not (Test-Path $InputPdfPath)) {
        throw "PDF not found: $InputPdfPath"
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPdfPath)
    $xmlOutput = Join-Path $OutputDir "$baseName.xml"
    $htmlOutput = Join-Path $OutputDir "$baseName.slim.html"

    $args = @(
        "-xml"            # XML format
        "-p"              # Extract images
        "-zoom", "1.0"    # No zoom distortion
        "-noframes"       # Single output file
        "-nomerge"        # Keep layout simple
        "-enc", "UTF-8"
        "-nodrm"
        "`"$InputPdfPath`"",
        "`"$xmlOutput`""
    )

    # Run conversion to XML
    Start-Process -FilePath $PdfToHtmlPath -ArgumentList $args -NoNewWindow -Wait

    if (-not (Test-Path $xmlOutput)) {
        throw "XML output was not created."
    }

    # Convert XML to lightweight HTML
    Convert-PdfXmlToHtml -XmlPath $xmlOutput -OutputHtmlPath $htmlOutput
    return $htmlOutput
}

function Convert-PdfXmlToHtml {
    param (
        [Parameter(Mandatory)][string]$XmlPath,
        [string]$OutputHtmlPath = "$XmlPath.html"
    )

    if (-not (Test-Path $XmlPath)) {
        throw "Input XML not found: $XmlPath"
    }

    [xml]$doc = Get-Content $XmlPath
    $html = @()
    $html += '<!DOCTYPE html>'
    $html += '<html><head><meta charset="UTF-8">'
    $html += '<style>body{font-family:sans-serif;font-size:12pt;line-height:1.4}</style></head><body>'

    foreach ($page in $doc.pdf2xml.page) {
        $html += "<div class='page' style='margin-bottom:2em'>"
        foreach ($text in $page.text) {
            $content = ($text.'#text' -replace '\s+', ' ').Trim()
            if ($content) {
                $html += "<p>$content</p>"
            }
        }
        $html += "</div>"
    }

    $html += '</body></html>'
    Set-Content -Path $OutputHtmlPath -Value ($html -join "`n") -Encoding UTF8
    Set-PrintAndLog -message  "Generated slim HTML: $OutputHtmlPath" -Color Green
}
function Convert-PdfToHtml {
    param (
        [string]$inputPath,
        [string]$outputDir = (Split-Path $inputPath),
        [string]$pdftohtmlPath = "C:\tools\poppler\bin\pdftohtml.exe",
        [bool]$includeHiddenText = $true,
        [bool]$complexLayoutMode = $true
    )

    $filename = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
    $outputHtml = Join-Path $outputDir "$filename.html"

    $popplerArgs = @()

    # Preserve layout with less nesting
    if ($complexLayoutMode) {
        $popplerArgs += "-c"            # complex layout mode
    }

    # Enable image extraction
    $popplerArgs += "-p"                # extract images
    $popplerArgs += "-zoom 1.0"         # avoid automatic zoom bloat

    # Output options
    $popplerArgs += "-noframes"        # single HTML file instead of one per page
    $popplerArgs += "-nomerge"         # don't merge text blocks (more control)
    $popplerArgs += "-enc UTF-8"       # UTF-8 encoding
    $popplerArgs += "-nodrm"           # ignore any DRM restrictions

    if ($includeHiddenText) {
        $popplerArgs += "-hidden"
    }

    # Wrap file paths
    $popplerArgs += "`"$inputPath`""
    $popplerArgs += "`"$outputHtml`""

    Start-Process -FilePath $pdftohtmlPath `
        -ArgumentList $popplerArgs -Wait -NoNewWindow

    return (Test-Path $outputHtml) ? $outputHtml : $null
}


function Save-Base64ToFile {
    param (
        [Parameter(Mandatory)]
        [string]$Base64String,

        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    # Remove data URI prefix if present (e.g., "data:image/png;base64,...")
    if ($Base64String -match '^data:.*?;base64,') {
        $Base64String = $Base64String -replace '^data:.*?;base64,', ''
    }

    $bytes = [System.Convert]::FromBase64String($Base64String)
    [System.IO.File]::WriteAllBytes($OutputPath, $bytes)

    Set-PrintAndLog -message  "Saved Base64 content to: $OutputPath" -Color Cyan
}

