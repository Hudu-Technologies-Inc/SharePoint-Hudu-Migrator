
function Relink-DocumentUploads {
    param (
        [AllowEmptyCollection()] [array]$Docs = @()
        
    )

    $Docs = @($Docs | Where-Object { $null -ne $_ })
    if ($Docs.Count -lt 1) {
        Set-PrintAndLog -message "No stubbed articles queued for relinking." -Color DarkGray
        return
    }

    foreach ($doc in $Docs) {
        if (-not $doc -or -not $doc.stub -or -not $doc.stub.id) {
            Set-PrintAndLog -message "Skipping relink for item without a Hudu stub: $($doc.title ?? $doc.LocalPath ?? $doc.Name ?? 'unknown')" -Color DarkGray
            continue
        }

        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($doc.FilePath)
        $htmlPath = $doc.NewPath

        # Paths to supporting JSON files
        $linksPath      = "$htmlPath-links.json"
        $uploadedPath   = "$htmlPath-uploaded.json"
        $attachmentsPath = "$htmlPath-attachments.json"
        # Load data
        $uploadedInfo = @($doc.UploadedFiles)
        $foundLinks   = Get-LinksFromHTML -htmlContent $doc.ReplacedContent -title ($doc.title ?? $doc.localpath) -includeImages $true -suppressOutput $true
        $attachments  = $doc.AllAttachments
        $webViewUrl = $doc.webViewUrl
        if (-not $webViewUrl) {
            $webViewUrl = @($doc.OriginalLinks)[0]
        }

        $originalFilename = $doc.OriginalFilename ?? $doc.LocalPath
        $filenameOnly = if ($originalFilename) {
            [System.IO.Path]::GetFileName($originalFilename).ToLowerInvariant()
        } else {
            ""
        }

        $docAsAttachment = @(
            $uploadedInfo |
                Where-Object {
                    $uploadName = $_.OriginalFilename ?? $_.name ?? $_.filename
                    $uploadFileName = if ($uploadName) { [System.IO.Path]::GetFileName([string]$uploadName).ToLowerInvariant() } else { "" }
                    $filenameOnly -and $uploadFileName -eq $filenameOnly
                } |
                Select-Object -First 1
        )[0]
        $docAsAttachmentUrl = $docAsAttachment.MappedUrl ?? $docAsAttachment.url
        if (-not $docAsAttachmentUrl -and $docAsAttachment.id) {
            $docAsAttachmentUrl = if ($docAsAttachment.UploadType -eq 'image') {
                "$HuduBaseURL/public_photo/$($docAsAttachment.id)"
            } else {
                "$HuduBaseURL/file/$($docAsAttachment.id)"
            }
        }
        $AttachmentMap = @{}
        foreach ($upload in $doc.UploadedFiles) {
            if (-not $upload.PSObject.Properties['ext']) {
                $upload | Add-Member -NotePropertyName 'ext' -NotePropertyValue `
                    ([System.IO.Path]::GetExtension($upload.OriginalFilename).TrimStart('.')) -Force
            }
            $filename = [System.IO.Path]::GetFileName($upload.OriginalFilename).ToLowerInvariant()
            $AttachmentMap[$filename] = $upload
        }
        # Read HTML
        $html = $doc.replacedContent
        if (-not $doc.PSObject.Properties['OverrideContent']) {
        # Replace all links or filenames matching the original filename, then attachments
            $html = Replace-HuduAttachmentLinkBlock -html $html -sourceFile $doc
            foreach ($link in $foundLinks) {
                if ($filenameOnly -and $docAsAttachmentUrl -and $link.ToLowerInvariant() -like "*$filenameOnly*") {
                    Set-PrintandLog -Message "linking $($link.ToLowerInvariant()) => $docAsAttachmentUrl via $filenameOnly"
                    $html = $html -replace [regex]::Escape($link), $docAsAttachmentUrl
                }
                foreach ($attachedFile in $doc.UploadedFiles){
                    $attachedSourceName = $attachedFile.OriginalFilename ?? $attachedFile.name ?? $attachedFile.filename
                    if ([string]::IsNullOrWhiteSpace([string]$attachedSourceName)) { continue }
                    $attachedfilenameOnly = [System.IO.Path]::GetFileName($attachedSourceName).ToLowerInvariant()
                    $attachedUrl = $attachedFile.MappedUrl ?? $attachedFile.url
                    if (-not $attachedUrl -and $attachedFile.id) {
                        $attachedUrl = if ($attachedFile.UploadType -eq 'image') {
                            "$HuduBaseURL/public_photo/$($attachedFile.id)"
                        } else {
                            "$HuduBaseURL/file/$($attachedFile.id)"
                        }
                    }
                    if ([string]::IsNullOrWhiteSpace([string]$attachedUrl)) { continue }
                    if ($link.ToLowerInvariant() -like "*$attachedfilenameOnly*") {
                        Set-PrintandLog -Message "linking attachment $($link.ToLowerInvariant()) => $attachedUrl via $attachedfilenameOnly"
                        $html = $html -replace [regex]::Escape($link), $attachedUrl
                    }
                }
            }
            $updatedHTML = if ($originalFilename -and $docAsAttachmentUrl) {
                $html -replace [regex]::Escape($originalFilename), $docAsAttachmentUrl
            } else {
                $html
            }
            $updatedHTML = Replace-SharePointAttachmentTags -Html $updatedHTML -AttachmentMap $AttachmentMap -HuduBaseUrl $HuduBaseURL
            $updatedHTML = Replace-SharePointLinkBlock -html $updatedHTML -webViewUrl $webViewUrl        
        } else {
            $updatedHTML = $doc.OverrideContent
        }


        $doc.replacedContent =$updatedHTML
        $doc | Add-Member -NotePropertyName FinalContent -NotePropertyValue $updatedHTML -Force

        # Save back
        $doc | Add-Member -NotePropertyName ReplacedLinks -NotePropertyValue (Get-LinksFromHTML -htmlContent $updatedHTML -title ($doc.title ?? $doc.localpath) -includeImages $true -suppressOutput $false) -Force
        Save-HtmlSnapshot -PageId $doc.id -Title $doc.title -Content $updatedHTML -Suffix "relinked" -OutDir $tmpfolder
        Export-DocPropertyJson -Doc $doc -Property 'ReplacedLinks'
        Set-PrintAndLog "Relinked HTML locally: $htmlPath" -Color Green
    }
}
Relink-DocumentUploads -docs @($stubbedArticles)
