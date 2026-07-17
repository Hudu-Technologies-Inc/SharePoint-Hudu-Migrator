
function Relink-DocumentUploads {
    param (
        [Parameter(Mandatory)] [array]$Docs
        
    )

    foreach ($doc in $Docs) {
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

        $originalFilename = @($uploadedInfo | Where-Object { $_.OriginalFilename } | Select-Object -First 1).OriginalFilename
        if (-not $originalFilename) {
            $originalFilename = $doc.OriginalFilename ?? $doc.LocalPath
        }
        $filenameOnly = if ($originalFilename) {
            [System.IO.Path]::GetFileName($originalFilename).ToLowerInvariant()
        } else {
            ""
        }

        $docAsAttachmentUrl = @($uploadedInfo | Where-Object { $_.url } | Select-Object -First 1).url
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
                    $attachedfilenameOnly = [System.IO.Path]::GetFileName($attachedFile.name).ToLowerInvariant()
                    if ($link.ToLowerInvariant() -like "*$attachedfilenameOnly*") {
                        Set-PrintandLog -Message "linking attachment $($link.ToLowerInvariant()) => $($attachedFile.url) via $attachedfilenameOnly"
                        $html = $html -replace [regex]::Escape($link), $($attachedFile.url)
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
        if ($null -ne $doc.companyId -and $doc.companyId -ge 1) {
            Set-HuduArticle -id $doc.stub.id -content $updatedHTML -CompanyId $doc.companyId
        } else {
            Set-HuduArticle -id $doc.stub.id -content $updatedHTML
        }
        # Save back
        $doc.ReplacedLinks = Get-LinksFromHTML -htmlContent $updatedHTML -title ($doc.title ?? $doc.localpath) -includeImages $true -suppressOutput $false
        Save-HtmlSnapshot -PageId $doc.id -Title $doc.title -Content $updatedHTML -Suffix "relinked" -OutDir $tmpfolder
        Export-DocPropertyJson -Doc $doc -Property 'ReplacedLinks'
        Set-PrintAndLog "Relinked HTML: $htmlPath" -Color Green
    }
}
Relink-DocumentUploads -docs $stubbedArticles
