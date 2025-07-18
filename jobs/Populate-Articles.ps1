$docIDX=0
foreach ($doc in $StubbedArticles) {
    write-host "starting $doc "
    $docIDX=$docIDX+1
    $completionPercentage = Get-PercentDone -Current $docIDX -Total $StubbedArticles.count
    $UploadedAsDoc=$false
    if ($doc.UsingGeneratedHTML) {
        Set-PrintAndLog -message "Using generated content, skipping population" -color DarkMagenta
        continue
    }
    if ([string]::IsNullOrWhiteSpace($doc.ReplacedContent)) {
        $doc.ReplacedContent = "No Content Present"
    }

    Save-HtmlSnapshot -PageId $doc.id -Title $doc.title -Content $doc.RawContent -Suffix "raw" -OutDir $tmpfolder
    $doc.ReplacedContent = compress-html -Html $doc.ReplacedContent 

    $doc.charsTrimmed =  $doc.rawContent.length - $($doc.ReplacedContent).length
    Set-PrintAndLog -Message "Removed $($doc.charsTrimmed) characters of bloat from $($doc.title)" -Color Green
    Save-HtmlSnapshot -PageId $doc.id -Title $doc.title -Content $doc.RawContent -Suffix "raw" -OutDir $tmpfolder
    $FinalContents = $doc.ReplacedContent

    Set-PrintAndLog "Populating Article: $($doc.title) to $($($doc.CompanyId) ?? 'Global KB') with relinked contents" -Color Green

    if ($($doc.ReplacedContent).Length -gt $HUDU_MAX_DOCSIZE) {
        $UploadedAsDoc=$true
        if (-not ($doc.FileTooLarge)) {
            Set-PrintAndLog "Content Length Warning: Content, even after stripping bloat is too large. Safe-Maximum is $HUDU_MAX_DOCSIZE Characters, and this is $($($doc.ReplacedContent).length) chars long! Adding as attached document!"
            $uploadable=$($(Get-HuduUploads | where-object {$_.uploadable_type -eq 'Article' -and $_.uploadable_id -eq $doc.stub.id}) ?? $(New-HuduUpload -FilePath $doc.LocalPath -record_id $($doc.stub).id -record_type 'Article'))
            $Link = "<br><a href='$($uploadable.url ?? $doc.webViewUrl)' target='_blank'>View Original</a>"
            $note = "</p>File was too large to attach to Hudu.</p>"    
            $htmlPath = Join-Path $tmpfolder -ChildPath ("LargeDoc_{0}.html" -f (Get-SafeFilename ([IO.Path]::GetFileNameWithoutExtension($($doc.title)))))
            $doc.NewPath = $htmlPath
            Get-GeneratedAttachmentLinkLargeDocs -sourceFile $doc -outputFile $htmlPath -link $link -note $note
            $doc.replacedContent = Get-Content $doc.NewPath -Raw
            if ($null -ne $uploadable) {$doc.UploadedFiles.add($uploadable)}
            $doc | Add-Member -NotePropertyName OverrideContent -NotePropertyValue $doc.ReplacedContent -Force
            $FinalContents =  $doc.replacedContent
        } else {
            Set-PrintAndLog "Original File is too large. Adding Link to Sharepoint"
            $Link = "<br><a href='$($doc.webViewUrl)' target='_blank'>View In Sharepoint (too large)</a>"
            $note = "</p>File was too large to attach to Hudu.</p>"
            $htmlPath = Join-Path $tmpfolder -ChildPath ("RemoteDoc_{0}.html" -f (Get-SafeFilename ([IO.Path]::GetFileNameWithoutExtension($($doc.title)))))
            $doc.NewPath = $htmlPath
            Get-GeneratedAttachmentLinkLargeDocs -sourceFile $doc -outputFile $htmlPath -link $link -note $note
            $doc.replacedContent = Get-Content $doc.NewPath -Raw
            $doc | Add-Member -NotePropertyName OverrideContent -NotePropertyValue $doc.ReplacedContent -Force
            $FinalContents =  $doc.replacedContent
        }
    } else {
        $FinalContents=$($($doc.ReplacedContent) ?? "unknown contents")
    }
    try {
        if ($null -ne $($doc.CompanyId) -and $($doc.CompanyId) -ne -1) {
            $doc.HuduArticle = $(Set-HuduArticle -ArticleId $($doc.stub).id -Content $FinalContents -name $($($doc.title) ?? "Unknown Title") -CompanyId $($doc.CompanyId)).Article
        } else {
            $doc.HuduArticle = $(Set-HuduArticle -ArticleId $($doc.stub).id -Content $FinalContents -name $($($doc.title) ?? "Unknown Title")).Article
        }
        
    } catch {
        # Handle articles that are too large having an issue during file upload / linking
        $ErrorInfo=@{
            Message="Error Uploading article with content that is too long: $($doc.title)"
            Error=$_
            HuduArticle=$(Get-HuduArticles -id $($doc.stub).id).Article
            doc = "Sharepoint page with Id $($doc.id), titled $($doc.title)- $($doc.FullUrl ?? '')"
            ArticleURL=$($doc.stub.url ?? "URL not found")
        }
        $RunSummary.Errors.add($ErrorInfo)
        $RunSummary.JobInfo.ArticlesErrored+=1
        Write-ErrorObjectsToFile -name "largearticle-$($doc.title)" -ErrorObject $ErrorInfo
        continue
    }

    if ($true -eq $UploadedAsDoc) {
        # Add a warning for articles that are too large being uploaded as linked standalone file
        $RunSummary.Warnings.add(@{
            Warning="Document from page $($doc.title) was too large and was uploaded as standalone HTML File; Please review."
            ArticleURL=$htmlAttachment.Article.url ?? ($doc.stub.url ?? "URL not found")
            PageURL=$doc.FullUrl
        })
        continue
    }

    # Track relinking info. we'll want to relink articles/pages after all are created.
    $Article_Relinking[$doc.Stub.id] = [PSCustomObject]@{
        HuduArticle = $doc.Stub
        doc         = $doc
        content     = $doc.ReplacedContent
        Links       = $doc.AllAttachments
    }
    Write-Progress -Activity "Processing content for $($doc.title)" -Status "$completionPercentage%" -PercentComplete $completionPercentage
}