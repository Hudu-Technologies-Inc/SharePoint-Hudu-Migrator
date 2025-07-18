$docIDX=0
foreach ($doc in $StubbedArticles) {
    $docIDX=$docIDX+1
    $completionPercentage = Get-PercentDone -Current $docIDX -Total $StubbedArticles.count

    if ($doc.PSObject.Properties["OverrideContent"]) {continue}
    # get attachment / embedded images
    Set-PrintAndLog -message "Starting ul of $($doc.AllAttachments.count) attachments found for $($doc.title)" -Color Green

    # download attachments + upload attachments and atytach to stub

    #Base64EmbeddedImages

    $AttachIDX=0
    foreach ($att in $doc.AllAttachments) {
        $AttachIDX+=1
        $localPath = $att
        $fileSize = (Get-Item $localPath).Length
        $tooLarge = $fileSize -gt 100MB
        $isImage = $att -match '\.(jpg|jpeg|png|gif|bmp)$'

        $record = [PSCustomObject]@{
            FileName           = $att
            Extension          = [IO.Path]::GetExtension($att).ToLower()
            IsImage            = $isImage
            PageId             = $doc.id
            PageTitle          = $doc.title
            SourceUrl          = $null
            LocalPath          = $localPath
            UploadResult       = $null
            HuduArticleId      = $null
            HuduUploadType     = $null
            SuccessDownload    = $true
            AttachmentSize     = $fileSize
            AttachmentTooLarge = $tooLarge
        }
        
        # handle attachment not present
        $exists = Test-Path $localPath
        if (-not $exists) {
            Set-PrintAndLog -Message "Attachment missing on disk: $localPath" -Color Yellow
            $RunSummary.Errors += @{
                Attachment = $att
                Problem    = "File not found on disk."
                Page       = "$($doc.title) (ID: $($doc.id))"
                Article    = "$($doc.Stub.url)"
                doc        = $doc
            }
            continue
        }
        # handle attachment is too large
        if ($true -eq $record.AttachmentTooLarge) {
            $ErrorObject=@{
                Attachment = $record.Filename
                Problem    = "$($record.Filename) is TOO LARGE for Hudu. Manual Action is required. Skipping."
                page       = "Sharepoint page with Id $($doc.id), titled $($doc.title)"
                Article    = "Hudu stub with id $($($doc.stub).id) at $($($doc.stub).url)"
            }
            $RunSummary.Errors = $ErrorObject
            $RunSummary.JobInfo.UploadsErrored+=1
            Write-ErrorObjectsToFile -ErrorObject $ErrorObject -name "Attach-Error-$($record.Filename)"
            continue
        }


        Set-PrintAndLog -message "Downloaded Attachment $AttachIDX of $($doc.AllAttachments.Count) for $($doc.title) - $($attachment.filename ?? "File")" -Color Yellow
        if ($record -and $record.SuccessDownload -and $record.LocalPath) {
            try {
                Set-PrintAndLog -Message "Uploading image: $($record.FileName) => record_id=$($($doc.stub).id) record_type=Article" -Color Green
                $upload=$null
                if ($record.IsImage) {
                    $upload = $((New-HuduPublicPhoto -FilePath $record.LocalPath -record_id $($doc.stub).id -record_type 'Article').public_photo)
                } else {
                    $upload = New-HuduUpload -FilePath $record.LocalPath -record_id $($doc.stub).id -record_type 'Article'
                }

                $mapEntry=[PSCustomObject]@{
                    doc           = $doc.id
                    PageTitle     = $doc.title
                    LocalFile     = $record.FileName
                    HuduUrl       = $upload.url
                    HuduUploadId  = $upload.id
                }
                $AllNewLinks.Add($mapEntry)
                $normalizedFileName = $record.FileName.ToLowerInvariant()
                $ImageMap[$normalizedFileName] = @{
                    Id   = $upload.id
                    Type = if ($record.IsImage) { 'image' } else { 'upload' }
                }
                $upload | Add-Member -NotePropertyName OriginalFilename -NotePropertyValue $record.FileName -Force
                $upload | Add-Member -NotePropertyName MappedUrl -NotePropertyValue $upload.url -Force
                $upload | Add-Member -NotePropertyName UploadType -NotePropertyValue $record.HuduUploadType -Force
                $doc.UploadedFiles.add($upload)                

                $record.UploadResult    = $upload
                $record.HuduUploadType  = $ImageMap[$normalizedFileName].Type
                $record.HuduArticleId   = $($doc.stub).id
                $RunSummary.JobInfo.UploadsCreated += 1
            } catch {
                $ErrorInfo=@{
                    Error       =$_
                    Record      = $record.AttachmentSize ?? 0
                    Message     = "Error During Attachment Upload"
                    Article     = "Hudu Article id $($doc.stub.id) at $($doc.stub.url)"
                    Page        = "Sharepoint page with Id $($doc.id), titled $($doc.title)- $($doc.FullUrl ?? '')"
                }
                $RunSummary.Errors.add($ErrorInfo)
                $RunSummary.JobInfo.UploadsErrored+=1
                Write-ErrorObjectsToFile -Name "$($record.FileName)" -ErrorObject $ErrorInfo
            }
        }
    }
    Write-Progress -Activity "Processing attachments for $($doc.title)" -Status "$completionPercentage%" -PercentComplete $completionPercentage
    Export-DocPropertyJson -Doc $doc -Property 'UploadedFiles'
}

$ImageMap | ConvertTo-Json -depth 45 | Out-File "$(join-path $tmpfolder -ChildPath "imagemap.json")"