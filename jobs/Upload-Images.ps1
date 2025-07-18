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
        $localPath      = $att
        $fileSize       = (Get-Item $localPath).Length
        $tooLarge       = [bool]$($fileSize -gt 100MB)
        $isImage        = [bool]$($att -match '\.(jpg|jpeg|png)$')
        $exists = Test-Path $localPath

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
            exists             = $exists
        }
        # handle image/attachment doesnt exist
        if (-not $exists) {
            Set-PrintAndLog -Message "Attachment missing on disk: $localPath" -Color Yellow
            $errorObject=@{
                Attachment = "$att"
                Problem    = "File not found on disk."
                doc        = "$($doc.title), $($doc.id)"
                Article    = "Hudu stub with id $($($doc.stub).id) at $($($doc.stub).url)"
            }
            Write-ErrorObjectsToFile -ErrorObject $errorObject -name "Nofile-$($att)" -color Red
            $RunSummary.Errors += $errorObject
            continue
        }
        # handle attachment is too large
        if ($true -eq $record.AttachmentTooLarge) {
            $ErrorObject=@{
                Attachment = "$att"
                Problem    = "$($record.Filename) is TOO LARGE for Hudu. Manual Action is required. Skipping."
                doc        = "$($doc.title), $($doc.id)"
                Article    = "Hudu stub with id $($($doc.stub).id) at $($($doc.stub).url)"
            }
            Write-ErrorObjectsToFile -ErrorObject $ErrorObject -name "Attach-Error-$($record.Filename)"
            $RunSummary.Errors += $errorObject
            continue
        }

        Set-PrintAndLog -message "Applying Attachment $AttachIDX of $($doc.AllAttachments.Count) for $($doc.title) - $($attachment.filename ?? "File")" -Color Yellow
        if ($record -and $record.SuccessDownload -and $record.LocalPath) {
            try {
                Set-PrintAndLog -Message "Uploading image: $($record.FileName) => record_id=$($($doc.stub).id) record_type=Article" -Color Green
                $HuduUpload=$null
                if ($record.IsImage) {
                    $HuduUpload = $((New-HuduPublicPhoto -FilePath $record.LocalPath -record_id $($doc.stub).id -record_type 'Article').public_photo)
                } else {
                    $HuduUpload = New-HuduUpload -FilePath $record.LocalPath -record_id $($doc.stub).id -record_type 'Article'
                }

                $mapEntry=[PSCustomObject]@{
                    doc           = $doc.id
                    PageTitle     = $doc.title
                    LocalFile     = $record.FileName
                    HuduUrl       = $HuduUpload.url
                    HuduUploadId  = $HuduUpload.id
                }
                $AllNewLinks.Add($mapEntry)
                $normalizedFileName = $record.FileName.ToLowerInvariant()
                $ImageMap[$normalizedFileName] = @{
                    Id   = $HuduUpload.id
                    Type = if ($record.IsImage) { 'image' } else { 'upload' }
                }
                $HuduUpload | Add-Member -NotePropertyName OriginalFilename -NotePropertyValue $record.FileName -Force
                $HuduUpload | Add-Member -NotePropertyName MappedUrl -NotePropertyValue $HuduUpload.url -Force
                $HuduUpload | Add-Member -NotePropertyName UploadType -NotePropertyValue $record.HuduUploadType -Force
                $doc.UploadedFiles.add($HuduUpload)                

                $record.UploadResult    = $HuduUpload
                $record.HuduUploadType  = $ImageMap[$normalizedFileName].Type
                $record.HuduArticleId   = $($doc.stub).id
                $RunSummary.JobInfo.UploadsCreated += 1
            } catch {
                $ErrorInfo=@{
                    Error       = $_
                    Record      = $record.AttachmentSize ?? 0
                    Message     = "Error During Attachment Upload"
                    Article     = "Hudu Article id $($doc.stub.id) at $($doc.stub.url)"
                    Doc         = "Sharepoint doc with Id $($doc.id), titled $($doc.title)- $($doc.FullUrl ?? '')"
                }
                $RunSummary.Errors.add($ErrorInfo)
                $RunSummary.JobInfo.UploadsErrored+=1
                Write-ErrorObjectsToFile -Name "uploaderr-$($record.FileName)" -ErrorObject $ErrorInfo
            }
        }
    }
    Write-Progress -Activity "Processing attachments for $($doc.title)" -Status "$completionPercentage%" -PercentComplete $completionPercentage
    Export-DocPropertyJson -Doc $doc -Property 'UploadedFiles'
}

$ImageMap | ConvertTo-Json -depth 45 | Out-File "$(join-path $tmpfolder -ChildPath "imagemap.json")"