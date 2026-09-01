$docIDX=0
$HuduPublicPhotoExtensions = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp")
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
        $exists = Test-Path -LiteralPath $localPath
        $fileSize       = if ($exists) { (Get-Item -LiteralPath $localPath).Length } else { 0 }
        $tooLarge       = [bool]$($exists -and $fileSize -ge 100MB)
        $extension      = [IO.Path]::GetExtension($att).ToLowerInvariant()
        $isImage        = [bool]($HuduPublicPhotoExtensions -contains $extension)

        $record = [PSCustomObject]@{
            FileName           = $att
            Extension          = $extension
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
            Add-RunSummaryError -ErrorObject $errorObject
            continue
        }
        # handle attachment is too large
        if ($true -eq $record.AttachmentTooLarge) {
            Set-PrintAndLog -Message "Attachment is 100 MB or larger; skipping Hudu upload and relying on SharePoint link: $localPath" -Color Yellow
            Add-RunSummaryWarning -Warning @{
                Attachment = "$att"
                Warning    = "$($record.Filename) is 100 MB or larger and was not uploaded to Hudu."
                doc        = "$($doc.title), $($doc.id)"
                Article    = "Hudu stub with id $($($doc.stub).id) at $($($doc.stub).url)"
            }
            continue
        }

        Set-PrintAndLog -message "Applying Attachment $AttachIDX of $($doc.AllAttachments.Count) for $($doc.title) - $($attachment.filename ?? "File")" -Color Yellow
        if ($record -and $record.SuccessDownload -and $record.LocalPath) {
            try {
                Set-PrintAndLog -Message "Uploading image: $($record.FileName) => record_id=$($($doc.stub).id) record_type=Article" -Color Green
                $HuduUpload=$null
                if ($record.IsImage) {
                    $HuduUpload = $((New-HuduPublicPhoto -FilePath $record.LocalPath -record_id $($doc.stub).id -record_type 'Article'))
                    $HuduUpload = $HuduUpload.public_photo ?? $HuduUpload
                    $record.HuduUploadType = 'image'
                } else {
                    $HuduUpload = New-HuduUpload -FilePath $record.LocalPath -record_id $($doc.stub).id -record_type 'Article'
                    $HuduUpload = $HuduUpload.upload ?? $HuduUpload
                    $record.HuduUploadType = 'upload'
                }
                $mappedUrl = Get-HuduUploadArticleUrl -Upload $HuduUpload -UploadType $record.HuduUploadType -OriginalFilename $record.FileName -HuduBaseUrl $HuduBaseURL

                $mapEntry=[PSCustomObject]@{
                    doc           = $doc.id
                    PageTitle     = $doc.title
                    LocalFile     = $record.FileName
                    HuduUrl       = $mappedUrl
                    HuduUploadId  = $HuduUpload.id
                }
                $AllNewLinks.Add($mapEntry)
                $normalizedFileName = $record.FileName.ToLowerInvariant()
                $ImageMap[$normalizedFileName] = @{
                    Id   = $HuduUpload.id
                    Type = $record.HuduUploadType
                }
                $HuduUpload | Add-Member -NotePropertyName OriginalFilename -NotePropertyValue $record.FileName -Force
                $HuduUpload | Add-Member -NotePropertyName MappedUrl -NotePropertyValue $mappedUrl -Force
                $HuduUpload | Add-Member -NotePropertyName UploadType -NotePropertyValue $record.HuduUploadType -Force
                $doc.UploadedFiles.add($HuduUpload)                

                $record.UploadResult    = $HuduUpload
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
                Add-RunSummaryError -ErrorObject $ErrorInfo
                $RunSummary.JobInfo.UploadsErrored+=1
                Write-ErrorObjectsToFile -Name "uploaderr-$($record.FileName)" -ErrorObject $ErrorInfo
            }
        }
    }
    Write-Progress -Activity "Processing attachments for $($doc.title)" -Status "$completionPercentage%" -PercentComplete $completionPercentage
    Export-DocPropertyJson -Doc $doc -Property 'UploadedFiles'
}

$ImageMap | ConvertTo-Json -depth 45 | Out-File "$(join-path $tmpfolder -ChildPath "imagemap.json")"
