Set-PrintAndLog -message "Reading converted file contents into each file entry and generating preview" -color DarkMagenta
$ReadContentsIdx=0
# Read raw contents for linking later
# Use b64-removed content for preview now
foreach ($file in $successConverted) {
    Write-Host "Checking file: $($file.LocalPath); NewPath: $($file.NewPath)" -ForegroundColor Cyan
    $ReadContentsIdx=$ReadContentsIdx+1
    $completionPercentage = Get-PercentDone -Current $ReadContentsIdx -Total $successConverted.count
    if ($file.NewPath -and (Test-Path $file.NewPath)) {
        $file.ContentPreview = $(Get-ArticlePreviewBlock -Title $file.title -docId $file.id -content $file.ReplacedContent -MaxLength $RunSummary.SetupInfo.PreviewLength)
        $file.Links = $(Get-LinksFromHTML -htmlContent $file.ReplacedContent -title $file.title ?? $file.localpath -includeImages $true -suppressOutput $false) 
    } elseif ($true -eq $file.UsingGeneratedHTML -and ($null -ne $file.RawContent)) {
        $file.ContentPreview = $(Get-ArticlePreviewBlock -Title $file.title -docId $file.id -content $file.ReplacedContent -MaxLength $RunSummary.SetupInfo.PreviewLength)
        $file.Links = $(Get-LinksFromHTML -htmlContent $file.ReplacedContent -title $file.title ?? $file.localpath -includeImages $true -suppressOutput $false) 
    } else {
        continue
    }
    Export-DocPropertyJson -Doc $file -Property 'AllAttachments'
    Export-DocPropertyJson -Doc $file -Property 'ReplacedLinks'

    Write-Progress -Activity "Reading Converted File Contents and Generating Previews, reading raw links." -Status "$completionPercentage%" -PercentComplete $completionPercentage
}
Set-PrintAndLog -message "Writing out converted file definitions to $("$($RunSummary.OutputJsonFiles.ConvertedFiles)")...!" -color DarkMagenta

$successConverted | ConvertTo-Json -Depth 45 | Out-File "$($RunSummary.OutputJsonFiles.ConvertedFiles)"

