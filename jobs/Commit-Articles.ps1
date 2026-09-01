$docIDX = 0
foreach ($doc in $StubbedArticles) {
    $docIDX += 1
    $completionPercentage = Get-PercentDone -Current $docIDX -Total $StubbedArticles.Count

    if (-not $doc -or -not $doc.stub -or -not $doc.stub.id) {
        Set-PrintAndLog -message "Skipping final article commit for item without a Hudu stub: $($doc.title ?? $doc.LocalPath ?? $doc.Name ?? 'unknown')" -Color DarkGray
        continue
    }

    $finalContents = if ($doc.PSObject.Properties["FinalContent"] -and -not [string]::IsNullOrWhiteSpace([string]$doc.FinalContent)) {
        [string]$doc.FinalContent
    } elseif (-not [string]::IsNullOrWhiteSpace([string]$doc.ReplacedContent)) {
        [string]$doc.ReplacedContent
    } else {
        "No Content Present"
    }

    Set-PrintAndLog -message "Committing final HTML for article: $($doc.title) ($($doc.stub.id))" -Color Green

    try {
        $huduArticle = $null
        if ($null -ne $doc.CompanyId -and $doc.CompanyId -ge 1) {
            $huduArticle = Set-HuduArticle -ArticleId $doc.stub.id -Content $finalContents -name ($doc.title ?? "Unknown Title") -CompanyId $doc.CompanyId
        } else {
            $huduArticle = Set-HuduArticle -ArticleId $doc.stub.id -Content $finalContents -name ($doc.title ?? "Unknown Title")
        }
        $huduArticle = $huduArticle.Article ?? $huduArticle
        $doc | Add-Member -NotePropertyName HuduArticle -NotePropertyValue $huduArticle -Force
    } catch {
        $huduArticle = $null
        try {
            $huduArticle = Get-HuduArticles -id $doc.stub.id
            $huduArticle = $huduArticle.Article ?? $huduArticle
        } catch {}

        $errorInfo = @{
            Message    = "Error committing final HTML to Hudu article: $($doc.title)"
            Error      = $_
            HuduArticle = $huduArticle
            Doc        = "SharePoint item with Id $($doc.id), titled $($doc.title)- $($doc.FullUrl ?? $doc.webViewUrl ?? '')"
            ArticleURL = $doc.stub.url ?? "URL not found"
        }
        $RunSummary.Errors.Add($errorInfo)
        $RunSummary.JobInfo.ArticlesErrored += 1
        Write-ErrorObjectsToFile -name "finalarticle-$($doc.title)" -ErrorObject $errorInfo
        continue
    }

    if ($RunSummary.SetupInfo.ResumeFromState -and -not [string]::IsNullOrWhiteSpace([string]$doc.SourceKey)) {
        $stateEntry = Write-SharePointMigrationStateEntry `
            -Path $RunSummary.OutputJsonFiles.MigrationState `
            -Item $doc `
            -Status Completed `
            -HuduType Article `
            -HuduId ($doc.HuduArticle.id ?? $doc.stub.id) `
            -Message "Final article HTML committed successfully"

        $SharePointMigrationState[$doc.SourceKey] = $stateEntry
    }

    Write-Progress -Activity "Committing final article HTML for $($doc.title)" -Status "$completionPercentage%" -PercentComplete $completionPercentage
}
