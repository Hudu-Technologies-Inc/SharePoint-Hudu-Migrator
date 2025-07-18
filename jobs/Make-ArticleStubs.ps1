##### Step 5, Stub Articles
$FolderResolutionCache = @{}
$BaseSitePath = (Get-Item $allSitesfolder).FullName
Set-PrintAndLog -message "BaseSitePath set to $BaseSitePath" -Color Cyan
$docIDX=0
foreach ($doc in $successConverted | Where-Object { $_.PSObject.Properties.Name -contains 'ContentPreview' -and $_.ContentPreview }) {
    $docIDX += 1
    $completionPercentage = Get-PercentDone -Current $docIDX -Total $Sourcedocs.count

    # Determine CompanyId
    switch ([int]$RunSummary.JobInfo.MigrationDest.Identifier) {
        0 { $doc.CompanyId = $SingleCompanyChoice.id }
        1 { $doc.CompanyId = $null }
        default {
            $doc.CompanyId = (
                Select-ObjectFromList `
                    -message "Migrating Article: $($doc.ContentPreview ?? "no preview")... Which company to migrate into?" `
                    -objects $Attribution_Options
            ).CompanyId
        }
    }

    # Resolve relative folder path
    $relativeFolderPath = $null
    if ($doc.LocalPath -and $BaseSitePath) {
        $relativeFolderPath = Split-Path -Path $doc.LocalPath -Parent
        $relativeFolderPath = $relativeFolderPath.Substring($BaseSitePath.Length).TrimStart('\')
    }

    # Build key and resolve/create folder via cache
    $key = "$($doc.CompanyId)-$relativeFolderPath"
    if (-not $FolderResolutionCache.ContainsKey($key) -and $relativeFolderPath) {
        $folderParts = $relativeFolderPath -split '\\'
        $resolvedFolder = Initialize-HuduFolder -FolderPath $folderParts -CompanyId $doc.CompanyId

        if ($resolvedFolder) {
            $FolderResolutionCache[$key] = $resolvedFolder
            Set-PrintAndLog -message "Created folder for path: $relativeFolderPath with ID $($resolvedFolder.id)" -Color Cyan
        } else {
            Set-PrintAndLog -message "Failed to create folder for: $relativeFolderPath" -Color Red
        }
    }

    if ($FolderResolutionCache.ContainsKey($key)) {
        $doc.HuduFolder = $FolderResolutionCache[$key]
        $doc.HuduFolderId = $doc.HuduFolder.id
    }

    # Stub article
    if ($null -eq $doc.CompanyId -or $doc.CompanyId -eq 0) {
        Set-PrintAndLog -message "Stubbing global KB article" -Color Yellow
        $doc.stub = New-HuduStubArticle -Title $doc.title -Content "$($doc.ContentPreview)" -FolderId $doc.HuduFolderId
    }
    elseif ($doc.CompanyId -lt 0) {
        Set-PrintAndLog -message "Skipping doc/article transfer for $($doc.title)" -Color Gray
        $RunSummary.Warnings += @{
            Message     = "User elected to skip doc/article transfer for $($doc.title)"
            docSkipped  = "doc with ID $($doc.id), titled $($doc.title) was skipped. $($doc.FullUrl ?? '')"
        }
        $RunSummary.JobInfo.Skipped++
        continue
    }
    else {
        Set-PrintAndLog -message "Stubbing KB article for Hudu company ID: $($doc.CompanyId)" -Color Yellow
        $doc.stub = New-HuduStubArticle -Title $doc.title -Content "$($doc.ContentPreview)" -CompanyId $doc.CompanyId -FolderId $doc.HuduFolderId
    }

    # Post-processing
    Set-PrintAndLog -message "Article $($doc.title)  with id $($doc.stub.id); $($doc.stub | ConvertTo-Json -Depth 3)" -Color Green

    if (-not $doc.stub) {
        $ErrorObject = @{
            Error = "Error stubbing article with id $($doc.id), title $($doc.title)"
        }
        Write-ErrorObjectsToFile -name "Stub-$($doc.title)" -ErrorObject $ErrorObject
        $RunSummary.Errors.Add($ErrorObject)
        $RunSummary.JobInfo.ArticlesErrored++
        continue
    }

    $RunSummary.JobInfo.ArticlesCreated++
    $RunSummary.JobInfo.LinksCreated++
    $AllNewLinks.Add([PSCustomObject]@{
        PageId    = $doc.id
        PageTitle = $doc.title
        HuduUrl   = $doc.stub.url
        ArticleId = $doc.stub.id
    })

    $StubbedArticles += $doc
    Write-Progress -Activity "Stubbing $($doc.title)" -Status "$completionPercentage%" -PercentComplete $completionPercentage
}
