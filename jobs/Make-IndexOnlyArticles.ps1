##### Step 5, Index-only files

function ConvertTo-HuduIndexHtmlText {
    param ($Value)

    if ($null -eq $Value) { return "" }
    return [System.Web.HttpUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-HuduIndexHtmlAttribute {
    param ($Value)

    if ($null -eq $Value) { return "" }
    return [System.Web.HttpUtility]::HtmlAttributeEncode([string]$Value)
}

function Get-SharePointIndexSourceUrl {
    param ($File)

    if ($File.webViewUrl) { return $File.webViewUrl }
    return @($File.OriginalLinks)[0]
}

function Format-SharePointIndexDate {
    param ($Value)

    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return "" }

    try {
        return ([datetime]$Value).ToString("yyyy-MM-dd HH:mm")
    } catch {
        return [string]$Value
    }
}

function New-SharePointIndexArticleHtml {
    param (
        [Parameter(Mandatory)] [array]$Files,
        [Parameter(Mandatory)] [string]$Title
    )

    $rows = foreach ($file in ($Files | Sort-Object LocalPath)) {
        $sourceUrl = Get-SharePointIndexSourceUrl -File $file
        $sourceUrlAttr = ConvertTo-HuduIndexHtmlAttribute $sourceUrl
        $fileName = ConvertTo-HuduIndexHtmlText ([System.IO.Path]::GetFileName($file.LocalPath))
        $huduUrl = $file.IndexHuduUrl
        $huduText = if ($file.IndexUploadStatus) { $file.IndexUploadStatus } else { "View in Hudu" }
        $huduCell = if ($huduUrl) {
            "<a href=""$(ConvertTo-HuduIndexHtmlAttribute $huduUrl)"" target=""_blank"">$(ConvertTo-HuduIndexHtmlText $huduText)</a>"
        } else {
            ConvertTo-HuduIndexHtmlText $huduText
        }
        $created = ConvertTo-HuduIndexHtmlText (Format-SharePointIndexDate $file.CreatedDateTime)

@"
      <tr>
        <td><a href="$sourceUrlAttr" target="_blank">$fileName</a></td>
        <td>$huduCell</td>
        <td>$created</td>
      </tr>
"@
    }

    $safeTitle = ConvertTo-HuduIndexHtmlText $Title

@"
<h1>$safeTitle</h1>
<table>
  <thead>
    <tr>
      <th>Original SharePoint Link</th>
      <th>Hudu Link</th>
      <th>Date Added to SharePoint</th>
    </tr>
  </thead>
  <tbody>
$($rows -join "`n")
  </tbody>
</table>
"@
}

function Get-IndexOnlyCompanyId {
    param (
        [string]$RelativeFolderPath,
        [array]$Files
    )

    switch ([int]$RunSummary.JobInfo.MigrationDest.Identifier) {
        0 { return $SingleCompanyChoice.id }
        1 { return $null }
        default {
            $sample = @($Files | Select-Object -First 3 | ForEach-Object { $_.Name }) -join ", "
            return (
                Select-ObjectFromList `
                    -message "Index-only folder: $RelativeFolderPath ($(@($Files).Count) file(s); $sample). Which company to migrate into?" `
                    -objects $Attribution_Options
            ).CompanyId
        }
    }
}

if ($null -eq $IndexOnlyFiles -or $IndexOnlyFiles.Count -eq 0) {
    Set-PrintAndLog -message "No index-only files queued." -Color DarkGray
    return
}

$BaseSitePath = (Get-Item $allSitesfolder).FullName
$FolderResolutionCache = @{}
$indexGroups = $IndexOnlyFiles |
    Where-Object { $_.LocalPath } |
    Group-Object -Property {
        $folderPath = Split-Path -Path $_.LocalPath -Parent
        $folderPath.Substring($BaseSitePath.Length).TrimStart('\')
    }

$groupIndex = 0
foreach ($group in $indexGroups) {
    $groupIndex += 1
    $completionPercentage = Get-PercentDone -Current $groupIndex -Total $indexGroups.Count
    $relativeFolderPath = $group.Name
    $files = @($group.Group)

    $companyId = Get-IndexOnlyCompanyId -RelativeFolderPath $relativeFolderPath -Files $files
    if ($null -ne $companyId -and $companyId -lt 0) {
        Set-PrintAndLog -message "Skipping index-only folder: $relativeFolderPath" -Color Gray
        $RunSummary.Warnings += @{
            Message = "User elected to skip index-only folder"
            Folder  = $relativeFolderPath
        }
        continue
    }

    $huduFolderId = $null
    $key = "$companyId-$relativeFolderPath"
    if (-not [string]::IsNullOrWhiteSpace($relativeFolderPath)) {
        if (-not $FolderResolutionCache.ContainsKey($key)) {
            $folderParts = $relativeFolderPath -split '\\'
            $resolvedFolder = Initialize-HuduFolder -FolderPath $folderParts -CompanyId $companyId

            if ($resolvedFolder) {
                $FolderResolutionCache[$key] = $resolvedFolder
                Set-PrintAndLog -message "Created folder for index path: $relativeFolderPath with ID $($resolvedFolder.id)" -Color Cyan
            } else {
                Set-PrintAndLog -message "Failed to create folder for index path: $relativeFolderPath" -Color Red
            }
        }

        if ($FolderResolutionCache.ContainsKey($key)) {
            $huduFolderId = $FolderResolutionCache[$key].id
        }
    }

    $leafName = if ([string]::IsNullOrWhiteSpace($relativeFolderPath)) { "SharePoint Root" } else { Split-Path -Path $relativeFolderPath -Leaf }
    $articleTitle = "$(Get-SafeTitle $leafName) - File Index"
    Set-PrintAndLog -message "Creating index-only article '$articleTitle' for $($files.Count) file(s)." -Color Yellow

    if ($null -eq $companyId -or $companyId -eq 0) {
        $stub = New-HuduStubArticle -Title $articleTitle -Content "Preparing file index..." -FolderId $huduFolderId
    } else {
        $stub = New-HuduStubArticle -Title $articleTitle -Content "Preparing file index..." -CompanyId $companyId -FolderId $huduFolderId
    }

    if (-not $stub) {
        $errorObject = @{
            Error  = "Error stubbing index-only article"
            Title  = $articleTitle
            Folder = $relativeFolderPath
        }
        Write-ErrorObjectsToFile -name "IndexOnlyStub-$articleTitle" -ErrorObject $errorObject
        [void]$RunSummary.Errors.Add($errorObject)
        $RunSummary.JobInfo.ArticlesErrored++
        continue
    }

    foreach ($file in $files) {
        $file | Add-Member -NotePropertyName CompanyId -NotePropertyValue $companyId -Force
        $file | Add-Member -NotePropertyName Stub -NotePropertyValue $stub -Force
        $file | Add-Member -NotePropertyName HuduFolderId -NotePropertyValue $huduFolderId -Force
        $file | Add-Member -NotePropertyName IndexHuduUrl -NotePropertyValue $null -Force
        $file | Add-Member -NotePropertyName IndexUploadStatus -NotePropertyValue $null -Force

        if ($file.FileTooLarge -or ((Test-Path $file.LocalPath) -and (Get-Item $file.LocalPath).Length -ge 100MB)) {
            $file.IndexUploadStatus = "100 MB or larger; use SharePoint link"
            Set-PrintAndLog -message "Index-only file too large for Hudu upload: $($file.LocalPath)" -Color Yellow
            continue
        }

        if (-not (Test-Path $file.LocalPath)) {
            $file.IndexUploadStatus = "Missing local file"
            Set-PrintAndLog -message "Index-only file missing on disk: $($file.LocalPath)" -Color Yellow
            continue
        }

        try {
            Set-PrintAndLog -message "Uploading index-only file: $($file.LocalPath) => article $($stub.id)" -Color Green
            $huduUpload = New-HuduUpload -FilePath $file.LocalPath -record_id $stub.id -record_type 'Article'
            $huduUpload = $huduUpload.upload ?? $huduUpload
            $huduUrl = $huduUpload.url
            if (-not $huduUrl -and $huduUpload.id) {
                $huduUrl = "$HuduBaseURL/file/$($huduUpload.id)"
            }

            $file.IndexHuduUrl = $huduUrl
            $file.IndexUploadStatus = [System.IO.Path]::GetFileName($file.LocalPath)
            $huduUpload | Add-Member -NotePropertyName OriginalFilename -NotePropertyValue $file.LocalPath -Force
            $huduUpload | Add-Member -NotePropertyName MappedUrl -NotePropertyValue $huduUrl -Force
            [void]$file.UploadedFiles.Add($huduUpload)

            [void]$AllNewLinks.Add([PSCustomObject]@{
                PageId       = $file.id
                PageTitle    = $file.title
                HuduUrl      = $huduUrl
                ArticleId    = $stub.id
                OriginalPath = $file.LocalPath
            })
            $RunSummary.JobInfo.UploadsCreated++
        } catch {
            $file.IndexUploadStatus = "Failed to upload"
            $errorObject = @{
                Error   = $_
                Message = "Error uploading index-only file"
                File    = $file.LocalPath
                Article = "Hudu article id $($stub.id) at $($stub.url)"
            }
            [void]$RunSummary.Errors.Add($errorObject)
            $RunSummary.JobInfo.UploadsErrored++
            Write-ErrorObjectsToFile -Name "IndexOnlyUpload-$($file.title)" -ErrorObject $errorObject
        }
    }

    $indexHtml = New-SharePointIndexArticleHtml -Files $files -Title $articleTitle
    try {
        if ($null -ne $companyId -and $companyId -ge 1) {
            $huduArticle = Set-HuduArticle -ArticleId $stub.id -Content $indexHtml -name $articleTitle -CompanyId $companyId
        } else {
            $huduArticle = Set-HuduArticle -ArticleId $stub.id -Content $indexHtml -name $articleTitle
        }
        $huduArticle = $huduArticle.Article ?? $huduArticle
    } catch {
        $errorObject = @{
            Error   = $_
            Message = "Error updating index-only article content"
            Article = "Hudu article id $($stub.id) at $($stub.url)"
            Folder  = $relativeFolderPath
        }
        [void]$RunSummary.Errors.Add($errorObject)
        $RunSummary.JobInfo.ArticlesErrored++
        Write-ErrorObjectsToFile -Name "IndexOnlyArticle-$articleTitle" -ErrorObject $errorObject
        continue
    }

    $indexArticle = [PSCustomObject]@{
        Title              = $articleTitle
        RelativeFolderPath = $relativeFolderPath
        CompanyId          = $companyId
        Stub               = $stub
        HuduArticle        = $huduArticle
        Files              = $files
        Content            = $indexHtml
    }
    [void]$IndexOnlyArticles.Add($indexArticle)
    $RunSummary.JobInfo.ArticlesCreated++
    $RunSummary.JobInfo.LinksCreated++
    [void]$AllNewLinks.Add([PSCustomObject]@{
        PageId    = "index-only:$relativeFolderPath"
        PageTitle = $articleTitle
        HuduUrl   = $stub.url
        ArticleId = $stub.id
    })

    Save-HtmlSnapshot -PageId $stub.id -Title $articleTitle -Content $indexHtml -Suffix "index-only" -OutDir $tmpfolder
    Write-Progress -Activity "Creating index-only articles" -Status "$completionPercentage%" -PercentComplete $completionPercentage
}
