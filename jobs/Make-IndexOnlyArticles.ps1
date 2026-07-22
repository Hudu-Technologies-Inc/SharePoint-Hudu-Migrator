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

function Get-SharePointIndexAttributionSourceText {
    param (
        [string]$RelativeFolderPath,
        [array]$Files
    )

    @(
        $RelativeFolderPath
        @($Files | Select-Object -First 10 | ForEach-Object { $_.SiteName })
        @($Files | Select-Object -First 10 | ForEach-Object { $_.DriveName })
        @($Files | Select-Object -First 10 | ForEach-Object { $_.parentDrivePath })
        @($Files | Select-Object -First 10 | ForEach-Object { $_.RelativePath })
        @($Files | Select-Object -First 10 | ForEach-Object { $_.Name })
    ) -join ' '
}

function Resolve-SharePointIndexClientAttribution {
    param (
        [Parameter(Mandatory)] [string]$SourceText
    )

    if (-not $RunSummary.SetupInfo.ClientAttributionAutoApply -or $ClientAttributionMap.Count -lt 1) {
        return $null
    }

    $attributionMatch = Resolve-HuduCompanyFromSharePointAttributionMap `
        -SourceText $SourceText `
        -AttributionMap ($ClientAttributionResolver ?? $ClientAttributionMap) `
        -AutoOnly `
        -AllowUnmatchedClientEntry:$RunSummary.SetupInfo.ClientAttributionCreateMissing `
        -MinScore $RunSummary.SetupInfo.ClientAttributionListItemMinScore `
        -MinGap $RunSummary.SetupInfo.ClientAttributionListItemMinGap

    if (-not $attributionMatch) { return $null }

    try {
        $attributionEntry = Confirm-HuduCompanyForSharePointAttributionMatch `
            -AttributionMatch $attributionMatch `
            -AttributionMap $ClientAttributionMap `
            -CreateMissing:$RunSummary.SetupInfo.ClientAttributionCreateMissing
    } catch {
        Set-PrintAndLog -message "Failed to create Hudu company for client list item '$($attributionMatch.Entry.ClientName)': $($_.Exception.Message)" -Color Red
        return $null
    }

    if (-not $attributionEntry -or -not $attributionEntry.HuduCompanyId) { return $null }

    [PSCustomObject]@{
        Entry = $attributionEntry
        Match = $attributionMatch
    }
}

function Resolve-SharePointIndexClientDesignation {
    param ([array]$Files)

    $sampleFile = @($Files | Select-Object -First 1)[0]
    if (-not $sampleFile) { return $null }

    Resolve-HuduCompanyFromClientDesignationMap `
        -SiteId $sampleFile.SiteId `
        -ListId $sampleFile.sharepointListId `
        -ClientDesignationMap $ClientDesignationMap `
        -UseSiteDesignation:$RunSummary.SetupInfo.ClientAttributionUseSiteDesignations `
        -UseListDesignation:$RunSummary.SetupInfo.ClientAttributionUseListDesignations
}

function Get-IndexOnlyCompanyId {
    param (
        [string]$RelativeFolderPath,
        [array]$Files
    )

    switch ([int]$RunSummary.JobInfo.MigrationDest.Identifier) {
        0 { return $SingleCompanyChoice.id }
        1 { return $null }
        3 {
            $clientDesignation = if ($RunSummary.SetupInfo.PreferClientAttributionOverSiteCompany) {
                Resolve-SharePointIndexClientDesignation -Files $Files
            } else {
                $null
            }

            $clientAttribution = if (-not $clientDesignation -and $RunSummary.SetupInfo.PreferClientAttributionOverSiteCompany) {
                Resolve-SharePointIndexClientAttribution -SourceText (Get-SharePointIndexAttributionSourceText -RelativeFolderPath $RelativeFolderPath -Files $Files)
            } else {
                $null
            }

            if ($clientDesignation) {
                Set-PrintAndLog -message "Assigned index-only folder '$RelativeFolderPath' to per-site client designation '$($clientDesignation.HuduCompanyName)' ($($clientDesignation.Votes)/$($clientDesignation.TotalVotes) votes)." -Color Cyan
                return $clientDesignation.HuduCompanyId
            } elseif ($clientAttribution) {
                Set-PrintAndLog -message "Auto-attributed index-only folder '$RelativeFolderPath' to client list item '$($clientAttribution.Entry.RawTitle)' => Hudu company '$($clientAttribution.Entry.HuduCompanyName)' via '$($clientAttribution.Match.Alias)' ($($clientAttribution.Match.Confidence)%)." -Color Cyan
                return $clientAttribution.Entry.HuduCompanyId
            } else {
                $sampleFile = @($Files | Select-Object -First 1)[0]
                $siteCompany = Resolve-HuduCompanyFromSiteCompanyMap -SiteId $sampleFile.SiteId -SiteName $sampleFile.SiteName -SiteCompanyMap $SiteCompanyMap
                if ($siteCompany -and $siteCompany.HuduCompanyId) {
                    Set-PrintAndLog -message "Assigned index-only folder '$RelativeFolderPath' to per-site Hudu company '$($siteCompany.HuduCompanyName)'." -Color Cyan
                    return $siteCompany.HuduCompanyId
                }

                Set-PrintAndLog -message "No client or per-site Hudu company available for index-only folder '$RelativeFolderPath'; falling back to manual company selection." -Color Yellow
                $sample = @($Files | Select-Object -First 3 | ForEach-Object { $_.Name }) -join ", "
                return (
                    Select-ObjectFromList `
                        -message "Index-only folder: $RelativeFolderPath ($(@($Files).Count) file(s); $sample). Which company to migrate into?" `
                        -objects $Attribution_Options
                ).CompanyId
            }
        }
        default {
            $clientDesignation = Resolve-SharePointIndexClientDesignation -Files $Files
            $clientAttribution = if (-not $clientDesignation) {
                Resolve-SharePointIndexClientAttribution -SourceText (Get-SharePointIndexAttributionSourceText -RelativeFolderPath $RelativeFolderPath -Files $Files)
            } else {
                $null
            }

            if ($clientDesignation) {
                Set-PrintAndLog -message "Assigned index-only folder '$RelativeFolderPath' to per-site client designation '$($clientDesignation.HuduCompanyName)' ($($clientDesignation.Votes)/$($clientDesignation.TotalVotes) votes)." -Color Cyan
                return $clientDesignation.HuduCompanyId
            } elseif ($clientAttribution) {
                Set-PrintAndLog -message "Auto-attributed index-only folder '$RelativeFolderPath' to client list item '$($clientAttribution.Entry.RawTitle)' => Hudu company '$($clientAttribution.Entry.HuduCompanyName)' via '$($clientAttribution.Match.Alias)' ($($clientAttribution.Match.Confidence)%)." -Color Cyan
                return $clientAttribution.Entry.HuduCompanyId
            }

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

    $leafName = if ([string]::IsNullOrWhiteSpace($relativeFolderPath)) { "SharePoint Root" } else { Split-Path -Path $relativeFolderPath -Leaf }
    $articleTitle = "$(Get-SafeTitle $leafName) - File Index"

    if ($RunSummary.SetupInfo.SkipExistingArticles -and ($null -eq $companyId -or $companyId -ge 0)) {
        $existingArticle = Get-HuduExistingArticleByExactName -Title $articleTitle -CompanyId $companyId
        if ($existingArticle) {
            $existingArticleId = $existingArticle.id ?? $existingArticle.Id
            $existingArticleUrl = $existingArticle.url ?? $existingArticle.Url
            Set-PrintAndLog -message "Skipping index-only article '$articleTitle' because Hudu article already exists in target company/global KB: $existingArticleUrl" -Color Yellow

            $RunSummary.JobInfo.ArticlesSkipped++
            $RunSummary.Warnings += @{
                Message     = "Skipped index-only folder because matching Hudu article already exists"
                Title       = $articleTitle
                CompanyId   = $companyId
                ExistingId  = $existingArticleId
                ExistingUrl = $existingArticleUrl
                Folder      = $relativeFolderPath
            }

            [void]$AllNewLinks.Add([PSCustomObject]@{
                PageId    = "index-only:$relativeFolderPath"
                PageTitle = $articleTitle
                HuduUrl   = $existingArticleUrl
                ArticleId = $existingArticleId
            })

            if ($RunSummary.SetupInfo.ResumeFromState) {
                foreach ($file in $files) {
                    if ([string]::IsNullOrWhiteSpace([string]$file.SourceKey)) { continue }

                    Write-SharePointExistingArticleSkipState `
                        -Doc $file `
                        -ExistingArticle $existingArticle `
                        -Message "Skipped because matching Hudu file index article already exists"
                }
            }

            Write-Progress -Activity "Creating index-only articles" -Status "$completionPercentage%" -PercentComplete $completionPercentage
            continue
        }
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

        if ($file.FileTooLarge -or ((Test-Path -LiteralPath $file.LocalPath) -and (Get-Item -LiteralPath $file.LocalPath).Length -ge 100MB)) {
            $file.IndexUploadStatus = "100 MB or larger; use SharePoint link"
            Set-PrintAndLog -message "Index-only file too large for Hudu upload: $($file.LocalPath)" -Color Yellow
            continue
        }

        if (-not (Test-Path -LiteralPath $file.LocalPath)) {
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

    if ($RunSummary.SetupInfo.ResumeFromState) {
        foreach ($file in $files) {
            if ([string]::IsNullOrWhiteSpace([string]$file.SourceKey)) { continue }

            $stateEntry = Write-SharePointMigrationStateEntry `
                -Path $RunSummary.OutputJsonFiles.MigrationState `
                -Item $file `
                -Status Completed `
                -HuduType Article `
                -HuduId ($huduArticle.id ?? $stub.id) `
                -Message "Indexed in Hudu file index article"

            $SharePointMigrationState[$file.SourceKey] = $stateEntry
        }
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
