##### One-off helper, move tagged articles into matching company KBs while preserving folder paths

function ConvertTo-HuduMoveTaggedBool {
    param (
        $Value,
        [bool]$Default = $false
    )

    if ($null -eq $Value) { return $Default }
    if ($Value -is [bool]) { return $Value }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $Default }

    switch -Regex ($text.ToLowerInvariant()) {
        '^(1|true|yes|y|on)$' { return $true }
        '^(0|false|no|n|off)$' { return $false }
        default { return [bool]$Value }
    }
}

$moveTaggedDryRun = ConvertTo-HuduMoveTaggedBool ($HuduMoveTaggedArticlesDryRun ?? $true) $true
$moveTaggedPreserveFolderPath = ConvertTo-HuduMoveTaggedBool ($HuduMoveTaggedArticlesPreserveFolderPath ?? $true) $true
$moveTaggedSkipAlreadyInTargetCompany = ConvertTo-HuduMoveTaggedBool ($HuduMoveTaggedArticlesSkipAlreadyInTargetCompany ?? $true) $true
$moveTaggedDestinationRootFolderName = [string]($HuduMoveTaggedArticlesDestinationRootFolderName ?? "")
$moveTaggedMaxArticles = [int]($HuduMoveTaggedArticlesMaxArticles ?? 0)
$moveTaggedArticleIds = @(
    if ($null -ne $HuduMoveTaggedArticlesArticleIds) {
        @($HuduMoveTaggedArticlesArticleIds | ForEach-Object { [int]$_ })
    } else {
        @()
    }
)
$moveTaggedCompanyIds = @(
    if ($null -ne $HuduMoveTaggedArticlesCompanyIds) {
        @($HuduMoveTaggedArticlesCompanyIds | ForEach-Object { [int]$_ })
    } else {
        @()
    }
)
$moveTaggedSourceCompanyIds = @(
    if ($null -ne $HuduMoveTaggedArticlesSourceCompanyIds) {
        @($HuduMoveTaggedArticlesSourceCompanyIds | ForEach-Object { [string]$_ })
    } else {
        @()
    }
)

$moveTaggedRoot = if (-not [string]::IsNullOrWhiteSpace([string]$workdir)) {
    [string]$workdir
} elseif (-not [string]::IsNullOrWhiteSpace([string]$PSScriptRoot)) {
    Split-Path -Parent $PSScriptRoot
} else {
    (Get-Location).Path
}
$moveTaggedReportPath = if (-not [string]::IsNullOrWhiteSpace([string]$HuduMoveTaggedArticlesReportPath)) {
    if ([System.IO.Path]::IsPathRooted($HuduMoveTaggedArticlesReportPath)) {
        [System.IO.Path]::GetFullPath($HuduMoveTaggedArticlesReportPath)
    } else {
        [System.IO.Path]::GetFullPath((Join-Path $moveTaggedRoot $HuduMoveTaggedArticlesReportPath))
    }
} else {
    Join-Path (Join-Path $moveTaggedRoot "logs") "moved-tagged-articles-to-companies.csv"
}

function Write-HuduMoveTaggedLog {
    param (
        [Parameter(Mandatory)] [string]$Message,
        [string]$Color = "White"
    )

    if (Get-Command Set-PrintAndLog -ErrorAction SilentlyContinue) {
        Set-PrintAndLog -message $Message -Color $Color
    } else {
        Write-Host $Message -ForegroundColor $Color
    }
}

function ConvertTo-HuduMoveTaggedKey {
    param ($Value)

    if ($null -eq $Value) { return "" }
    $text = ([string]$Value).Normalize([Text.NormalizationForm]::FormD).ToLowerInvariant()
    $text = $text -replace '\p{Mn}', ''
    try {
        $text = [System.Web.HttpUtility]::HtmlDecode($text)
    } catch {
        $text = [System.Net.WebUtility]::HtmlDecode($text)
    }
    $text = $text -replace '&', ' and '
    $text = $text -replace '[^a-z0-9]+', ' '
    return ($text -replace '\s+', ' ').Trim()
}

function Get-HuduMoveTaggedTrailingTag {
    param ($Value)

    if ($null -eq $Value) { return $null }
    $match = [regex]::Match([string]$Value, '\((?<tag>[^()]*)\)\s*$')
    if (-not $match.Success) { return $null }

    $tag = $match.Groups['tag'].Value.Trim()
    if ([string]::IsNullOrWhiteSpace($tag)) { return $null }
    return $tag
}

function Get-HuduMoveTaggedArticleObject {
    param ($Article)

    return ($Article.article ?? $Article.Article ?? $Article)
}

function Expand-HuduMoveTaggedArticles {
    param ($InputObject)

    $expanded = [System.Collections.Generic.List[object]]::new()
    foreach ($item in @($InputObject)) {
        if ($null -eq $item) { continue }

        $articleSet = $item.articles ?? $item.Articles
        if ($articleSet) {
            foreach ($wrappedArticle in @($articleSet)) {
                $expanded.Add((Get-HuduMoveTaggedArticleObject $wrappedArticle))
            }
            continue
        }

        $expanded.Add((Get-HuduMoveTaggedArticleObject $item))
    }

    return @($expanded)
}

function Get-HuduMoveTaggedArticleId {
    param ($Article)

    $articleObject = Get-HuduMoveTaggedArticleObject $Article
    return ($articleObject.id ?? $articleObject.Id ?? $articleObject.article_id ?? $articleObject.ArticleId)
}

function Get-HuduMoveTaggedArticleName {
    param ($Article)

    $articleObject = Get-HuduMoveTaggedArticleObject $Article
    return [string]($articleObject.name ?? $articleObject.Name ?? $articleObject.title ?? $articleObject.Title ?? "Untitled Article")
}

function Get-HuduMoveTaggedArticleCompanyId {
    param ($Article)

    $articleObject = Get-HuduMoveTaggedArticleObject $Article
    return ($articleObject.company_id ?? $articleObject.companyId ?? $articleObject.CompanyId ?? $articleObject.company.id ?? $articleObject.Company.Id)
}

function Get-HuduMoveTaggedFolderId {
    param ($Object)

    return ($Object.folder_id ?? $Object.FolderId ?? $Object.folder.id ?? $Object.Folder.Id)
}

function Get-HuduMoveTaggedFolderParentId {
    param ($Folder)

    return ($Folder.parent_folder_id ?? $Folder.ParentFolderId ?? $Folder.parentFolderId)
}

function Get-HuduMoveTaggedFolderCompanyId {
    param ($Folder)

    return ($Folder.company_id ?? $Folder.companyId ?? $Folder.CompanyId ?? $Folder.company.id ?? $Folder.Company.Id)
}

function Get-HuduMoveTaggedFolderName {
    param ($Folder)

    return [string]($Folder.name ?? $Folder.Name)
}

function Get-HuduMoveTaggedCompanyId {
    param ($Company)

    return ($Company.id ?? $Company.Id)
}

function Get-HuduMoveTaggedCompanyName {
    param ($Company)

    return [string]($Company.name ?? $Company.Name)
}

function Get-HuduMoveTaggedCompanyKey {
    param ($CompanyId)

    if ($null -eq $CompanyId -or [string]::IsNullOrWhiteSpace([string]$CompanyId)) {
        return ""
    }

    $parsedCompanyId = 0
    if (-not [int]::TryParse([string]$CompanyId, [ref]$parsedCompanyId) -or $parsedCompanyId -lt 1) {
        return ""
    }

    return [string]$parsedCompanyId
}

function Get-HuduMoveTaggedFolderPath {
    param (
        $Folder,
        [Parameter(Mandatory)] [hashtable]$FolderById
    )

    if (-not $Folder) { return @() }

    $path = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $current = $Folder

    while ($current) {
        $currentId = [string]($current.id ?? $current.Id)
        if ($currentId -and -not $seen.Add($currentId)) { break }

        $name = Get-HuduMoveTaggedFolderName $current
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $path.Insert(0, $name)
        }

        $parentId = Get-HuduMoveTaggedFolderParentId $current
        if (-not $parentId -or -not $FolderById.ContainsKey([string]$parentId)) { break }
        $current = $FolderById[[string]$parentId]
    }

    return @($path)
}

function New-HuduMoveTaggedFolderIndex {
    param ($Folders)

    $folderById = @{}
    $childrenByParent = @{}
    foreach ($folder in @($Folders)) {
        if ($null -eq $folder) { continue }

        $id = [string]($folder.id ?? $folder.Id)
        if ($id) { $folderById[$id] = $folder }

        $parentId = [string](Get-HuduMoveTaggedFolderParentId $folder)
        if ([string]::IsNullOrWhiteSpace($parentId)) { $parentId = "" }
        if (-not $childrenByParent.ContainsKey($parentId)) {
            $childrenByParent[$parentId] = [System.Collections.Generic.List[object]]::new()
        }
        $childrenByParent[$parentId].Add($folder)
    }

    [PSCustomObject]@{
        FolderById       = $folderById
        ChildrenByParent = $childrenByParent
    }
}

function Find-HuduMoveTaggedChildFolder {
    param (
        [Parameter(Mandatory)] [hashtable]$ChildrenByParent,
        $ParentId,
        [Parameter(Mandatory)] [string]$Name
    )

    $parentKey = if ($ParentId) { [string]$ParentId } else { "" }
    if (-not $ChildrenByParent.ContainsKey($parentKey)) { return $null }

    $nameKey = ConvertTo-HuduMoveTaggedKey $Name
    return @($ChildrenByParent[$parentKey] | Where-Object {
        (ConvertTo-HuduMoveTaggedKey (Get-HuduMoveTaggedFolderName $_)) -eq $nameKey
    } | Select-Object -First 1)
}

function Add-HuduMoveTaggedFolderToIndex {
    param (
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)] [hashtable]$FolderById,
        [Parameter(Mandatory)] [hashtable]$ChildrenByParent
    )

    $id = [string]($Folder.id ?? $Folder.Id)
    if ($id) { $FolderById[$id] = $Folder }

    $parentId = [string](Get-HuduMoveTaggedFolderParentId $Folder)
    if ([string]::IsNullOrWhiteSpace($parentId)) { $parentId = "" }
    if (-not $ChildrenByParent.ContainsKey($parentId)) {
        $ChildrenByParent[$parentId] = [System.Collections.Generic.List[object]]::new()
    }
    $ChildrenByParent[$parentId].Add($Folder)
}

function Invoke-HuduMoveTaggedRequest {
    param (
        [Parameter(Mandatory)] [ValidateSet("Get", "Post", "Put")] [string]$Method,
        [Parameter(Mandatory)] [string]$Resource,
        [string]$Body
    )

    if (Get-Command Invoke-HuduRequest -ErrorAction SilentlyContinue) {
        if ($Body) {
            return Invoke-HuduRequest -Method $Method.ToLowerInvariant() -Resource $Resource -Body $Body
        }
        return Invoke-HuduRequest -Method $Method.ToLowerInvariant() -Resource $Resource
    }

    $baseUrl = [string](Get-HuduBaseURL)
    $apiKey = (New-Object PSCredential 'user', (Get-HuduApiKey)).GetNetworkCredential().Password
    if ([string]::IsNullOrWhiteSpace($baseUrl) -or [string]::IsNullOrWhiteSpace($apiKey)) {
        throw "Hudu base URL/API key are not initialized."
    }

    $uri = "{0}{1}" -f $baseUrl.TrimEnd('/'), $Resource
    $headers = @{ 'x-api-key' = $apiKey }
    $params = @{
        Method      = $Method
        Uri         = $uri
        Headers     = $headers
        ErrorAction = 'Stop'
    }
    if ($Body) {
        $params.Body = $Body
        $params.ContentType = 'application/json'
    }

    Invoke-RestMethod @params
}

function New-HuduMoveTaggedFolder {
    param (
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [int]$CompanyId,
        $ParentFolderId
    )

    $command = Get-Command New-HuduFolder -ErrorAction SilentlyContinue
    if ($command) {
        $newFolderParams = @{ Name = $Name }

        if ($ParentFolderId) {
            if ($command.Parameters.ContainsKey('ParentFolderId')) {
                $newFolderParams.ParentFolderId = $ParentFolderId
            } elseif ($command.Parameters.ContainsKey('parent_folder_id')) {
                $newFolderParams.parent_folder_id = $ParentFolderId
            }
        }

        if ($command.Parameters.ContainsKey('CompanyId')) {
            $newFolderParams.CompanyId = $CompanyId
            $created = New-HuduFolder @newFolderParams
            return ($created.folder ?? $created)
        }

        if ($command.Parameters.ContainsKey('company_id')) {
            $newFolderParams.company_id = $CompanyId
            $created = New-HuduFolder @newFolderParams
            return ($created.folder ?? $created)
        }
    }

    $folder = @{
        name       = $Name
        company_id = $CompanyId
    }
    if ($ParentFolderId) {
        $folder.parent_folder_id = $ParentFolderId
    }

    $createdResponse = Invoke-HuduMoveTaggedRequest `
        -Method Post `
        -Resource "/api/v1/folders" `
        -Body (@{ folder = $folder } | ConvertTo-Json -Depth 20)

    return ($createdResponse.folder ?? $createdResponse)
}

function Ensure-HuduMoveTaggedFolderPath {
    param (
        [Parameter(Mandatory)] [string[]]$Path,
        [Parameter(Mandatory)] [int]$CompanyId,
        [Parameter(Mandatory)] [hashtable]$DestinationFolderById,
        [Parameter(Mandatory)] [hashtable]$DestinationChildrenByParent,
        [switch]$DryRun
    )

    if (@($Path).Count -lt 1) { return $null }

    $parentId = $null
    $lastFolder = $null
    $pathSoFar = [System.Collections.Generic.List[string]]::new()
    foreach ($folderName in @($Path)) {
        if ([string]::IsNullOrWhiteSpace($folderName)) { continue }
        $pathSoFar.Add($folderName)

        $existing = Find-HuduMoveTaggedChildFolder -ChildrenByParent $DestinationChildrenByParent -ParentId $parentId -Name $folderName
        if ($existing) {
            $lastFolder = $existing
            $parentId = $existing.id ?? $existing.Id
            continue
        }

        if ($DryRun) {
            $syntheticId = "dryrun:${CompanyId}:$(@($pathSoFar) -join '/')"
            $createdFolder = [PSCustomObject]@{
                id               = $syntheticId
                name             = $folderName
                parent_folder_id = $parentId
                company_id       = $CompanyId
            }
        } else {
            $createdFolder = New-HuduMoveTaggedFolder -Name $folderName -CompanyId $CompanyId -ParentFolderId $parentId
        }

        Add-HuduMoveTaggedFolderToIndex -Folder $createdFolder -FolderById $DestinationFolderById -ChildrenByParent $DestinationChildrenByParent
        $lastFolder = $createdFolder
        $parentId = $createdFolder.id ?? $createdFolder.Id
    }

    return $lastFolder
}

function Set-HuduMoveTaggedArticle {
    param (
        [Parameter(Mandatory)] [int]$ArticleId,
        [Parameter(Mandatory)] [int]$CompanyId,
        $FolderId
    )

    $object = Invoke-HuduMoveTaggedRequest -Method Get -Resource "/api/v1/articles/$ArticleId"
    $article = $object.article ?? $object
    if (-not $article) { throw "Hudu article $ArticleId was not returned." }

    $article.company_id = $CompanyId
    if ($FolderId) {
        $article.folder_id = $FolderId
    } else {
        $article.folder_id = $null
    }

    $body = @{ article = $article } | ConvertTo-Json -Depth 20
    Invoke-HuduMoveTaggedRequest -Method Put -Resource "/api/v1/articles/$ArticleId" -Body $body
}

function Get-HuduMoveTaggedFolderIndexForCompany {
    param (
        $CompanyId,
        [Parameter(Mandatory)] [hashtable]$FolderIndexByCompany
    )

    $companyKey = Get-HuduMoveTaggedCompanyKey $CompanyId
    if ($FolderIndexByCompany.ContainsKey($companyKey)) {
        return $FolderIndexByCompany[$companyKey]
    }

    $folders = if ([string]::IsNullOrWhiteSpace($companyKey)) {
        @(Get-HuduFolders | Where-Object { $null -eq (Get-HuduMoveTaggedFolderCompanyId $_) })
    } else {
        @(Get-HuduFolders -CompanyId ([int]$companyKey))
    }

    $index = New-HuduMoveTaggedFolderIndex -Folders $folders
    $FolderIndexByCompany[$companyKey] = $index
    return $index
}

function New-HuduMoveTaggedCompanyTagIndex {
    param ($Companies)

    $tagIndex = @{}
    foreach ($company in @($Companies)) {
        $companyId = Get-HuduMoveTaggedCompanyId $company
        $companyName = Get-HuduMoveTaggedCompanyName $company
        $companyTag = Get-HuduMoveTaggedTrailingTag $companyName
        $companyTagKey = ConvertTo-HuduMoveTaggedKey $companyTag

        if (-not $companyId -or [string]::IsNullOrWhiteSpace($companyTagKey)) { continue }
        if (-not $tagIndex.ContainsKey($companyTagKey)) {
            $tagIndex[$companyTagKey] = [System.Collections.Generic.List[object]]::new()
        }

        $tagIndex[$companyTagKey].Add([PSCustomObject]@{
            Id     = [int]$companyId
            Name   = $companyName
            Tag    = $companyTag
            TagKey = $companyTagKey
            Object = $company
        })
    }

    return $tagIndex
}

$reportDir = Split-Path -Parent $moveTaggedReportPath
if (-not (Test-Path -LiteralPath $reportDir -PathType Container)) {
    $null = New-Item -ItemType Directory -Path $reportDir -Force
}

Write-HuduMoveTaggedLog -Message "Moving tagged Hudu articles to matching companies. DryRun=$moveTaggedDryRun; PreserveFolderPath=$moveTaggedPreserveFolderPath; DestinationRoot='$moveTaggedDestinationRootFolderName'." -Color Cyan

$allCompanies = @(Get-HuduCompanies)
if ($moveTaggedCompanyIds.Count -gt 0) {
    $companyIdSet = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($companyId in @($moveTaggedCompanyIds)) { [void]$companyIdSet.Add([string]$companyId) }

    $allCompanies = @($allCompanies | Where-Object {
        $companyIdSet.Contains([string](Get-HuduMoveTaggedCompanyId $_))
    })
}

$companyTagIndex = New-HuduMoveTaggedCompanyTagIndex -Companies $allCompanies
$duplicateCompanyTags = @($companyTagIndex.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 })
if ($duplicateCompanyTags.Count -gt 0) {
    Write-HuduMoveTaggedLog -Message "Found $($duplicateCompanyTags.Count) duplicate company tag(s). Articles with those tags will be skipped." -Color Yellow
}

$articles = if ($moveTaggedArticleIds.Count -gt 0) {
    Expand-HuduMoveTaggedArticles ($moveTaggedArticleIds | ForEach-Object { Get-HuduArticles -Id $_ })
} else {
    Expand-HuduMoveTaggedArticles (Get-HuduArticles)
}

if ($moveTaggedSourceCompanyIds.Count -gt 0) {
    $sourceCompanyIdSet = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($companyId in @($moveTaggedSourceCompanyIds)) { [void]$sourceCompanyIdSet.Add([string]$companyId) }

    $articles = @($articles | Where-Object {
        $sourceCompanyId = Get-HuduMoveTaggedArticleCompanyId $_
        $sourceKey = if ($null -eq $sourceCompanyId -or [string]::IsNullOrWhiteSpace([string]$sourceCompanyId)) { "" } else { [string]$sourceCompanyId }
        $sourceCompanyIdSet.Contains($sourceKey)
    })
}

if ($moveTaggedMaxArticles -gt 0) {
    $articles = @($articles | Select-Object -First $moveTaggedMaxArticles)
}

$folderIndexByCompany = @{}
$report = [System.Collections.Generic.List[object]]::new()
$moved = 0
$dryRun = 0
$failed = 0
$skipped = 0
$articleIndex = 0

foreach ($article in @($articles)) {
    $articleIndex++
    $articleId = Get-HuduMoveTaggedArticleId $article
    $articleName = Get-HuduMoveTaggedArticleName $article
    $articleTag = Get-HuduMoveTaggedTrailingTag $articleName
    $articleTagKey = ConvertTo-HuduMoveTaggedKey $articleTag
    $sourceCompanyId = Get-HuduMoveTaggedArticleCompanyId $article
    $sourceCompanyKey = Get-HuduMoveTaggedCompanyKey $sourceCompanyId
    $sourceFolderId = Get-HuduMoveTaggedFolderId $article
    $sourceFolder = $null
    $sourcePath = @()
    $destinationCompanyId = $null
    $destinationCompanyName = $null
    $destinationFolder = $null
    $destinationFolderId = $null
    $destinationPath = [System.Collections.Generic.List[string]]::new()
    $status = $null
    $errorMessage = $null

    try {
        if (-not $articleId) {
            throw "Article id was not detected."
        }

        if ([string]::IsNullOrWhiteSpace($articleTagKey)) {
            $status = "SkippedNoArticleTag"
            $skipped++
            continue
        }

        if (-not $companyTagIndex.ContainsKey($articleTagKey)) {
            $status = "SkippedNoMatchingCompanyTag"
            $skipped++
            continue
        }

        $companyMatches = @($companyTagIndex[$articleTagKey])
        if ($companyMatches.Count -gt 1) {
            $status = "SkippedDuplicateCompanyTag"
            $skipped++
            continue
        }

        $matchedCompany = $companyMatches[0]
        $destinationCompanyId = [int]$matchedCompany.Id
        $destinationCompanyName = [string]$matchedCompany.Name

        if (
            $moveTaggedSkipAlreadyInTargetCompany -and
            -not [string]::IsNullOrWhiteSpace($sourceCompanyKey) -and
            [string]$destinationCompanyId -eq $sourceCompanyKey
        ) {
            $status = "SkippedAlreadyInTargetCompany"
            $skipped++
            continue
        }

        if ($sourceFolderId) {
            $sourceIndex = Get-HuduMoveTaggedFolderIndexForCompany -CompanyId $sourceCompanyId -FolderIndexByCompany $folderIndexByCompany
            if ($sourceIndex.FolderById.ContainsKey([string]$sourceFolderId)) {
                $sourceFolder = $sourceIndex.FolderById[[string]$sourceFolderId]
                $sourcePath = @(Get-HuduMoveTaggedFolderPath -Folder $sourceFolder -FolderById $sourceIndex.FolderById)
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($moveTaggedDestinationRootFolderName)) {
            $destinationPath.Add($moveTaggedDestinationRootFolderName)
        }
        if ($moveTaggedPreserveFolderPath) {
            foreach ($part in @($sourcePath)) { $destinationPath.Add($part) }
        }

        if ($destinationPath.Count -gt 0) {
            $destinationIndex = Get-HuduMoveTaggedFolderIndexForCompany -CompanyId $destinationCompanyId -FolderIndexByCompany $folderIndexByCompany
            $destinationFolder = Ensure-HuduMoveTaggedFolderPath `
                -Path $destinationPath.ToArray() `
                -CompanyId $destinationCompanyId `
                -DestinationFolderById $destinationIndex.FolderById `
                -DestinationChildrenByParent $destinationIndex.ChildrenByParent `
                -DryRun:$moveTaggedDryRun
            $destinationFolderId = $destinationFolder.id ?? $destinationFolder.Id
        }

        if ($moveTaggedDryRun) {
            $status = "DryRunMove"
            $dryRun++
        } else {
            Set-HuduMoveTaggedArticle -ArticleId $articleId -CompanyId $destinationCompanyId -FolderId $destinationFolderId | Out-Null
            $status = "Moved"
            $moved++
            Write-HuduMoveTaggedLog -Message "Moved article $articleIndex/$($articles.Count): '$articleName' to company '$destinationCompanyName' folder '$(@($destinationPath) -join '\')'." -Color Green
        }
    } catch {
        $status = "Failed"
        $errorMessage = $_.Exception.Message
        $failed++
        Write-HuduMoveTaggedLog -Message "Failed to move article '$articleName' ($articleId): $errorMessage" -Color Red
    } finally {
        if (-not $status) {
            $status = "Skipped"
            $skipped++
        }

        $report.Add([PSCustomObject]@{
            Status                 = $status
            ArticleId              = $articleId
            ArticleName            = $articleName
            ArticleTag             = $articleTag
            SourceCompanyId        = $sourceCompanyId
            SourceFolderId         = $sourceFolderId
            SourceFolderPath       = (@($sourcePath) -join '\')
            DestinationCompanyId   = $destinationCompanyId
            DestinationCompanyName = $destinationCompanyName
            DestinationFolderId    = $destinationFolderId
            DestinationPath        = (@($destinationPath) -join '\')
            Error                  = $errorMessage
        })
    }
}

$report | Export-Csv -LiteralPath $moveTaggedReportPath -NoTypeInformation -Encoding UTF8
Write-HuduMoveTaggedLog -Message "Tagged article move complete: $moved moved, $dryRun dry-run move(s), $skipped skipped, $failed failed. Report: $moveTaggedReportPath" -Color Cyan
