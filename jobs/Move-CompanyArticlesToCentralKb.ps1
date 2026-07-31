##### One-off helper, move company KB articles into central KB while preserving folder paths

$moveCentralDryRun = [bool]($HuduMoveArticlesToCentralDryRun ?? $true)
$moveCentralSourceCompanyId = [int]($HuduMoveArticlesToCentralSourceCompanyId ?? 0)
$moveCentralDestinationRootFolderName = [string]($HuduMoveArticlesToCentralDestinationRootFolderName ?? "")
$moveCentralPreserveFolderPath = [bool]($HuduMoveArticlesToCentralPreserveFolderPath ?? $true)
$moveCentralIncludeFolderDescendants = [bool]($HuduMoveArticlesToCentralIncludeFolderDescendants ?? $true)
$moveCentralMaxArticles = [int]($HuduMoveArticlesToCentralMaxArticles ?? 0)
$moveCentralArticleIds = @(
    if ($null -ne $HuduMoveArticlesToCentralArticleIds) {
        @($HuduMoveArticlesToCentralArticleIds | ForEach-Object { [int]$_ })
    } else {
        @()
    }
)
$moveCentralSourceFolderIds = @(
    if ($null -ne $HuduMoveArticlesToCentralSourceFolderIds) {
        @($HuduMoveArticlesToCentralSourceFolderIds | ForEach-Object { [int]$_ })
    } else {
        @()
    }
)

$moveCentralRoot = if (-not [string]::IsNullOrWhiteSpace([string]$workdir)) {
    [string]$workdir
} elseif (-not [string]::IsNullOrWhiteSpace([string]$PSScriptRoot)) {
    Split-Path -Parent $PSScriptRoot
} else {
    (Get-Location).Path
}
$moveCentralReportPath = if (-not [string]::IsNullOrWhiteSpace([string]$HuduMoveArticlesToCentralReportPath)) {
    if ([System.IO.Path]::IsPathRooted($HuduMoveArticlesToCentralReportPath)) {
        [System.IO.Path]::GetFullPath($HuduMoveArticlesToCentralReportPath)
    } else {
        [System.IO.Path]::GetFullPath((Join-Path $moveCentralRoot $HuduMoveArticlesToCentralReportPath))
    }
} else {
    Join-Path (Join-Path $moveCentralRoot "logs") "moved-company-articles-to-central-kb.csv"
}

function Write-HuduMoveCentralLog {
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

function ConvertTo-HuduMoveCentralKey {
    param ($Value)

    if ($null -eq $Value) { return "" }
    $text = ([string]$Value).Normalize([Text.NormalizationForm]::FormD).ToLowerInvariant()
    $text = $text -replace '\p{Mn}', ''
    $text = [System.Web.HttpUtility]::HtmlDecode($text)
    $text = $text -replace '&', ' and '
    $text = $text -replace '[^a-z0-9]+', ' '
    return ($text -replace '\s+', ' ').Trim()
}

function Get-HuduMoveCentralArticleObject {
    param ($Article)

    return ($Article.article ?? $Article.Article ?? $Article)
}

function Get-HuduMoveCentralArticleId {
    param ($Article)

    $articleObject = Get-HuduMoveCentralArticleObject $Article
    return ($articleObject.id ?? $articleObject.Id)
}

function Get-HuduMoveCentralFolderId {
    param ($Object)

    return ($Object.folder_id ?? $Object.FolderId ?? $Object.folder.id ?? $Object.Folder.Id)
}

function Get-HuduMoveCentralFolderParentId {
    param ($Folder)

    return ($Folder.parent_folder_id ?? $Folder.ParentFolderId)
}

function Get-HuduMoveCentralFolderCompanyId {
    param ($Folder)

    return ($Folder.company_id ?? $Folder.CompanyId)
}

function Get-HuduMoveCentralFolderName {
    param ($Folder)

    return [string]($Folder.name ?? $Folder.Name)
}

function Get-HuduMoveCentralFolderPath {
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

        $name = Get-HuduMoveCentralFolderName $current
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $path.Insert(0, $name)
        }

        $parentId = Get-HuduMoveCentralFolderParentId $current
        if (-not $parentId -or -not $FolderById.ContainsKey([string]$parentId)) { break }
        $current = $FolderById[[string]$parentId]
    }

    return @($path)
}

function New-HuduMoveCentralFolderIndex {
    param ($Folders)

    $folderById = @{}
    $childrenByParent = @{}
    foreach ($folder in @($Folders)) {
        $id = [string]($folder.id ?? $folder.Id)
        if ($id) { $folderById[$id] = $folder }

        $parentId = [string](Get-HuduMoveCentralFolderParentId $folder)
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

function Find-HuduMoveCentralChildFolder {
    param (
        [Parameter(Mandatory)] [hashtable]$ChildrenByParent,
        $ParentId,
        [Parameter(Mandatory)] [string]$Name
    )

    $parentKey = if ($ParentId) { [string]$ParentId } else { "" }
    if (-not $ChildrenByParent.ContainsKey($parentKey)) { return $null }

    $nameKey = ConvertTo-HuduMoveCentralKey $Name
    return @($ChildrenByParent[$parentKey] | Where-Object {
        (ConvertTo-HuduMoveCentralKey (Get-HuduMoveCentralFolderName $_)) -eq $nameKey
    } | Select-Object -First 1)
}

function Add-HuduMoveCentralFolderToIndex {
    param (
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)] [hashtable]$FolderById,
        [Parameter(Mandatory)] [hashtable]$ChildrenByParent
    )

    $id = [string]($Folder.id ?? $Folder.Id)
    if ($id) { $FolderById[$id] = $Folder }

    $parentId = [string](Get-HuduMoveCentralFolderParentId $Folder)
    if ([string]::IsNullOrWhiteSpace($parentId)) { $parentId = "" }
    if (-not $ChildrenByParent.ContainsKey($parentId)) {
        $ChildrenByParent[$parentId] = [System.Collections.Generic.List[object]]::new()
    }
    $ChildrenByParent[$parentId].Add($Folder)
}

function Ensure-HuduMoveCentralFolderPath {
    param (
        [Parameter(Mandatory)] [string[]]$Path,
        [Parameter(Mandatory)] [hashtable]$CentralFolderById,
        [Parameter(Mandatory)] [hashtable]$CentralChildrenByParent,
        [switch]$DryRun
    )

    if (@($Path).Count -lt 1) { return $null }

    $parentId = $null
    $lastFolder = $null
    $pathSoFar = [System.Collections.Generic.List[string]]::new()
    foreach ($folderName in @($Path)) {
        if ([string]::IsNullOrWhiteSpace($folderName)) { continue }
        $pathSoFar.Add($folderName)

        $existing = Find-HuduMoveCentralChildFolder -ChildrenByParent $CentralChildrenByParent -ParentId $parentId -Name $folderName
        if ($existing) {
            $lastFolder = $existing
            $parentId = $existing.id ?? $existing.Id
            continue
        }

        if ($DryRun) {
            $syntheticId = "dryrun:$(@($pathSoFar) -join '/')"
            $createdFolder = [PSCustomObject]@{
                id               = $syntheticId
                name             = $folderName
                parent_folder_id = $parentId
                company_id       = $null
            }
        } else {
            $newFolderParams = @{ Name = $folderName }
            if ($parentId) { $newFolderParams.ParentFolderId = $parentId }
            $created = New-HuduFolder @newFolderParams
            $createdFolder = $created.folder ?? $created
        }

        Add-HuduMoveCentralFolderToIndex -Folder $createdFolder -FolderById $CentralFolderById -ChildrenByParent $CentralChildrenByParent
        $lastFolder = $createdFolder
        $parentId = $createdFolder.id ?? $createdFolder.Id
    }

    return $lastFolder
}

function Invoke-HuduMoveCentralRequest {
    param (
        [Parameter(Mandatory)] [ValidateSet("Get", "Put")] [string]$Method,
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

function Set-HuduMoveCentralArticle {
    param (
        [Parameter(Mandatory)] [int]$ArticleId,
        $FolderId
    )

    $object = Invoke-HuduMoveCentralRequest -Method Get -Resource "/api/v1/articles/$ArticleId"
    $article = $object.article ?? $object
    if (-not $article) { throw "Hudu article $ArticleId was not returned." }

    $article.company_id = $null
    if ($FolderId) {
        $article.folder_id = $FolderId
    } else {
        $article.folder_id = $null
    }

    $body = @{ article = $article } | ConvertTo-Json -Depth 20
    Invoke-HuduMoveCentralRequest -Method Put -Resource "/api/v1/articles/$ArticleId" -Body $body
}

if ($moveCentralSourceCompanyId -lt 1 -and $moveCentralArticleIds.Count -lt 1) {
    throw "Set `$HuduMoveArticlesToCentralSourceCompanyId or `$HuduMoveArticlesToCentralArticleIds before running this job."
}

$reportDir = Split-Path -Parent $moveCentralReportPath
if (-not (Test-Path -LiteralPath $reportDir -PathType Container)) {
    $null = New-Item -ItemType Directory -Path $reportDir -Force
}

Write-HuduMoveCentralLog -Message "Moving Hudu articles to central KB. DryRun=$moveCentralDryRun; SourceCompanyId=$moveCentralSourceCompanyId; DestinationRoot='$moveCentralDestinationRootFolderName'; PreserveFolderPath=$moveCentralPreserveFolderPath." -Color Cyan

$sourceFolders = if ($moveCentralSourceCompanyId -gt 0) { @(Get-HuduFolders -CompanyId $moveCentralSourceCompanyId) } else { @() }
$centralFolders = @(Get-HuduFolders | Where-Object { $null -eq (Get-HuduMoveCentralFolderCompanyId $_) })
$sourceIndex = New-HuduMoveCentralFolderIndex -Folders $sourceFolders
$centralIndex = New-HuduMoveCentralFolderIndex -Folders $centralFolders

$articles = if ($moveCentralArticleIds.Count -gt 0) {
    @($moveCentralArticleIds | ForEach-Object { Get-HuduArticles -Id $_ } | ForEach-Object { Get-HuduMoveCentralArticleObject $_ })
} else {
    @(Get-HuduArticles -CompanyId $moveCentralSourceCompanyId | ForEach-Object { Get-HuduMoveCentralArticleObject $_ })
}

if ($moveCentralSourceFolderIds.Count -gt 0) {
    $folderIdsToInclude = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($folderId in @($moveCentralSourceFolderIds)) {
        [void]$folderIdsToInclude.Add([string]$folderId)
        if ($moveCentralIncludeFolderDescendants) {
            $queue = [System.Collections.Generic.Queue[object]]::new()
            $queue.Enqueue([string]$folderId)
            while ($queue.Count -gt 0) {
                $currentId = [string]$queue.Dequeue()
                if (-not $sourceIndex.ChildrenByParent.ContainsKey($currentId)) { continue }
                foreach ($child in @($sourceIndex.ChildrenByParent[$currentId])) {
                    $childId = [string]($child.id ?? $child.Id)
                    if ($childId -and $folderIdsToInclude.Add($childId)) {
                        $queue.Enqueue($childId)
                    }
                }
            }
        }
    }

    $articles = @($articles | Where-Object {
        $folderId = Get-HuduMoveCentralFolderId $_
        $folderId -and $folderIdsToInclude.Contains([string]$folderId)
    })
}

if ($moveCentralMaxArticles -gt 0) {
    $articles = @($articles | Select-Object -First $moveCentralMaxArticles)
}

$report = [System.Collections.Generic.List[object]]::new()
$moved = 0
$dryRun = 0
$failed = 0
$skipped = 0
$articleIndex = 0

foreach ($article in @($articles)) {
    $articleIndex++
    $articleId = Get-HuduMoveCentralArticleId $article
    $articleName = [string]($article.name ?? $article.Name ?? $article.title ?? $article.Title ?? "Untitled Article")
    $sourceFolderId = Get-HuduMoveCentralFolderId $article
    $sourceFolder = if ($sourceFolderId -and $sourceIndex.FolderById.ContainsKey([string]$sourceFolderId)) {
        $sourceIndex.FolderById[[string]$sourceFolderId]
    } else {
        $null
    }

    $sourcePath = @(Get-HuduMoveCentralFolderPath -Folder $sourceFolder -FolderById $sourceIndex.FolderById)
    $destinationPath = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($moveCentralDestinationRootFolderName)) {
        $destinationPath.Add($moveCentralDestinationRootFolderName)
    }
    if ($moveCentralPreserveFolderPath) {
        foreach ($part in @($sourcePath)) { $destinationPath.Add($part) }
    }

    $destinationFolder = $null
    $destinationFolderId = $null
    $status = $null
    $errorMessage = $null

    try {
        if (-not $articleId) {
            throw "Article id was not detected."
        }

        if ($destinationPath.Count -gt 0) {
            $destinationFolder = Ensure-HuduMoveCentralFolderPath `
                -Path $destinationPath.ToArray() `
                -CentralFolderById $centralIndex.FolderById `
                -CentralChildrenByParent $centralIndex.ChildrenByParent `
                -DryRun:$moveCentralDryRun
            $destinationFolderId = $destinationFolder.id ?? $destinationFolder.Id
        }

        if ($moveCentralDryRun) {
            $status = "DryRunMove"
            $dryRun++
        } else {
            Set-HuduMoveCentralArticle -ArticleId $articleId -FolderId $destinationFolderId | Out-Null
            $status = "Moved"
            $moved++
            Write-HuduMoveCentralLog -Message "Moved article $articleIndex/$($articles.Count): '$articleName' to central KB folder '$(@($destinationPath) -join '\')'." -Color Green
        }
    } catch {
        $status = "Failed"
        $errorMessage = $_.Exception.Message
        $failed++
        Write-HuduMoveCentralLog -Message "Failed to move article '$articleName' ($articleId): $errorMessage" -Color Red
    }

    if (-not $status) {
        $status = "Skipped"
        $skipped++
    }

    $report.Add([PSCustomObject]@{
        Status              = $status
        ArticleId           = $articleId
        ArticleName         = $articleName
        SourceCompanyId     = $moveCentralSourceCompanyId
        SourceFolderId      = $sourceFolderId
        SourceFolderPath    = (@($sourcePath) -join '\')
        DestinationFolderId = $destinationFolderId
        DestinationPath     = (@($destinationPath) -join '\')
        Error               = $errorMessage
    })
}

$report | Export-Csv -LiteralPath $moveCentralReportPath -NoTypeInformation -Encoding UTF8
Write-HuduMoveCentralLog -Message "Move to central KB complete: $moved moved, $dryRun dry-run move(s), $skipped skipped, $failed failed. Report: $moveCentralReportPath" -Color Cyan
