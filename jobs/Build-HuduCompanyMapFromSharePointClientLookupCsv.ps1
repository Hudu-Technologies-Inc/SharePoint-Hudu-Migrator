<#
.SYNOPSIS
Build a SharePoint lookup-id to Hudu company attribution map from a client lookup CSV.

.DESCRIPTION
Consumes the lookup candidate CSV produced by Dump-SharePointFileClientMetadata.ps1 and
matches CandidateTitle values against Hudu companies. Writes a review CSV for all rows and
a JSON map containing only confident Auto matches by default.
#>

param(
    [string]$InputPath = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'logs') 'sharepoint-file-client-lookup-candidates.csv'),
    [string]$OutputPath = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'logs') 'client-attribution-map-from-lookup-list.json'),
    [string]$ReviewPath = (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'logs') 'client-attribution-map-from-lookup-list-review.csv'),

    [string]$CompaniesJsonPath = "",
    [string[]]$ListNames = @('Client List'),

    [ValidateRange(0, 100)]
    [int]$MinScore = 95,

    [ValidateRange(0, 100)]
    [int]$MinGap = 5,

    [switch]$IncludeReviewMatchesInJson
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'helpers\attribution.ps1')

function Resolve-CompanyLookupPath {
    param([string]$Path)

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $Path))
}

function Select-CompanyLookupFirstValue {
    param(
        [Parameter(ValueFromRemainingArguments)]
        [object[]]$Values
    )

    foreach ($value in @($Values)) {
        if ($null -eq $value) { continue }
        $text = [string]$value
        if (-not [string]::IsNullOrWhiteSpace($text)) { return $text }
    }

    return $null
}

function ConvertTo-CompanyLookupBool {
    param($Value)

    if ($Value -is [bool]) { return [bool]$Value }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    if ($text -match '^(?i:true|yes|1)$') { return $true }
    if ($text -match '^(?i:false|no|0)$') { return $false }
    return $null
}

function Import-CompanyLookupCompanies {
    param([string]$Path)

    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        $resolvedPath = Resolve-CompanyLookupPath -Path $Path
        if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
            throw "Companies JSON not found: $resolvedPath"
        }

        return @(Get-Content -LiteralPath $resolvedPath -Raw | ConvertFrom-Json)
    }

    if (-not (Get-Command Get-HuduCompanies -ErrorAction SilentlyContinue)) {
        throw "Get-HuduCompanies is not available. Load the Hudu API module/auth first, or pass -CompaniesJsonPath."
    }

    return @(Get-HuduCompanies)
}

$resolvedInputPath = Resolve-CompanyLookupPath -Path $InputPath
$resolvedOutputPath = Resolve-CompanyLookupPath -Path $OutputPath
$resolvedReviewPath = Resolve-CompanyLookupPath -Path $ReviewPath

if (-not (Test-Path -LiteralPath $resolvedInputPath -PathType Leaf)) {
    throw "Input CSV not found: $resolvedInputPath"
}

foreach ($path in @($resolvedOutputPath, $resolvedReviewPath)) {
    $directory = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) {
        $null = New-Item -ItemType Directory -Path $directory -Force
    }
}

$listNameKeys = @($ListNames | ForEach-Object { ConvertTo-AttributionNormalizedText $_ })
$csvRows = @(Import-Csv -LiteralPath $resolvedInputPath)
$rows = @(
    foreach ($row in $csvRows) {
        $lookupId = Select-CompanyLookupFirstValue $row.LookupId $row.SharePointItemId $row.Id
        $candidateTitle = Select-CompanyLookupFirstValue $row.CandidateTitle $row.RawTitle $row.Title $row.ClientName
        $listName = Select-CompanyLookupFirstValue $row.ListName

        if ([string]::IsNullOrWhiteSpace($lookupId) -or [string]::IsNullOrWhiteSpace($candidateTitle)) { continue }
        if ($listNameKeys.Count -gt 0 -and $listNameKeys -notcontains (ConvertTo-AttributionNormalizedText $listName)) { continue }

        $clientActive = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$row.CandidateFieldsJson)) {
            try {
                $fields = $row.CandidateFieldsJson | ConvertFrom-Json
                $clientActive = ConvertTo-CompanyLookupBool (Select-CompanyLookupFirstValue $fields.ClientActive)
            } catch {}
        }

        $parsed = ConvertFrom-SharePointClientTitle -Title $candidateTitle
        [PSCustomObject]@{
            SharePointItemId  = [string]$lookupId
            ListName          = $listName
            SiteName          = $null
            SiteId            = $null
            WebUrl            = $row.CandidateUrl
            ClientActive      = $clientActive
            RawTitle          = $parsed.RawTitle
            ClientName        = $parsed.ClientName
            ClientCode        = $parsed.ClientCode
            Provider          = $parsed.Provider
            NormalizedName    = $parsed.NormalizedName
            StrippedName      = $parsed.StrippedName
            AttributionSource = 'sharepoint_lookup_candidate_csv'
        }
    }
)

$companies = @(Import-CompanyLookupCompanies -Path $CompaniesJsonPath)
$map = @(New-HuduClientAttributionMapFromEntries -Entries $rows -Companies $companies -MinScore $MinScore -MinGap $MinGap)

$map |
    Select-Object `
        MatchStatus,
        AutoMatched,
        Confidence,
        ConfidenceGap,
        SharePointItemId,
        ListName,
        RawTitle,
        ClientName,
        ClientCode,
        Provider,
        ClientActive,
        HuduCompanyId,
        HuduCompanyName,
        MatchReason,
        @{ Name = 'Candidate2'; Expression = { @($_.TopCandidates | Select-Object -Skip 1 -First 1).CompanyName } },
        @{ Name = 'Candidate2Score'; Expression = { @($_.TopCandidates | Select-Object -Skip 1 -First 1).Score } },
        @{ Name = 'Candidate3'; Expression = { @($_.TopCandidates | Select-Object -Skip 2 -First 1).CompanyName } },
        @{ Name = 'Candidate3Score'; Expression = { @($_.TopCandidates | Select-Object -Skip 2 -First 1).Score } },
        WebUrl |
    Export-Csv -LiteralPath $resolvedReviewPath -NoTypeInformation -Encoding UTF8

$jsonRows = if ($IncludeReviewMatchesInJson) {
    @($map | Where-Object { $_.HuduCompanyId })
} else {
    @($map | Where-Object { $_.AutoMatched -and $_.HuduCompanyId })
}

$jsonRows |
    ConvertTo-Json -Depth 20 |
    Set-Content -LiteralPath $resolvedOutputPath -Encoding UTF8

[PSCustomObject]@{
    InputRows       = $csvRows.Count
    CandidateRows   = $rows.Count
    Companies       = $companies.Count
    AutoMatches     = @($map | Where-Object { $_.AutoMatched -and $_.HuduCompanyId }).Count
    ReviewMatches   = @($map | Where-Object { -not $_.AutoMatched -and $_.HuduCompanyId }).Count
    NoMatches       = @($map | Where-Object { -not $_.HuduCompanyId }).Count
    JsonRows        = $jsonRows.Count
    OutputPath      = $resolvedOutputPath
    ReviewPath      = $resolvedReviewPath
}
