function ConvertTo-AttributionNormalizedText {
    param ($Value)

    if ($null -eq $Value) { return "" }

    $text = ([string]$Value).ToLowerInvariant()
    $text = [System.Web.HttpUtility]::HtmlDecode($text)
    $text = $text -replace '&', ' and '
    $text = $text -replace '[^a-z0-9]+', ' '
    $text = $text -replace '\s+', ' '
    return $text.Trim()
}

function Remove-AttributionLegalSuffixes {
    param ($Value)

    $text = ConvertTo-AttributionNormalizedText $Value
    if (-not $text) { return "" }

    $suffixes = @(
        'incorporated', 'corporation', 'corp', 'limited', 'ltd',
        'inc', 'llc', 'llp', 'lp', 'plc', 'co', 'company'
    )

    $tokens = [System.Collections.Generic.List[string]]::new()
    foreach ($token in ($text -split '\s+')) {
        if ($suffixes -notcontains $token) {
            $tokens.Add($token)
        }
    }

    return (($tokens | Where-Object { $_ }) -join ' ').Trim()
}

function ConvertFrom-SharePointClientTitle {
    param (
        [Parameter(Mandatory)]
        [string]$Title
    )

    $trimmed = $Title.Trim()
    $provider = $null
    $code = $null
    $name = $trimmed

    if ($name -match '\s*\[(?<provider>[^\]]+)\]\s*$') {
        $provider = $Matches.provider.Trim()
        $name = ($name -replace '\s*\[[^\]]+\]\s*$', '').Trim()
    }

    if ($name -match '\s*\((?<code>[^)]+)\)\s*$') {
        $code = $Matches.code.Trim()
        $name = ($name -replace '\s*\([^)]+\)\s*$', '').Trim()
    }

    [PSCustomObject]@{
        RawTitle       = $trimmed
        ClientName     = $name
        ClientCode     = $code
        Provider       = $provider
        NormalizedName = ConvertTo-AttributionNormalizedText $name
        StrippedName   = Remove-AttributionLegalSuffixes $name
    }
}

function Get-AttributionLevenshteinDistance {
    param (
        [string]$Left,
        [string]$Right
    )

    if ($null -eq $Left) { $Left = "" }
    if ($null -eq $Right) { $Right = "" }

    $n = $Left.Length
    $m = $Right.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }

    $d = New-Object 'int[,]' ($n + 1), ($m + 1)
    for ($i = 0; $i -le $n; $i++) { $d[$i, 0] = $i }
    for ($j = 0; $j -le $m; $j++) { $d[0, $j] = $j }

    for ($i = 1; $i -le $n; $i++) {
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($Left[$i - 1] -eq $Right[$j - 1]) { 0 } else { 1 }
            $delete = $d[($i - 1), $j] + 1
            $insert = $d[$i, ($j - 1)] + 1
            $substitute = $d[($i - 1), ($j - 1)] + $cost
            $d[$i, $j] = [Math]::Min([Math]::Min($delete, $insert), $substitute)
        }
    }

    return $d[$n, $m]
}

function Get-AttributionTokenScore {
    param (
        [string]$Left,
        [string]$Right
    )

    $leftTokens = @(ConvertTo-AttributionNormalizedText $Left -split '\s+' | Where-Object { $_ -and $_.Length -gt 1 } | Sort-Object -Unique)
    $rightTokens = @(ConvertTo-AttributionNormalizedText $Right -split '\s+' | Where-Object { $_ -and $_.Length -gt 1 } | Sort-Object -Unique)

    if ($leftTokens.Count -eq 0 -or $rightTokens.Count -eq 0) { return 0 }

    $leftSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$leftTokens)
    $rightSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$rightTokens)
    $intersection = [System.Collections.Generic.HashSet[string]]::new($leftSet)
    $intersection.IntersectWith($rightSet)
    $union = [System.Collections.Generic.HashSet[string]]::new($leftSet)
    $union.UnionWith($rightSet)

    if ($union.Count -eq 0) { return 0 }
    return [Math]::Round(($intersection.Count / $union.Count) * 100, 2)
}

function Get-AttributionSimilarityScore {
    param (
        [string]$Left,
        [string]$Right
    )

    $leftNormalized = ConvertTo-AttributionNormalizedText $Left
    $rightNormalized = ConvertTo-AttributionNormalizedText $Right

    if (-not $leftNormalized -or -not $rightNormalized) { return 0 }
    if ($leftNormalized -eq $rightNormalized) { return 100 }

    $maxLength = [Math]::Max($leftNormalized.Length, $rightNormalized.Length)
    $distance = Get-AttributionLevenshteinDistance -Left $leftNormalized -Right $rightNormalized
    $levenshteinScore = [Math]::Round((1 - ($distance / $maxLength)) * 100, 2)
    $tokenScore = Get-AttributionTokenScore -Left $leftNormalized -Right $rightNormalized

    return [Math]::Max($levenshteinScore, $tokenScore)
}

function Get-AttributionBestSourceWindowScore {
    param (
        [string]$Alias,
        [string[]]$SourceTokens
    )

    $aliasTokens = @(ConvertTo-AttributionNormalizedText $Alias -split '\s+' | Where-Object { $_ })
    if ($aliasTokens.Count -eq 0 -or $SourceTokens.Count -eq 0) { return 0 }

    $windowSizes = @(
        [Math]::Max(1, $aliasTokens.Count - 1),
        $aliasTokens.Count,
        ($aliasTokens.Count + 1)
    ) | Sort-Object -Unique

    $best = 0
    $aliasText = $aliasTokens -join ' '

    foreach ($windowSize in $windowSizes) {
        if ($windowSize -gt $SourceTokens.Count) { continue }
        for ($i = 0; $i -le ($SourceTokens.Count - $windowSize); $i++) {
            $windowText = @($SourceTokens[$i..($i + $windowSize - 1)]) -join ' '
            $score = Get-AttributionSimilarityScore -Left $aliasText -Right $windowText
            if ($score -gt $best) { $best = $score }
        }
    }

    return [double]$best
}

function Get-SharePointClientListItemSourceMatchCandidates {
    param (
        [Parameter(Mandatory)]
        [string]$SourceText,

        [Parameter(Mandatory)]
        [array]$AttributionMap,

        [switch]$AutoOnly,

        [switch]$AllowUnmatchedClientEntry
    )

    $normalizedSource = ConvertTo-AttributionNormalizedText $SourceText
    if (-not $normalizedSource) { return @() }

    $sourceTokens = @($normalizedSource -split '\s+' | Where-Object { $_ })
    $sourceTokenSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$sourceTokens)

    foreach ($entry in @($AttributionMap)) {
        if ($AutoOnly -and -not $entry.AutoMatched -and -not $AllowUnmatchedClientEntry) { continue }

        $bestScore = 0
        $bestAlias = $null
        $bestReason = $null
        $bestAliasLength = 0

        $normalizedCode = ConvertTo-AttributionNormalizedText ($entry.NormalizedClientCode ?? $entry.ClientCode)
        if ($normalizedCode -and $normalizedCode.Length -ge 2 -and $sourceTokenSet.Contains($normalizedCode)) {
            $bestScore = 100
            $bestAlias = $normalizedCode
            $bestReason = 'exact_client_code_in_source'
            $bestAliasLength = $normalizedCode.Length
        }

        foreach ($alias in @($entry.Aliases)) {
            $normalizedAlias = ConvertTo-AttributionNormalizedText $alias
            if (-not $normalizedAlias -or $normalizedAlias.Length -lt 3) { continue }

            $score = 0
            $reason = $null
            $aliasTokens = @($normalizedAlias -split '\s+' | Where-Object { $_ -and $_.Length -gt 1 })
            $sharedTokenCount = if ($aliasTokens.Count -gt 0) {
                @($aliasTokens | Where-Object { $sourceTokenSet.Contains($_) }).Count
            } else {
                0
            }

            if ($normalizedSource -eq $normalizedAlias -or $normalizedSource.Contains($normalizedAlias)) {
                $score = 99
                $reason = 'client_list_item_alias_in_source'
            } elseif ($aliasTokens.Count -gt 0 -and $sharedTokenCount -eq $aliasTokens.Count) {
                $score = 97
                $reason = 'client_list_item_tokens_in_source'
            } elseif ($sharedTokenCount -ge [Math]::Max(1, [Math]::Ceiling($aliasTokens.Count * 0.5))) {
                $windowScore = Get-AttributionBestSourceWindowScore -Alias $normalizedAlias -SourceTokens $sourceTokens
                if ($windowScore -ge 85) {
                    $score = [double]$windowScore
                    $reason = 'client_list_item_fuzzy_source_window'
                }
            }

            if ($score -gt $bestScore -or ($score -eq $bestScore -and $normalizedAlias.Length -gt $bestAliasLength)) {
                $bestScore = [double]$score
                $bestAlias = $normalizedAlias
                $bestReason = $reason
                $bestAliasLength = $normalizedAlias.Length
            }
        }

        if ($bestScore -gt 0) {
            [PSCustomObject]@{
                Entry                 = $entry
                Alias                 = $bestAlias
                AliasLength           = $bestAliasLength
                Confidence            = [double]$bestScore
                ClientMatchConfidence = [double]$bestScore
                HuduMatchConfidence   = [double]($entry.Confidence ?? 0)
                Reason                = $bestReason
            }
        }
    }
}

function Get-SharePointClientListEntries {
    param (
        [Parameter(Mandatory)] $ManifestSet,
        [string[]]$ListNames = @('Client List')
    )

    $normalizedListNames = @($ListNames | ForEach-Object { ConvertTo-AttributionNormalizedText $_ })

    foreach ($manifest in @($ManifestSet.Manifests)) {
        foreach ($siteEntry in @($manifest.sites)) {
            foreach ($listEntry in @($siteEntry.lists)) {
                $listName = $listEntry.metadata.displayName ?? $listEntry.metadata.name
                if ($normalizedListNames -notcontains (ConvertTo-AttributionNormalizedText $listName)) {
                    continue
                }

                foreach ($item in @($listEntry.items)) {
                    $title = $item.fields.Title ?? $item.fields.LinkTitle ?? $item.webUrl ?? $item.id
                    if ([string]::IsNullOrWhiteSpace([string]$title)) { continue }

                    $parsed = ConvertFrom-SharePointClientTitle -Title $title
                    [PSCustomObject]@{
                        SharePointItemId = $item.id
                        ListName         = $listName
                        SiteName         = $siteEntry.metadata.displayName
                        SiteId           = $siteEntry.metadata.id
                        WebUrl           = $item.webUrl
                        ClientActive     = $item.fields.ClientActive
                        WhiteLabelled    = $item.fields.WhiteLabelled
                        RawTitle         = $parsed.RawTitle
                        ClientName       = $parsed.ClientName
                        ClientCode       = $parsed.ClientCode
                        Provider         = $parsed.Provider
                        NormalizedName   = $parsed.NormalizedName
                        StrippedName     = $parsed.StrippedName
                        AttributionSource = 'sharepoint_list'
                    }
                }
            }
        }
    }
}

function Import-SharePointClientAttributionClientFile {
    param (
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return @()
    }

    $raw = Get-Content -LiteralPath $Path -Raw
    $values = @()

    try {
        $json = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($json.PSObject.Properties.Name -contains 'clients') {
            $values = @($json.clients)
        } elseif ($json.PSObject.Properties.Name -contains 'Clients') {
            $values = @($json.Clients)
        } elseif ($json.PSObject.Properties.Name -contains 'items') {
            $values = @($json.items)
        } elseif ($json.PSObject.Properties.Name -contains 'Items') {
            $values = @($json.Items)
        } elseif ($json -is [array]) {
            $values = @($json)
        } else {
            $values = @($json)
        }
    } catch {
        $values = @(
            [regex]::Matches($raw, '"(?<value>(?:\\"|[^"])*)"') |
                ForEach-Object {
                    $_.Groups['value'].Value -replace '\\"', '"'
                }
        )
    }

    foreach ($value in $values) {
        $title = if ($null -ne $value -and $value.PSObject.Properties.Name -contains 'title') {
            [string]$value.title
        } elseif ($null -ne $value -and $value.PSObject.Properties.Name -contains 'Title') {
            [string]$value.Title
        } elseif ($null -ne $value -and $value.PSObject.Properties.Name -contains 'name') {
            [string]$value.name
        } elseif ($null -ne $value -and $value.PSObject.Properties.Name -contains 'Name') {
            [string]$value.Name
        } elseif ($null -ne $value -and $value.PSObject.Properties.Name -contains 'clientName') {
            [string]$value.clientName
        } elseif ($null -ne $value -and $value.PSObject.Properties.Name -contains 'ClientName') {
            [string]$value.ClientName
        } else {
            [string]$value
        }

        if ([string]::IsNullOrWhiteSpace($title)) { continue }
        $parsed = ConvertFrom-SharePointClientTitle -Title $title
        [PSCustomObject]@{
            SharePointItemId = $null
            ListName         = 'clients.json'
            SiteName         = $null
            SiteId           = $null
            WebUrl           = $null
            ClientActive     = $null
            WhiteLabelled    = $null
            RawTitle         = $parsed.RawTitle
            ClientName       = $parsed.ClientName
            ClientCode       = $parsed.ClientCode
            Provider         = $parsed.Provider
            NormalizedName   = $parsed.NormalizedName
            StrippedName     = $parsed.StrippedName
            AttributionSource = 'client_file'
        }
    }
}

function Get-HuduCompanyAttributionCandidates {
    param (
        [Parameter(Mandatory)] $ClientEntry,
        [Parameter(Mandatory)] [array]$Companies
    )

    foreach ($company in @($Companies)) {
        $companyName = [string]$company.Name
        $companyNormalized = ConvertTo-AttributionNormalizedText $companyName
        $companyStripped = Remove-AttributionLegalSuffixes $companyName
        $score = 0
        $reason = 'fuzzy'

        if ($ClientEntry.NormalizedName -and $ClientEntry.NormalizedName -eq $companyNormalized) {
            $score = 100
            $reason = 'exact_normalized_name'
        }
        elseif ($ClientEntry.ClientCode -and $ClientEntry.ClientCode.Length -ge 3 -and $companyNormalized -match "(^| )$([regex]::Escape((ConvertTo-AttributionNormalizedText $ClientEntry.ClientCode)))( |$)") {
            $score = 98
            $reason = 'client_code_in_company_name'
        }
        elseif ($ClientEntry.StrippedName -and $ClientEntry.StrippedName -eq $companyStripped) {
            $score = 96
            $reason = 'exact_legal_suffix_stripped'
        }
        else {
            $score = [Math]::Max(
                (Get-AttributionSimilarityScore -Left $ClientEntry.ClientName -Right $companyName),
                (Get-AttributionSimilarityScore -Left $ClientEntry.StrippedName -Right $companyStripped)
            )
        }

        [PSCustomObject]@{
            CompanyId   = $company.Id
            CompanyName = $companyName
            Score       = [double]$score
            Reason      = $reason
        }
    }
}

function New-HuduClientAttributionMapFromEntries {
    param (
        [Parameter(Mandatory)] [array]$Entries,
        [Parameter(Mandatory)] [array]$Companies,
        [int]$MinScore = 95,
        [int]$MinGap = 5
    )

    foreach ($entry in @($Entries)) {
        $candidates = @(Get-HuduCompanyAttributionCandidates -ClientEntry $entry -Companies $Companies | Sort-Object Score -Descending)
        $best = $candidates | Select-Object -First 1
        $second = $candidates | Select-Object -Skip 1 -First 1
        $gap = if ($second) { [double]$best.Score - [double]$second.Score } elseif ($best) { [double]$best.Score } else { 0 }
        $autoMatched = ($best -and [double]$best.Score -ge $MinScore -and $gap -ge $MinGap)
        $aliases = [System.Collections.Generic.List[string]]::new()

        foreach ($alias in @($entry.ClientName, $entry.StrippedName)) {
            $normalizedAlias = ConvertTo-AttributionNormalizedText $alias
            if ($normalizedAlias -and $normalizedAlias.Length -ge 3 -and -not $aliases.Contains($normalizedAlias)) {
                $aliases.Add($normalizedAlias)
            }
        }

        $normalizedCode = ConvertTo-AttributionNormalizedText $entry.ClientCode

        [PSCustomObject]@{
            SharePointItemId      = $entry.SharePointItemId
            ListName              = $entry.ListName
            SiteName              = $entry.SiteName
            SiteId                = $entry.SiteId
            WebUrl                = $entry.WebUrl
            AttributionSource     = $entry.AttributionSource
            RawTitle              = $entry.RawTitle
            ClientName            = $entry.ClientName
            ClientCode            = $entry.ClientCode
            NormalizedClientCode = $normalizedCode
            Provider              = $entry.Provider
            ClientActive          = $entry.ClientActive
            HuduCompanyId         = $best.CompanyId
            HuduCompanyName       = $best.CompanyName
            Confidence            = if ($best) { [double]$best.Score } else { 0 }
            ConfidenceGap         = [double]$gap
            MatchReason           = $best.Reason
            AutoMatched           = [bool]$autoMatched
            MatchStatus           = if ($autoMatched) { 'Auto' } elseif ($best) { 'Review' } else { 'NoMatch' }
            Aliases               = @($aliases)
            TopCandidates         = @($candidates | Select-Object -First 5)
        }
    }
}

function New-SharePointClientAttributionMap {
    param (
        [Parameter(Mandatory)] $ManifestSet,
        [Parameter(Mandatory)] [array]$Companies,
        [string[]]$ListNames = @('Client List'),
        [int]$MinScore = 95,
        [int]$MinGap = 5
    )

    $entries = @(Get-SharePointClientListEntries -ManifestSet $ManifestSet -ListNames $ListNames)
    New-HuduClientAttributionMapFromEntries -Entries $entries -Companies $Companies -MinScore $MinScore -MinGap $MinGap
}

function Resolve-HuduCompanyFromSharePointAttributionMap {
    param (
        [Parameter(Mandatory)]
        [string]$SourceText,

        [Parameter(Mandatory)]
        [array]$AttributionMap,

        [switch]$AutoOnly,

        [switch]$AllowUnmatchedClientEntry,

        [int]$MinScore = 95,

        [int]$MinGap = 3
    )

    $matches = @(
        Get-SharePointClientListItemSourceMatchCandidates `
            -SourceText $SourceText `
            -AttributionMap $AttributionMap `
            -AutoOnly:$AutoOnly `
            -AllowUnmatchedClientEntry:$AllowUnmatchedClientEntry
    ) | Sort-Object Confidence, AliasLength, HuduMatchConfidence -Descending

    $best = $matches | Select-Object -First 1
    $second = $matches | Select-Object -Skip 1 -First 1
    if (-not $best -or [double]$best.Confidence -lt $MinScore) { return $null }

    $gap = if ($second) { [double]$best.Confidence - [double]$second.Confidence } else { [double]$best.Confidence }
    if ($gap -lt $MinGap) { return $null }

    $best | Add-Member -MemberType NoteProperty -Name ConfidenceGap -Value ([double]$gap) -Force
    return $best
}

function Confirm-HuduCompanyForSharePointAttributionMatch {
    param (
        $AttributionMatch,
        [array]$AttributionMap = @(),
        [switch]$CreateMissing
    )

    if (-not $AttributionMatch -or -not $AttributionMatch.Entry) { return $null }

    $entry = $AttributionMatch.Entry
    if ($entry.HuduCompanyId -and ($entry.AutoMatched -or $entry.MatchStatus -eq 'Created')) { return $entry }
    if (-not $CreateMissing) { return $null }

    $companyName = $entry.ClientName
    if ([string]::IsNullOrWhiteSpace([string]$companyName)) {
        $companyName = $entry.RawTitle
    }

    if ([string]::IsNullOrWhiteSpace([string]$companyName)) { return $null }

    $created = New-HuduCompany -Name $companyName
    $company = $created.company ?? $created
    if (-not $company -or -not $company.Id) {
        throw "New-HuduCompany did not return a company id for '$companyName'."
    }

    $normalizedName = ConvertTo-AttributionNormalizedText $entry.ClientName
    foreach ($relatedEntry in @($AttributionMap)) {
        if ((ConvertTo-AttributionNormalizedText $relatedEntry.ClientName) -eq $normalizedName) {
            $relatedEntry.HuduCompanyId = $company.Id
            $relatedEntry.HuduCompanyName = $company.Name
            $relatedEntry.Confidence = 100
            $relatedEntry.ConfidenceGap = 100
            $relatedEntry.MatchReason = 'created_missing_company_from_client_list_item'
            $relatedEntry.AutoMatched = $true
            $relatedEntry.MatchStatus = 'Created'
        }
    }

    return $entry
}

function Get-HuduCompanySiteCandidates {
    param (
        [Parameter(Mandatory)] $Site,
        [Parameter(Mandatory)] [array]$Companies
    )

    $siteName = $Site.displayName ?? $Site.name
    $siteSlug = $Site.name
    $siteEntry = [PSCustomObject]@{
        ClientName     = $siteName
        ClientCode     = $null
        NormalizedName = ConvertTo-AttributionNormalizedText $siteName
        StrippedName   = Remove-AttributionLegalSuffixes $siteName
    }

    $candidates = @(Get-HuduCompanyAttributionCandidates -ClientEntry $siteEntry -Companies $Companies)

    if ($siteSlug -and $siteSlug -ne $siteName) {
        $slugEntry = [PSCustomObject]@{
            ClientName     = $siteSlug
            ClientCode     = $null
            NormalizedName = ConvertTo-AttributionNormalizedText $siteSlug
            StrippedName   = Remove-AttributionLegalSuffixes $siteSlug
        }

        $slugCandidates = @(Get-HuduCompanyAttributionCandidates -ClientEntry $slugEntry -Companies $Companies)
        foreach ($candidate in $slugCandidates) {
            $existing = $candidates | Where-Object { $_.CompanyId -eq $candidate.CompanyId } | Select-Object -First 1
            if ($existing) {
                if ([double]$candidate.Score -gt [double]$existing.Score) {
                    $existing.Score = [double]$candidate.Score
                    $existing.Reason = "site_slug_$($candidate.Reason)"
                }
            } else {
                $candidate.Reason = "site_slug_$($candidate.Reason)"
                $candidates += $candidate
            }
        }
    }

    return $candidates | Sort-Object Score -Descending
}

function Resolve-HuduCompanyFromSiteCompanyMap {
    param (
        [string]$SiteId,
        [string]$SiteName,
        [array]$SiteCompanyMap
    )

    if ($SiteId) {
        $match = $SiteCompanyMap | Where-Object { $_.SiteId -eq $SiteId } | Select-Object -First 1
        if ($match) { return $match }
    }

    if ($SiteName) {
        $normalizedSiteName = ConvertTo-AttributionNormalizedText $SiteName
        $match = $SiteCompanyMap |
            Where-Object { (ConvertTo-AttributionNormalizedText $_.SiteName) -eq $normalizedSiteName } |
            Select-Object -First 1
        if ($match) { return $match }
    }

    return $null
}
