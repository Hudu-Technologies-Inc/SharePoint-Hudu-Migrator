function ConvertTo-AttributionNormalizedText {
    param ($Value)

    if ($null -eq $Value) { return "" }

    $text = ([string]$Value).Normalize([Text.NormalizationForm]::FormD).ToLowerInvariant()
    $text = $text -replace '\p{Mn}', ''
    $text = [System.Web.HttpUtility]::HtmlDecode($text)
    $text = $text -replace '&', ' and '
    $text = $text -replace '[^a-z0-9]+', ' '
    $text = $text -replace '\s+', ' '
    return $text.Trim()
}

function ConvertTo-AttributionCompactKey {
    param ($Value)

    if ($null -eq $Value) { return "" }

    $text = ([string]$Value).Normalize([Text.NormalizationForm]::FormD).ToLowerInvariant()
    $text = $text -replace '\p{Mn}', ''
    $text = [System.Web.HttpUtility]::HtmlDecode($text)
    $text = $text -replace '&', 'and'
    $text = $text -replace '[^a-z0-9]+', ''
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

function Get-AttributionSignificantTokens {
    param ($Value)

    $ignoredTokens = [System.Collections.Generic.HashSet[string]]::new([string[]]@(
        'the', 'and', 'for', 'with', 'from',
        'incorporated', 'corporation', 'corp', 'limited', 'ltd',
        'inc', 'llc', 'llp', 'lp', 'plc', 'co', 'company'
    ))

    @(
        ConvertTo-AttributionNormalizedText $Value -split '\s+' |
            Where-Object { $_ -and $_.Length -gt 2 -and -not $ignoredTokens.Contains($_) } |
            Sort-Object -Unique
    )
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

    $leftCompact = ConvertTo-AttributionCompactKey $Left
    $rightCompact = ConvertTo-AttributionCompactKey $Right
    if ($leftCompact -and $rightCompact -and $leftCompact -eq $rightCompact) { return 100 }

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

function Add-SharePointAttributionIndexValue {
    param (
        [Parameter(Mandatory)] [hashtable]$Table,
        [Parameter(Mandatory)] [string]$Key,
        [Parameter(Mandatory)] $Value
    )

    if (-not $Table.ContainsKey($Key)) {
        $Table[$Key] = [System.Collections.Generic.List[object]]::new()
    }

    $Table[$Key].Add($Value)
}

function Test-SharePointAttributionEntryEligible {
    param (
        $Entry,
        [switch]$AutoOnly,
        [switch]$AllowUnmatchedClientEntry
    )

    if ($AutoOnly -and -not $Entry.AutoMatched -and -not $AllowUnmatchedClientEntry) {
        return $false
    }

    return $true
}

function New-SharePointClientAttributionLookup {
    param (
        [AllowEmptyCollection()]
        [array]$AttributionMap
    )

    $lookup = [PSCustomObject]@{
        IsSharePointClientAttributionLookup = $true
        Entries                             = @($AttributionMap)
        CodeToItems                         = @{}
        TokenToAliasItems                   = @{}
        CompactAliasPrefixToItems           = @{}
        AliasItems                          = [System.Collections.Generic.List[object]]::new()
        SourceMatchCache                    = @{}
        SourceMatchCacheMaxItems            = 50000
    }

    $entryIndex = 0
    foreach ($entry in @($AttributionMap)) {
        $entryKey = [string]$entryIndex
        $entryIndex++

        $normalizedCode = ConvertTo-AttributionNormalizedText ($entry.NormalizedClientCode ?? $entry.ClientCode)
        if ($normalizedCode -and $normalizedCode.Length -ge 2) {
            Add-SharePointAttributionIndexValue `
                -Table $lookup.CodeToItems `
                -Key $normalizedCode `
                -Value ([PSCustomObject]@{
                    Entry    = $entry
                    EntryKey = $entryKey
                    Code     = $normalizedCode
                })
        }

        foreach ($alias in @($entry.Aliases)) {
            $normalizedAlias = ConvertTo-AttributionNormalizedText $alias
            if (-not $normalizedAlias -or $normalizedAlias.Length -lt 3) { continue }

            $compactAlias = ConvertTo-AttributionCompactKey $alias
            $aliasTokens = @(
                $normalizedAlias -split '\s+' |
                    Where-Object { $_ -and $_.Length -gt 1 } |
                    Sort-Object -Unique
            )

            if ($aliasTokens.Count -eq 0) { continue }

            $aliasItem = [PSCustomObject]@{
                Entry           = $entry
                EntryKey        = $entryKey
                Alias           = $normalizedAlias
                CompactAlias    = $compactAlias
                AliasLength     = $normalizedAlias.Length
                Tokens          = @($aliasTokens)
                RequiredMatches = [Math]::Max(1, [Math]::Ceiling($aliasTokens.Count * 0.5))
            }

            $lookup.AliasItems.Add($aliasItem)
            foreach ($token in $aliasTokens) {
                Add-SharePointAttributionIndexValue -Table $lookup.TokenToAliasItems -Key $token -Value $aliasItem
            }

            if ($compactAlias -and $compactAlias.Length -ge 5) {
                $prefixLength = [Math]::Min(4, $compactAlias.Length)
                $compactPrefix = $compactAlias.Substring(0, $prefixLength)
                Add-SharePointAttributionIndexValue -Table $lookup.CompactAliasPrefixToItems -Key $compactPrefix -Value $aliasItem
            }
        }
    }

    return $lookup
}

function Resolve-SharePointClientAttributionLookup {
    param (
        [Parameter(Mandatory)]
        $AttributionMap
    )

    if (
        $AttributionMap.PSObject.Properties.Name -contains 'IsSharePointClientAttributionLookup' -and
        $AttributionMap.IsSharePointClientAttributionLookup
    ) {
        return $AttributionMap
    }

    return New-SharePointClientAttributionLookup -AttributionMap @($AttributionMap)
}

function Set-SharePointAttributionBestMatch {
    param (
        [Parameter(Mandatory)] [hashtable]$BestByEntryKey,
        [Parameter(Mandatory)] $Entry,
        [Parameter(Mandatory)] [string]$EntryKey,
        [Parameter(Mandatory)] [string]$Alias,
        [Parameter(Mandatory)] [double]$Score,
        [Parameter(Mandatory)] [string]$Reason,
        [int]$AliasLength = 0
    )

    $existing = $BestByEntryKey[$EntryKey]
    if (
        -not $existing -or
        $Score -gt [double]$existing.Confidence -or
        ($Score -eq [double]$existing.Confidence -and $AliasLength -gt [int]$existing.AliasLength)
    ) {
        $BestByEntryKey[$EntryKey] = [PSCustomObject]@{
            Entry                 = $Entry
            Alias                 = $Alias
            AliasLength           = $AliasLength
            Confidence            = [double]$Score
            ClientMatchConfidence = [double]$Score
            HuduMatchConfidence   = [double]($Entry.Confidence ?? 0)
            Reason                = $Reason
        }
    }
}

function Get-SharePointClientListItemSourceMatchCandidates {
    param (
        [Parameter(Mandatory)]
        [string]$SourceText,

        [Parameter(Mandatory)]
        $AttributionMap,

        [switch]$AutoOnly,

        [switch]$AllowUnmatchedClientEntry
    )

    $normalizedSource = ConvertTo-AttributionNormalizedText $SourceText
    if (-not $normalizedSource) { return @() }

    $compactSource = ConvertTo-AttributionCompactKey $SourceText
    $sourceTokens = @($normalizedSource -split '\s+' | Where-Object { $_ })
    $sourceTokenSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$sourceTokens)
    $lookup = Resolve-SharePointClientAttributionLookup -AttributionMap $AttributionMap
    $cacheKey = "$normalizedSource`n$compactSource`n$([bool]$AutoOnly)`n$([bool]$AllowUnmatchedClientEntry)"

    if ($lookup.SourceMatchCache.ContainsKey($cacheKey)) {
        return @($lookup.SourceMatchCache[$cacheKey])
    }

    $bestByEntryKey = @{}

    foreach ($sourceToken in $sourceTokens) {
        if (-not $lookup.CodeToItems.ContainsKey($sourceToken)) { continue }

        foreach ($codeItem in @($lookup.CodeToItems[$sourceToken])) {
            if (-not (Test-SharePointAttributionEntryEligible -Entry $codeItem.Entry -AutoOnly:$AutoOnly -AllowUnmatchedClientEntry:$AllowUnmatchedClientEntry)) {
                continue
            }

            Set-SharePointAttributionBestMatch `
                -BestByEntryKey $bestByEntryKey `
                -Entry $codeItem.Entry `
                -EntryKey $codeItem.EntryKey `
                -Alias $codeItem.Code `
                -AliasLength $codeItem.Code.Length `
                -Score 100 `
                -Reason 'exact_client_code_in_source'
        }
    }

    $candidateAliasItems = [System.Collections.Generic.List[object]]::new()
    $seenAliasKeys = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($sourceToken in $sourceTokens) {
        if (-not $lookup.TokenToAliasItems.ContainsKey($sourceToken)) { continue }

        foreach ($aliasItem in @($lookup.TokenToAliasItems[$sourceToken])) {
            $aliasKey = "$($aliasItem.EntryKey)|$($aliasItem.Alias)"
            if ($seenAliasKeys.Add($aliasKey)) {
                $candidateAliasItems.Add($aliasItem)
            }
        }
    }

    if ($compactSource -and $compactSource.Length -ge 5) {
        foreach ($startIndex in 0..($compactSource.Length - 1)) {
            if (($compactSource.Length - $startIndex) -lt 4) { break }
            $prefix = $compactSource.Substring($startIndex, 4)
            if (-not $lookup.CompactAliasPrefixToItems.ContainsKey($prefix)) { continue }

            foreach ($aliasItem in @($lookup.CompactAliasPrefixToItems[$prefix])) {
                if (-not $aliasItem.CompactAlias -or $aliasItem.CompactAlias.Length -lt 5) { continue }
                if (-not $compactSource.Contains($aliasItem.CompactAlias)) { continue }

                $aliasKey = "$($aliasItem.EntryKey)|$($aliasItem.Alias)"
                if ($seenAliasKeys.Add($aliasKey)) {
                    $candidateAliasItems.Add($aliasItem)
                }
            }
        }
    }

    foreach ($aliasItem in @($candidateAliasItems)) {
        if (-not (Test-SharePointAttributionEntryEligible -Entry $aliasItem.Entry -AutoOnly:$AutoOnly -AllowUnmatchedClientEntry:$AllowUnmatchedClientEntry)) {
            continue
        }

        $sharedTokenCount = @($aliasItem.Tokens | Where-Object { $sourceTokenSet.Contains($_) }).Count
        $score = 0
        $reason = $null

        if ($normalizedSource -eq $aliasItem.Alias -or $normalizedSource.Contains($aliasItem.Alias)) {
            $score = 99
            $reason = 'client_list_item_alias_in_source'
        } elseif ($compactSource -and $aliasItem.CompactAlias -and $compactSource.Contains($aliasItem.CompactAlias)) {
            $score = 99
            $reason = 'client_list_item_compact_alias_in_source'
        } elseif ($sharedTokenCount -eq $aliasItem.Tokens.Count) {
            $score = 97
            $reason = 'client_list_item_tokens_in_source'
        } elseif ($sharedTokenCount -ge $aliasItem.RequiredMatches) {
            $windowScore = Get-AttributionBestSourceWindowScore -Alias $aliasItem.Alias -SourceTokens $sourceTokens
            if ($windowScore -ge 85) {
                $score = [double]$windowScore
                $reason = 'client_list_item_fuzzy_source_window'
            }
        }

        if ($score -gt 0) {
            Set-SharePointAttributionBestMatch `
                -BestByEntryKey $bestByEntryKey `
                -Entry $aliasItem.Entry `
                -EntryKey $aliasItem.EntryKey `
                -Alias $aliasItem.Alias `
                -AliasLength $aliasItem.AliasLength `
                -Score $score `
                -Reason $reason
        }
    }

    $matches = @($bestByEntryKey.Values)
    if ($lookup.SourceMatchCache.Count -lt $lookup.SourceMatchCacheMaxItems) {
        $lookup.SourceMatchCache[$cacheKey] = $matches
    }

    return $matches
}

function Get-SharePointClientListEntries {
    param (
        [Parameter(Mandatory)] $ManifestSet,
        [string[]]$ListNames = @('Client List'),
        [string[]]$TitleFieldNames = @('Title', 'LinkTitle', 'Client Name', 'Client', 'Company', 'Name')
    )

    $normalizedListNames = @($ListNames | ForEach-Object { ConvertTo-AttributionNormalizedText $_ })
    $normalizedTitleFieldNames = @(
        $TitleFieldNames |
            ForEach-Object { ConvertTo-AttributionNormalizedText $_ } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )

    foreach ($manifest in @($ManifestSet.Manifests)) {
        foreach ($siteEntry in @($manifest.sites)) {
            foreach ($listEntry in @($siteEntry.lists)) {
                $listName = $listEntry.metadata.displayName ?? $listEntry.metadata.name
                if ($normalizedListNames -notcontains (ConvertTo-AttributionNormalizedText $listName)) {
                    continue
                }

                foreach ($item in @($listEntry.items)) {
                    $title = $null
                    if ($item.fields) {
                        foreach ($property in @($item.fields.PSObject.Properties)) {
                            if ($property.Name -like '@odata*') { continue }
                            if ($null -eq $property.Value) { continue }

                            $normalizedPropertyName = ConvertTo-AttributionNormalizedText $property.Name
                            $decodedPropertyName = if (Get-Command ConvertFrom-SharePointInternalFieldName -ErrorAction SilentlyContinue) {
                                ConvertFrom-SharePointInternalFieldName $property.Name
                            } else {
                                [regex]::Replace($property.Name, '_x(?<hex>[0-9a-fA-F]{4})_', {
                                    param($Match)
                                    [string][char][Convert]::ToInt32($Match.Groups['hex'].Value, 16)
                                })
                            }
                            $normalizedDecodedPropertyName = ConvertTo-AttributionNormalizedText $decodedPropertyName

                            if (
                                $normalizedTitleFieldNames -contains $normalizedPropertyName -or
                                $normalizedTitleFieldNames -contains $normalizedDecodedPropertyName
                            ) {
                                $title = $property.Value
                                break
                            }
                        }
                    }

                    $title = $title ?? $item.fields.Title ?? $item.fields.LinkTitle ?? $item.webUrl ?? $item.id
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
        $valuePropertyNames = if ($null -ne $value -and $value.PSObject.Properties) {
            @($value.PSObject.Properties.Name)
        } else {
            @()
        }

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

        $importedAliases = [System.Collections.Generic.List[string]]::new()
        foreach ($aliasPropertyName in @('aliases', 'Aliases', 'alias', 'Alias')) {
            if ($valuePropertyNames -notcontains $aliasPropertyName) { continue }

            foreach ($alias in @($value.PSObject.Properties[$aliasPropertyName].Value)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$alias)) {
                    $importedAliases.Add([string]$alias)
                }
            }
        }

        $huduCompanyId = if ($valuePropertyNames -contains 'huduCompanyId') {
            $value.huduCompanyId
        } elseif ($valuePropertyNames -contains 'HuduCompanyId') {
            $value.HuduCompanyId
        } elseif ($valuePropertyNames -contains 'companyId') {
            $value.companyId
        } elseif ($valuePropertyNames -contains 'CompanyId') {
            $value.CompanyId
        } else {
            $null
        }

        $huduCompanyName = if ($valuePropertyNames -contains 'huduCompanyName') {
            [string]$value.huduCompanyName
        } elseif ($valuePropertyNames -contains 'HuduCompanyName') {
            [string]$value.HuduCompanyName
        } elseif ($valuePropertyNames -contains 'companyName') {
            [string]$value.companyName
        } elseif ($valuePropertyNames -contains 'CompanyName') {
            [string]$value.CompanyName
        } else {
            $null
        }

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
            ImportedAliases  = @($importedAliases)
            HuduCompanyId    = $huduCompanyId
            HuduCompanyName  = $huduCompanyName
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
        $clientCompact = ConvertTo-AttributionCompactKey $ClientEntry.ClientName
        $clientStrippedCompact = ConvertTo-AttributionCompactKey $ClientEntry.StrippedName
        $companyCompact = ConvertTo-AttributionCompactKey $companyName
        $companyStrippedCompact = ConvertTo-AttributionCompactKey $companyStripped
        $score = 0
        $reason = 'fuzzy'

        if ($ClientEntry.NormalizedName -and $ClientEntry.NormalizedName -eq $companyNormalized) {
            $score = 100
            $reason = 'exact_normalized_name'
        }
        elseif ($clientCompact -and $clientCompact -eq $companyCompact) {
            $score = 100
            $reason = 'exact_compact_name'
        }
        elseif ($ClientEntry.ClientCode -and $ClientEntry.ClientCode.Length -ge 3 -and $companyNormalized -match "(^| )$([regex]::Escape((ConvertTo-AttributionNormalizedText $ClientEntry.ClientCode)))( |$)") {
            $score = 98
            $reason = 'client_code_in_company_name'
        }
        elseif ($ClientEntry.StrippedName -and $ClientEntry.StrippedName -eq $companyStripped) {
            $score = 96
            $reason = 'exact_legal_suffix_stripped'
        }
        elseif ($clientStrippedCompact -and $clientStrippedCompact -eq $companyStrippedCompact) {
            $score = 96
            $reason = 'exact_compact_legal_suffix_stripped'
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
        [AllowEmptyCollection()] [array]$Entries,
        [Parameter(Mandatory)] [array]$Companies,
        [int]$MinScore = 95,
        [int]$MinGap = 5
    )

    if ($null -eq $Entries -or $Entries.Count -lt 1) {
        return @()
    }

    $companyByNormalizedName = @{}
    $companyByCompactName = @{}
    $companyByStrippedName = @{}
    $companyByCompactStrippedName = @{}
    $companyByToken = @{}

    foreach ($company in @($Companies)) {
        $companyName = [string]$company.Name
        if ([string]::IsNullOrWhiteSpace($companyName)) { continue }

        $strippedName = Remove-AttributionLegalSuffixes $companyName
        foreach ($pair in @(
            @{ Table = $companyByNormalizedName; Key = (ConvertTo-AttributionNormalizedText $companyName) },
            @{ Table = $companyByCompactName; Key = (ConvertTo-AttributionCompactKey $companyName) },
            @{ Table = $companyByStrippedName; Key = $strippedName },
            @{ Table = $companyByCompactStrippedName; Key = (ConvertTo-AttributionCompactKey $strippedName) }
        )) {
            if ([string]::IsNullOrWhiteSpace([string]$pair.Key)) { continue }
            if (-not $pair.Table.ContainsKey($pair.Key)) {
                $pair.Table[$pair.Key] = [System.Collections.Generic.List[object]]::new()
            }
            $pair.Table[$pair.Key].Add($company)
        }

        foreach ($token in @((Get-AttributionSignificantTokens $companyName) + (Get-AttributionSignificantTokens $strippedName) | Sort-Object -Unique)) {
            Add-SharePointAttributionIndexValue -Table $companyByToken -Key $token -Value $company
        }
    }

    foreach ($entry in @($Entries)) {
        $explicitCompanyId = $entry.HuduCompanyId ?? $entry.CompanyId
        $explicitCompanyName = $entry.HuduCompanyName ?? $entry.CompanyName
        $explicitCompany = $null

        if ($explicitCompanyId) {
            $explicitCompany = @($Companies | Where-Object { [string]$_.Id -eq [string]$explicitCompanyId } | Select-Object -First 1)[0]
        }

        if (-not $explicitCompany -and -not [string]::IsNullOrWhiteSpace([string]$explicitCompanyName)) {
            $normalizedExplicitCompanyName = ConvertTo-AttributionNormalizedText $explicitCompanyName
            $explicitCompany = @(
                $Companies |
                    Where-Object { (ConvertTo-AttributionNormalizedText $_.Name) -eq $normalizedExplicitCompanyName } |
                    Select-Object -First 1
            )[0]
        }

        $candidates = if ($explicitCompany) {
            @(
                [PSCustomObject]@{
                    CompanyId   = $explicitCompany.Id
                    CompanyName = $explicitCompany.Name
                    Score       = 100
                    Reason      = 'client_file_explicit_hudu_company'
                }
            )
        } else {
            $exactCompanies = @()
            $exactReason = $null

            foreach ($exactMatch in @(
                @{ Table = $companyByNormalizedName; Key = $entry.NormalizedName; Reason = 'exact_normalized_name' },
                @{ Table = $companyByCompactName; Key = (ConvertTo-AttributionCompactKey $entry.ClientName); Reason = 'exact_compact_name' },
                @{ Table = $companyByStrippedName; Key = $entry.StrippedName; Reason = 'exact_legal_suffix_stripped' },
                @{ Table = $companyByCompactStrippedName; Key = (ConvertTo-AttributionCompactKey $entry.StrippedName); Reason = 'exact_compact_legal_suffix_stripped' }
            )) {
                if ([string]::IsNullOrWhiteSpace([string]$exactMatch.Key)) { continue }
                if (-not $exactMatch.Table.ContainsKey($exactMatch.Key)) { continue }

                $exactCompanies = @($exactMatch.Table[$exactMatch.Key])
                $exactReason = $exactMatch.Reason
                break
            }

            if ($exactCompanies.Count -gt 0) {
                @(
                    $exactCompanies |
                        ForEach-Object {
                            [PSCustomObject]@{
                                CompanyId   = $_.Id
                                CompanyName = $_.Name
                                Score       = 100
                                Reason      = $exactReason
                            }
                        }
                )
            } else {
                $candidateCompaniesById = [ordered]@{}
                foreach ($token in @((Get-AttributionSignificantTokens $entry.ClientName) + (Get-AttributionSignificantTokens $entry.StrippedName) | Sort-Object -Unique)) {
                    if (-not $companyByToken.ContainsKey($token)) { continue }

                    foreach ($company in @($companyByToken[$token])) {
                        $companyKey = [string]($company.Id ?? $company.Name)
                        if (-not $candidateCompaniesById.Contains($companyKey)) {
                            $candidateCompaniesById[$companyKey] = $company
                        }
                    }
                }

                if ($candidateCompaniesById.Count -gt 0) {
                    @(Get-HuduCompanyAttributionCandidates -ClientEntry $entry -Companies @($candidateCompaniesById.Values) | Sort-Object Score -Descending)
                } else {
                    @()
                }
            }
        }

        $best = $candidates | Select-Object -First 1
        $second = $candidates | Select-Object -Skip 1 -First 1
        $gap = if ($second) { [double]$best.Score - [double]$second.Score } elseif ($best) { [double]$best.Score } else { 0 }
        $autoMatched = ($best -and [double]$best.Score -ge $MinScore -and $gap -ge $MinGap)
        $aliases = [System.Collections.Generic.List[string]]::new()

        foreach ($alias in (@($entry.ClientName, $entry.StrippedName) + @($entry.ImportedAliases))) {
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
        [string[]]$FieldNames = @('Title', 'LinkTitle', 'Client Name', 'Client', 'Company', 'Name'),
        [int]$MinScore = 95,
        [int]$MinGap = 5
    )

    $entries = @(Get-SharePointClientListEntries -ManifestSet $ManifestSet -ListNames $ListNames -TitleFieldNames $FieldNames)
    if ($entries.Count -lt 1) {
        return @()
    }

    New-HuduClientAttributionMapFromEntries -Entries $entries -Companies $Companies -MinScore $MinScore -MinGap $MinGap
}

function New-SharePointClientDesignationMap {
    param (
        [Parameter(Mandatory)] $ManifestSet,
        [Parameter(Mandatory)] $AttributionMap,
        [array]$SelectedSites = @(),
        [string[]]$FieldNames = @("Select a Client", "Client", "Customer", "Company", "LinkTitle"),
        [double]$MinShare = 0.8,
        [int]$MinItems = 1,
        [int]$MinScore = 95,
        [int]$MinGap = 3
    )

    $siteVotes = @{}
    $listVotes = @{}
    $siteNames = @{}
    $listNames = @{}
    $sourceMatchCache = @{}
    $selectedSiteIds = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($site in @($SelectedSites)) {
        if ($site.id) { [void]$selectedSiteIds.Add([string]$site.id) }
    }

    function Add-SharePointClientDesignationVote {
        param (
            [Parameter(Mandatory)] [hashtable]$Votes,
            [Parameter(Mandatory)] [string]$Key,
            [Parameter(Mandatory)] $Match
        )

        if (-not $Votes.ContainsKey($Key)) {
            $Votes[$Key] = @{}
        }

        $companyId = [string]$Match.Entry.HuduCompanyId
        if ([string]::IsNullOrWhiteSpace($companyId)) { return }

        if (-not $Votes[$Key].ContainsKey($companyId)) {
            $Votes[$Key][$companyId] = [PSCustomObject]@{
                HuduCompanyId   = $Match.Entry.HuduCompanyId
                HuduCompanyName = $Match.Entry.HuduCompanyName
                Count           = 0
                MatchAlias      = $Match.Alias
            }
        }

        $Votes[$Key][$companyId].Count++
    }

    function Resolve-SharePointClientDesignationWinners {
        param (
            [Parameter(Mandatory)] [hashtable]$Votes,
            [Parameter(Mandatory)] [hashtable]$Names,
            [Parameter(Mandatory)] [string]$Scope
        )

        foreach ($key in @($Votes.Keys)) {
            $voteRows = @($Votes[$key].Values | Sort-Object Count -Descending)
            if ($voteRows.Count -lt 1) { continue }

            $total = 0
            foreach ($voteRow in $voteRows) { $total += [int]$voteRow.Count }
            if ($total -lt 1) { continue }

            $winner = $voteRows | Select-Object -First 1
            $share = [double]$winner.Count / [double]$total
            if ([int]$winner.Count -lt $MinItems -or $share -lt $MinShare) { continue }

            [PSCustomObject]@{
                Scope           = $Scope
                Key             = $key
                Name            = $Names[$key]
                HuduCompanyId   = $winner.HuduCompanyId
                HuduCompanyName = $winner.HuduCompanyName
                Votes           = [int]$winner.Count
                TotalVotes      = [int]$total
                Share           = [Math]::Round($share, 4)
                MatchAlias      = $winner.MatchAlias
                TopCandidates   = @($voteRows | Select-Object -First 5)
            }
        }
    }

    foreach ($manifest in @($ManifestSet.Manifests)) {
        foreach ($siteEntry in @($manifest.sites)) {
            $siteId = [string]$siteEntry.metadata.id
            if ($selectedSiteIds.Count -gt 0 -and -not $selectedSiteIds.Contains($siteId)) {
                continue
            }

            $siteName = $siteEntry.metadata.displayName ?? $siteEntry.metadata.name
            if ($siteId) { $siteNames[$siteId] = $siteName }

            foreach ($listEntry in @($siteEntry.lists)) {
                $listId = [string]$listEntry.metadata.id
                if ([string]::IsNullOrWhiteSpace($listId)) { continue }

                $listKey = "$siteId|$listId"
                $listNames[$listKey] = $listEntry.metadata.displayName ?? $listEntry.metadata.name

                foreach ($item in @($listEntry.items)) {
                    $sourceText = Get-SharePointListItemPrimaryAttributionSourceText -Item $item -FieldNames $FieldNames
                    if ([string]::IsNullOrWhiteSpace([string]$sourceText)) { continue }

                    $sourceKey = ConvertTo-AttributionCompactKey $sourceText
                    if ([string]::IsNullOrWhiteSpace($sourceKey)) { continue }

                    if ($sourceMatchCache.ContainsKey($sourceKey)) {
                        $match = $sourceMatchCache[$sourceKey]
                    } else {
                        $match = Resolve-HuduCompanyFromSharePointAttributionMap `
                            -SourceText $sourceText `
                            -AttributionMap $AttributionMap `
                            -AutoOnly `
                            -MinScore $MinScore `
                            -MinGap $MinGap
                        $sourceMatchCache[$sourceKey] = $match
                    }

                    if (-not $match -or -not $match.Entry.HuduCompanyId) { continue }

                    if ($siteId) {
                        Add-SharePointClientDesignationVote -Votes $siteVotes -Key $siteId -Match $match
                    }
                    Add-SharePointClientDesignationVote -Votes $listVotes -Key $listKey -Match $match
                }
            }
        }
    }

    $sites = @(Resolve-SharePointClientDesignationWinners -Votes $siteVotes -Names $siteNames -Scope 'Site')
    $lists = @(Resolve-SharePointClientDesignationWinners -Votes $listVotes -Names $listNames -Scope 'List')
    $siteById = @{}
    $listByKey = @{}

    foreach ($site in $sites) { $siteById[[string]$site.Key] = $site }
    foreach ($list in $lists) { $listByKey[[string]$list.Key] = $list }

    [PSCustomObject]@{
        Sites     = $sites
        Lists     = $lists
        SiteById  = $siteById
        ListByKey = $listByKey
    }
}

function Resolve-HuduCompanyFromClientDesignationMap {
    param (
        [string]$SiteId,
        [string]$ListId,
        $ClientDesignationMap,
        [switch]$UseSiteDesignation,
        [switch]$UseListDesignation
    )

    if (-not $ClientDesignationMap) { return $null }

    if ($UseListDesignation -and $SiteId -and $ListId -and $ClientDesignationMap.ListByKey) {
        $listKey = "$SiteId|$ListId"
        if ($ClientDesignationMap.ListByKey.ContainsKey($listKey)) {
            return $ClientDesignationMap.ListByKey[$listKey]
        }
    }

    if ($UseSiteDesignation -and $SiteId -and $ClientDesignationMap.SiteById) {
        if ($ClientDesignationMap.SiteById.ContainsKey($SiteId)) {
            return $ClientDesignationMap.SiteById[$SiteId]
        }
    }

    return $null
}

function Resolve-HuduCompanyFromSharePointAttributionMap {
    param (
        [Parameter(Mandatory)]
        [string]$SourceText,

        [Parameter(Mandatory)]
        $AttributionMap,

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
    if ($entry.HuduCompanyId) { return $entry }
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
