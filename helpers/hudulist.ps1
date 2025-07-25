function New-HuduList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string[]]$Items
    )

    $payload = @{
        list = @{
            name = $Name
            list_items_attributes = @()
        }
    }

    foreach ($item in $Items) {
        $payload.list.list_items_attributes += @{ name = $item }
    }

    $jsonBody = $payload | ConvertTo-Json -Depth 100

    try {
        $response = Invoke-HuduRequest -Method POST -Resource "/api/v1/lists" -Body $jsonBody
        if ($response) {
            return $($response | ConvertFrom-Json -depth 6)
        }
    } catch {
        return $null
    }
}
function Get-HuduList {
    [CmdletBinding()]
    param(
        [string]$Name
    )

    $response = Invoke-HuduRequest -Method GET -Resource "/api/v1/lists"

    if (-not $response) {
        Write-Warning "⚠️ No lists returned from Hudu."
        return $null
    }

    # Flat array (not wrapped)
    $lists = $response

    if ($Name) {
        $filtered = $lists | Where-Object { $_.name -eq $Name }
        if ($filtered) {
            return $filtered
        } else {
            Write-Warning "No list found with name '$Name'"
            return $null
        }
    }

    return $lists
}
