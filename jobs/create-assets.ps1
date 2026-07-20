$ispFieldMap = @{
  "uxyn"       = "Name"
  "_x0067_ay1" = "Site Name"
  "saaf"       = "Provider"
  "xufd"       = "Type"
  "voyg"       = "Use"
  "m0mg"       = "Support #"
  "llh6"       = "Hardware Location"
  "fmwe"       = "Hardware Model #"
  "ibmi"       = "Public IP/Subnet"
}

$networkDeviceFieldMap = [ordered]@{
  # SharePoint internal name => Hudu "Network Devices" asset field label
  "xogl"      = "Site"
  "Title"     = "Device Name"
  "skdq"      = "Type"
  "w5ud"      = "Model"
  "joiu"      = "Serial"
  "zucp"      = "Network"
  "wr9t"      = "IP (s)"
  "qutm"      = "Management Link"
  "h64q"      = "Physical Location"
  "Edit"      = "Edit"
}



$networkDevicePrimaryFieldMap = @{
  PrimaryManufacturer = "_x0069_gy8"
  PrimaryModel        = "w5ud"
  PrimarySerial       = "joiu"
}

function ConvertTo-HuduAssetFields {
    param(
        $SharePointFields,
        [System.Collections.IDictionary]$FieldMap
    )

    $huduFields = @()

    foreach ($internalName in $FieldMap.Keys) {
        $huduLabel = $FieldMap[$internalName]
        $value = if ($SharePointFields.PSObject.Properties[$internalName]) {
            $SharePointFields.PSObject.Properties[$internalName].Value
        } else {
            $null
        }

        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
            $huduFields += @{ $huduLabel = $value }
        }
    }

    return $huduFields
}

function Get-MappedSharePointFieldValue {
    param(
        $SharePointFields,
        [string]$InternalName
    )

    if ([string]::IsNullOrWhiteSpace($InternalName)) { return $null }
    if (-not $SharePointFields.PSObject.Properties[$InternalName]) { return $null }

    $value = $SharePointFields.PSObject.Properties[$InternalName].Value
    if ([string]::IsNullOrWhiteSpace([string]$value)) { return $null }

    return $value
}

 $printershashtable = $netdevices | Group-Object { $_.fields.LinkTitle } -AsHashTable -AsString

$printershashtable.GetEnumerator().name | ForEach-Object {
 $coname = "$_".trim()
 $netitems = $printershashtable["$_"]

 write-host "$_ has $($netitems.count) items"
 $company = $null;
 $company = Get-HuduCompanies -Name $coname | Select-Object -first 1
 if ($null -eq $company) {

    $company = New-HuduCompany -name $coname
    $company = Get-HuduCompanies -Name $coname | Select-Object -First 1
 }
 foreach ($item in $netitems) {
 $fields = ConvertTo-HuduAssetFields -SharePointFields $item.Fields -FieldMap $networkDeviceFieldMap
 $fields | format-table
 New-HuduAsset -Name $item.Fields.title -CompanyId $company.Id  -AssetLayoutId 31  -Fields $fields
 }
}