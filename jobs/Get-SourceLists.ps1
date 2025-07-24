if ($true -eq $RunSummary.SetupInfo.includeSPLists) {
    foreach ($site in $userSelectedSites) {
        $sitelists = Invoke-RestMethod -Headers $SharePointHeaders -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists" -Method GET
        if ($null -ne $sitelists -and $sitelists.value.Count -gt 0) {
            foreach ($siteList in $sitelists.value) {
                try {
                    # Fetch list schema (columns) – to map field types
                    $columnsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($siteList.id)/columns"
                    $columns = Invoke-RestMethod -Headers $SharePointHeaders -Uri $columnsUri -Method GET

                    # Fetch list items
                    $itemsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($siteList.id)/items?expand=fields"
                    $items = Invoke-RestMethod -Headers $SharePointHeaders -Uri $itemsUri -Method GET

                    # Build a simplified list entry
                    $fieldsSummary = @{}

                    foreach ($col in $columns.value) {
                        $fieldName = $col.displayName
                        $fieldType = Get-SPColumnType $col
                        $defaultValue = $col.defaultValue ?? "null"
                        $fieldsSummary[$fieldName] = @{
                            Type=$fieldType 
                            Default=$defaultValue
                            HuduFieldType=Get-SPListItemTypeToHuduALType -SPListItemType $fieldType -FieldName $fieldName -SampleItems $items.value
                            SampleData = $items.value
                        }
                    }

                    $DiscoveredLists += [PSCustomObject]@{
                        ListName   = $siteList.displayName
                        SiteName   = $site.name
                        Fields     = $fieldsSummary
                    }

                } catch {
                    Write-Warning "Could not process list '$($siteList.displayName)': $_"
                }
            }
        }
    } 
}

foreach ($list in $DiscoveredLists) {
    $layoutName="$($list.SiteName)-$($list.ListName)"
    write-host "searching for or creating $layoutName"
    $AssetLayout=$($(Get-HuduAssetLayouts -name "$layoutName") ?? $(New-HuduAssetLayout -name "$layoutName")).assetlayout
    $layoutFields = @()
    foreach ($field in $list.Fields){
        $layoutFields+=@{
        label= "string"
        show_in_list= $true
        field_type= "string"
        required= $true
        hint= "string"
        min= 0
        max= 0
        linkable_id= 0
        expiration= $true
        options= ""
        position=
      }
    }


}
