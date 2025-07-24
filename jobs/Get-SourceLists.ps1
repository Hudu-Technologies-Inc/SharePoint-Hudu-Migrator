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
                        }
                    }

                    $DiscoveredLists += [PSCustomObject]@{
                        ListName   = $siteList.displayName
                        SiteName   = $site.name
                        Fields     = $fieldsSummary
                        SampleData = ($items.value | Select-Object -First 1).fields
                    }

                } catch {
                    Write-Warning "Could not process list '$($siteList.displayName)': $_"
                }
            }
        }
    } 
}
