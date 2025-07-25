if ($true -eq $RunSummary.SetupInfo.includeSPLists) {
    foreach ($site in $userSelectedSites) {
        $sitelists = Invoke-RestMethod -Headers $SharePointHeaders -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists" -Method GET
        set-Printandlog -message "starting for site $($site.Name)"

        if ($null -ne $sitelists -and $sitelists.value.Count -gt 0) {
            foreach ($siteList in $sitelists.value) {
                set-Printandlog -message "starting for list $($siteList.displayName)"  -Color DarkBlue
                # Get Data for Site-List
                try {
                    set-Printandlog -message "Obtaining data for Site-List $($siteList.displayName)"  -Color Blue
                    if ($siteList.displayName -in $BlockedSPInternalLists) {
                            set-Printandlog -message "Skipping internal-only site $($sitelist.displayName)"
                            continue                        
                    }
                    
                    # Fetch list schema (columns) – to map field types
                    $columnsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($siteList.id)/columns"
                    $columns = Invoke-RestMethod -Headers $SharePointHeaders -Uri $columnsUri -Method GET

                    # Fetch list items
                    $itemsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($siteList.id)/items?expand=fields"
                    $items = Invoke-RestMethod -Headers $SharePointHeaders -Uri $itemsUri -Method GET

                } catch {
                    Write-ErrorObjectsToFile -ErrorObject @{
                        Message         = "Could not process list '$($siteList.displayName)' for site $($site.Name): $_"
                        Error           = $_
                    } -name "site-fetching-$($siteList.displayName)-$($site.Name)"
                    continue
                }

                # Validate and Parse Columns
                $ValidColumns=@()
                try {
                    set-Printandlog -message "Validating Columns for Site-List $($siteList.displayName)"  -Color Blue
                    foreach ($col in $columns.value) {
                        if ($BlockedSPInternalColumns -contains $col.displayName) {
                            set-Printandlog -message "Skipping internal-only column $($col.displayName)"
                            continue
                        }
                        if ($col.readOnly -eq $true -or $col.columnGroup -eq "_hidden" -or $col.hidden -eq $true) {
                            set-Printandlog -message "Skipping hidden or read-only column $($col.displayName)"
                            continue
                        }
                        set-Printandlog -message "Valid column $($col.displayName)"
                        $ValidColumns += $col
                    }
                    set-Printandlog -message "Validated Columns for site: $($site.Name) list: $($siteList.Name): $($columns.value.count) -> $($ValidColumns.count)"
                    if (-not $ValidColumns -or $ValidColumns.Count -lt 1) {
                        set-Printandlog -message "Skipping list with not enough valid columns - site: $($site.Name) list: $($siteList.displayName)"
                        $sitelist | ConvertTo-Json -Depth 45 | Out-File $(Join-Path $tmpfolder "nocols_$($siteList.Name).json")
                        continue
                    }
                } catch {
                    Write-ErrorObjectsToFile -ErrorObject @{
                        Message         = "Could not process columns '$($siteList.displayName)' for site $($site.Name): $_"
                        Error           = $_
                    } -name "site-col-validation-$($siteList.displayName)-$($site.Name)"
                    continue
                }

                # Build a simplified list entry for attached files/documents
                $fieldsSummary = @{}
                $linkedFiles = @()
                try {
                    set-Printandlog -message "Building Attachments Array for Site-List $($siteList.displayName)"  -Color Blue
                    $linkedFiles = $items.value | Where-Object {$_.fields.ContentType -eq "Document" -and $_.fields.FileLeafRef} | `
                        ForEach-Object {
                            [PSCustomObject]@{
                                Name       = $_.fields.FileLeafRef
                                ID         = $_.id
                                Created    = $_.fields.Created
                                Modified   = $_.fields.Modified
                                LinkField  = $_.fields.LinkFilenameNoMenu
                                CheckinComment = $_.fields._CheckinComment
                        }}
                } catch {
                    Write-ErrorObjectsToFile -ErrorObject @{
                        Message         = "Could not process columns '$($siteList.displayName)' for site $($site.Name): $_"
                        Error           = $_
                    } -name "site-linkedfiles-$($siteList.displayName)-$($site.Name)"
                    continue
                }

                # Translate valid columns in list
                try {
                    set-Printandlog -message "Translating Columns for Site-List $($siteList.displayName)" -Color Blue
                    foreach ($col in $ValidColumns) {
                        $fieldType = Get-SPColumnType $col
                        if (-not $fieldType) {
                            Write-Warning "Could not resolve type for column $($col.displayName)"
                            $fieldType = "Text"
                        }                        
                        $defaultValue = $col.defaultValue ?? $null
                        $choices = Get-SPColumnChoices -col $col
                        $hudutype = Get-SPListItemTypeToHuduALType -SPListItemType $fieldType -FieldName $col.displayName -SampleItems $items.value
                        $nullable = $true
                        $multi = [bool]($fieldType -eq 'multichoice')

                        $fieldsSummary[$col.displayName] = @{
                            Type           = $fieldType 
                            Default        = $defaultValue
                            HuduFieldType  = $hudutype
                            Name           = $col.displayName
                            Nullable       = $nullable
                            Choices        = $choices
                            MultipleChoice = $multi
                        }
                    }

                } catch {
                    Write-ErrorObjectsToFile -ErrorObject @{
                        Message         = "Could not process columns '$($siteList.displayName)' for site $($site.Name): $_"
                        Error           = $_
                    } -name "site-col-translation-$($siteList.displayName)-$($site.Name)"
                    continue
                }

                if (-not $items.value -or $items.value.Count -lt 1) {
                    set-Printandlog -message "Skipping list with no values from site $($site.name) list $($siteList.displayName)"
                    $sitelist | ConvertTo-Json -Depth 45 | Out-File $(Join-Path $tmpfolder "noitems_$($siteList.Name).json")
                    continue
                }

                $ValidColumns | ConvertTo-Json -Depth 45 | Out-File $(Join-Path $tmpfolder "list_columns_$($siteList.Name).json")

                Set-PrintAndLog -Message "List: $($siteList.displayName), Total Columns: $($columns.value.Count), Valid: $($ValidColumns.Count)"

                $DiscoveredLists += [PSCustomObject]@{
                    ListName        = $siteList.displayName
                    SiteName        = $site.name
                    Fields          = $fieldsSummary
                    Values          = @($items.value)
                    LinkedFiles     = $linkedFiles
                    itemsUri        = $itemsUri
                }


            }
        } else {
            set-Printandlog -message "No Site Lists Found for Site $($site.Name)!"
        }
    } 
}
# Create or Modify Asset Layouts to be what we see in lists.
