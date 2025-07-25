$linkedFilesAndFolders=@()

$SkippableInternalColumns=@(
    "Folder Child Count","Item Child Count","Comment count",
    "Check In Comment","Retention label","Compliance Asset Id","Label applied by",
    "Like count","Source Version (Converted Document)","Source Version","Modified By",
    "Label setting","Source Name (Converted Document)","Source Name","Copy Source",
    "Item is a Record","App Modified By","App Created By"

)

if ($true -eq $RunSummary.SetupInfo.includeSPLists) {
    foreach ($site in $userSelectedSites) {
        $sitelists = Invoke-RestMethod -Headers $SharePointHeaders -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists" -Method GET
        $originalSitelistcount = $sitelists.Count
        $sitelists = $sitelists | Where-Object {
            $_.displayName -notmatch '^App|Site Assets|Form Templates|Shared Documents|Documents|Style Library|PortalSiteList|Hub Settings|CSPViolationReportList|Content and Structure Reports'
        }
        $validSiteListCount = $sitelists.Count
        set-Printandlog -message "Validated Sitelists from $originalSitelistcount -> $validSiteListCount"
        if ($null -ne $sitelists -and $sitelists.value.Count -gt 0) {
            foreach ($siteList in $sitelists.value) {
                try {

                    
                    # Fetch list schema (columns) – to map field types
                    $columnsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($siteList.id)/columns"
                    $columns = Invoke-RestMethod -Headers $SharePointHeaders -Uri $columnsUri -Method GET

                    # Fetch list items
                    $itemsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($siteList.id)/items?expand=fields"
                    $items = Invoke-RestMethod -Headers $SharePointHeaders -Uri $itemsUri -Method GET

                    $ValidColumns=@()
                    foreach ($col in $columns.value) {
                        if ($SkippableInternalColumns -contains $col.displayName) {
                            set-Printandlog -message "Skipping internal-only column $($col.displayName)"
                            continue
                        }
                        if ($col.readOnly -eq $true -or $col.hidden -eq $true) {
                            set-Printandlog -message "Skipping hidden or read-only column $($col.displayName)"
                            continue
                        }
                        set-Printandlog -message "Valid column $($col.displayName)"
                        $ValidColumns += $col
                    }
                    if ($ValidColumns.Count -lt 1){
                        set-Printandlog -message "Skipping list with not enough valid columns- site: $($site.Name) list: $($siteList.Name)"
                        continue
                    }

                    set-Printandlog -message "Validated Columns for site: $($site.Name) list: $($siteList.Name): $($columns.value.count) -> $($validatedColumns.count)"

                    # Build a simplified list entry
                    $fieldsSummary = @{}
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

                    foreach ($col in $ValidColumns) {
                        $fieldType = Get-SPColumnType $col
                        $defaultValue = $col.defaultValue ?? $null
                        $choices = Get-SPColumnChoices -col $col
                        $fieldsSummary[$col.displayName] = @{
                            Type            = $fieldType 
                            Default         = $defaultValue
                            HuduFieldType   = Get-SPListItemTypeToHuduALType -SPListItemType $fieldType -FieldName $col.displayName -SampleItems $items.value
                            Name            = $col.displayName
                            Nullable        = [bool]$(Get-SPColumnNullable -values @($items.value))
                            Choices         = $choices
                            MultipleChoice  = [bool]$($fieldType -eq 'multichoice')
                        }
                    }

                if (-not $items.value -or $items.value.Count -lt 1) {
                    set-Printandlog -message "Skipping list with no values from site $($site.name) list $($siteList.displayName)"
                    continue
                }
                Set-PrintAndLog -Message "List: $($siteList.displayName), Total Columns: $($columns.value.Count), Valid: $($ValidColumns.Count)"

                $DiscoveredLists += [PSCustomObject]@{
                    ListName        = $siteList.displayName
                    SiteName        = $site.name
                    Fields          = $fieldsSummary
                    Values          = @($items.value)
                    LinkedFiles     = $linkedFiles
                    itemsUri        = $itemsUri
                }

                } catch {
                    Write-Warning "Could not process list '$($siteList.displayName)': $_"
                }
            }
        } else {
            set-Printandlog -message "No Site Lists Found for Site $($site.Name)!"
        }
    } 
}
# Create or Modify Asset Layouts to be what we see in lists.
