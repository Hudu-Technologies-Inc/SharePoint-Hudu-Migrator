if ($RunSummary.SetupInfo.SPListsAsLayouts) {
    Set-PrintAndLog -message "Processing Lists as Asset Layouts" -Color Yellow
    foreach ($list in $DiscoveredLists) {
        $layoutName="$($list.SiteName)-$($list.ListName)"
        Set-PrintAndLog -message  "searching for or creating asset layout- $layoutName"
        $layoutIcon = $FontAwesomeIconMap[$(Select-ObjectFromList -allowNull $false -objects @($FontAwesomeIconMap.Keys) -message "Which Icon for layout $layoutName?")]
        $colorOptions = $HexColorMap.Keys | Sort-Object
        $selectedColorName = Select-ObjectFromList -allowNull $false -objects $colorOptions -message "Choose a color for $layoutName with icon $layoutIcon"
        $layoutColor = $HexColorMap[$selectedColorName]
        $layoutBackgroundColor = Get-ComplimentingBackgroundColor -HexColor $layoutcolor
        $TempLayoutFields = @(
            @{
                label        = 'Imported from SharePoint'
                field_type   = 'Text'
                show_in_list = 'false'
                position     = 500
            },
            @{
                label        = 'SharePoint URL'
                field_type   = 'Text'
                show_in_list = 'false'
                position     = 501
            },
            @{
                label        = 'Sharepoint ID'
                field_type   = 'Text'
                show_in_list = 'false'
                position     = 502
            }

        )   
    
        $AssetLayout = Get-HuduAssetLayouts -name "$layoutName"

        if (-not $AssetLayout) {
            Set-PrintAndLog -message "Creating layout $layoutName with icon $layoutIcon, background $layoutBackgroundColor, icon color $layoutColor and tempfields $($TempLayoutFields | ConvertTo-Json)"

            $AssetLayout = New-HuduAssetLayout -name "$layoutName" `
                -icon $layoutIcon `
                -color $layoutBackgroundColor `
                -icon_color $layoutColor `
                -include_passwords $true `
                -include_photos $true `
                -include_comments $true `
                -include_files $true `
                -fields $TempLayoutFields
        }
        $AssetLayout = $AssetLayout.asset_layout
        Set-PrintAndLog -message  "Layout Id $($AssetLayout.id) with $($list.Fields.Count) Fields and $($list.Values.Count) Values and $($list.LinkedFiles.Count) linked files..."

        $layoutFields = $TempLayoutFields
        $PosIDX = 499

        foreach ($field in $list.Fields.Values) {
            if (-not $field.HuduFieldType -or -not $field.Name) {
                Set-PrintAndLog -message "Skipping invalid field with null type or name"
                continue
            }

            $newField = @{
                field_type   = $field.HuduFieldType
                label        = $field.Name
                show_in_list = $true
                required     = -not $field.Nullable
                hint         = "original default - $($field.Default)"
                position     = $PosIDX
            }

            if ($field.HuduFieldType -eq "dropdown" -or $field.HuduFieldType -eq "checkboxes") {
                Set-PrintAndLog -message "Found $($field.Choices.Count) choices in '$($field.Name)'"
                $newField += @{
                    options          = $field.Choices
                    multiple_options = $field.MultipleChoice
                }
            }

            $layoutFields += $newField
            $PosIDX -= 1
        }
        $layoutFields | ConvertTo-Json -Depth 10 | Out-File "$(join-path $logsFolder -ChildPath "debug-fields-$layoutName.json")" 

        $LayoutObject = Set-HuduAssetLayout -id $AssetLayout.Id -name $AssetLayout.Name -fields @($layoutFields)
        $AssetLayout = $LayoutObject.assetlayout

        # $relationsToResolve=if ($list.LinkedFiles.Count -gt 0){
        #     $PosIDX=$PosIDX-1
        #     $layoutFields += @{
        #         field_type   = "AssetTag"
        #         label        = $field.Name
        #         linkable_id  = "REPLACEME"
        #         linked_files = $list.LinkedFiles
        #         Position     = $PosIDX
        #     }
        # } else {$null}
        # if ($relationsToResolve){
        #     $AssetLayout | Add-Member -NotePropertyName relationsToResolve -NotePropertyValue @($relationsToResolve) -Force
        # }

    }
}
