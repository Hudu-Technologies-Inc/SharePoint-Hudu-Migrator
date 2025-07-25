if ($RunSummary.SetupInfo.SPListsAsLayouts) {
    Set-PrintAndLog -message "Processing Lists as Asset Layouts" -Color Yellow
    foreach ($list in $DiscoveredLists) {
        $layoutName="$($list.SiteName)-$($list.ListName)"
        Set-PrintAndLog -message  "searching for or creating asset layout- $layoutName"
        $layoutIcon = $FontAwesomeIconMap[$(Select-ObjectFromList -allowNull $false -objects @($FontAwesomeIconMap.Keys) -message "Which Icon for layout $layoutName?")]
        $layoutcolor = $HexColorMap[$(Select-ObjectFromList -allowNull $false -objects @($HexColorMap.Keys) -message "Choose a color for $layoutName with icon $layoutIcon")]
        $layoutBackgroundColor = Get-ComplimentingBackgroundColor -HexColor $layoutcolor
        $TempLayoutFields = @(@{label        = 'Imported from SharePoint'
                                field_type   = 'Date'
                                show_in_list = 'false'
                                position     = 501},
                                                @{
                                label        = 'ITGlue URL'
                                field_type   = 'Text'
                                show_in_list = 'false'
                                position     = 502})        
    
        $AssetLayout=$($(Get-HuduAssetLayouts -name "$layoutName") ?? $(New-HuduAssetLayout -name "$layoutName" -icon $layoutIcon -color $layoutBackgroundColor -icon_color $layoutcolor  -include_passwords $true -include_photos $true -include_comments $true -include_files $true -fields @($TempLayoutFields) )).assetlayout
        Set-PrintAndLog -message  "Layout Id $($AssetLayout.id) with $($list.Fields.Count) Fields and $($list.Values.Count) Values and $($list.LinkedFiles.Count) linked files..."
        $layoutFields = @()
        $PosIDX=500
        foreach ($field in $list.Fields){
            Set-PrintAndLog -message  "Processing $(if ($field.Nullable) {'Nullable'} else {'Mandatory'}) $($field.Type) Field, $($field.Name), for $layoutName from site/list $($list.SiteName)/$($list.ListName)"
            $NewField= @{
                field_type   = $field.Type
                label        = $field.Name
                show_in_list = $true
                required     = $field.Nullable
                hint         = "original default - $($field.default)"
                position     = $PosIDX
            }
            if ($field.HuduFieldType -eq "listselect"){
                Set-PrintAndLog -message  "found $($field.Choices.Count) choices in $(if ($field.MultipleChoice) {'as multiple-choice'} else {'as choice'}) field, $($field.Name) in $layoutName"
                $newField=$newField+ @{
                    options           = "$($field.Choices -join ', ')"
                    multiple_options  = $field.MultipleChoice
                }
            }
            $layoutFields += $NewField
            $PosIDX=$PosIDX-1
        }
        $AssetLayout = $(Set-HuduAssetLayout -id $AssetLayout.Id -name $AssetLayout.Name -fields $layoutFields).assetlayout 
        $relationsToResolve=if ($list.LinkedFiles.Count -gt 0){
            $PosIDX=$PosIDX-1
            $layoutFields += @{
                field_type   = "AssetTag"
                label        = $field.Name
                linkable_id  = "REPLACEME"
                linked_files = $list.LinkedFiles
                Position     = $PosIDX
            }
        } else {$null}
        if ($relationsToResolve){
            $AssetLayout | Add-Member -NotePropertyName relationsToResolve -NotePropertyValue @($relationsToResolve) -Force
        }

    }
}
