if ($RunSummary.SetupInfo.SPListsAsLayouts) {
    Set-PrintAndLog -message "Processing Lists as Asset Layouts" -Color Yellow
    foreach ($list in $DiscoveredLists) {
        $layoutName="$($list.SiteName)-$($list.ListName)"
        Set-PrintAndLog -message  "searching for or creating $layoutName"
        $AssetLayout=$($(Get-HuduAssetLayouts -name "$layoutName") ?? $(New-HuduAssetLayout -name "$layoutName")).assetlayout
        Set-PrintAndLog -message  "Layout Id $($AssetLayout.id) with $($list.Fields.Count) Fields and $($list.Values.Count) Values and $($list.LinkedFiles.Count) linked files..."
        $layoutFields = @()
        $PosIDX=500
        $includeFiles=[bool]$($list.LinkedFiles.Count -gt 0)
        $includeComments=$true
        $includePasswords=$false
        $includePhotos=$false
        foreach ($field in $list.Fields){
            Set-PrintAndLog -message  "Processing $(if ($field.Nullable) {'Nullable'} else {'Mandatory'}) $($field.Type) Field, $($field.Name), for $layoutName from site/list $($list.SiteName)/$($list.ListName)"
            if ($field.HuduFieldType -eq "image"){
                Set-PrintAndLog -message  "Genuine photo field $($field.Name) found for $layoutName, AL can include photos"
                $includePhotos = $true
            }
            if ($field.Type -eq "user"){
                Set-PrintAndLog -message  "User field $($field.Name) found for $layoutName, assuming AL can be passworded"
                $includePasswords = $true
            }
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
        $AssetLayout = $(Set-HuduAssetLayout -id $AssetLayout.Id -name $AssetLayout.Name 
                     -include_passwords $includePasswords -include_photos $includePhotos -include_comments $includeComments -include_files $includeFiles
                     -fields @($layoutFields)).assetlayout
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
