if ($RunSummary.SetupInfo.SPListsAsLayouts) {
    Set-PrintAndLog -message "Processing Lists as Asset Layouts" -Color Yellow
    foreach ($list in $DiscoveredLists) {
        $layoutName="$($list.SiteName)-$($list.ListName)"
        if (-not $list.Fields.Values -or -not $list.Fields){
                Set-PrintAndLog -message "Skipping invalid layout with not enough fields"
                continue            
        }



        Set-PrintAndLog -message  "searching for or creating asset layout- $layoutName"
        $layoutIcon = $FontAwesomeIconMap[$(Select-ObjectFromList -allowNull $false -objects @($FontAwesomeIconMap.Keys) -message "Which Icon for layout $layoutName?")]
        $colorOptions = $HexColorMap.Keys | Sort-Object
        $selectedColorName = Select-ObjectFromList -allowNull $false -objects $colorOptions -message "Choose a color for $layoutName with icon $layoutIcon"
        $layoutColor = $HexColorMap[$selectedColorName]
        $layoutBackgroundColor = Get-ComplimentingBackgroundColor -HexColor $layoutcolor
        $TempLayoutFields = $BaseSPLayoutFields
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
                -fields @($TempLayoutFields)
        }
        $AssetLayout = $AssetLayout.asset_layout
        Set-PrintAndLog -message  "Layout Id $($AssetLayout.id) with $($list.Fields.Count) Fields and $($list.Values.Count) Values and $($list.LinkedFiles.Count) linked files..."

        $layoutFields = @()
        $PosIDX = 499

        foreach ($task in $list.Fields.Values) {
            if (-not $task.HuduFieldType -or -not $task.Name) {
                Set-PrintAndLog -message "Skipping invalid field with null type or name"
                continue
            }

            $newField = @{
                field_type   = $task.HuduFieldType
                label        = $task.Name
                show_in_list = $true
                required     = -not $task.Nullable
                hint         = "original default - $($task.Default)"
                position     = $PosIDX
            }
            if ($task.HuduFieldType -eq "ListSelect") {
                Set-PrintAndLog -message "Found $($task.Choices.Count) choices in '$($task.Name)'; Searching for or creating list for ListSelect Field"
                $ListName = "$($layoutName)-$($task.Name)"
                $huduList = Get-HuduLists -Name $ListName
                if (-not $huduList -and $task.Options) {
                    $huduList =New-HuduList -name $ListName -Items $task.Options
                }
                $newField.list_id = $huduList.id
                $newField.multiple_options = $task.MultipleChoice
            }
        }
        $layoutFields += $newField
        $PosIDX -= 1
    
        $layoutFields | ConvertTo-Json -Depth 10 | Out-File "$(join-path $logsFolder -ChildPath "debug-fields-$layoutName.json")" 
        $LayoutObject = Set-HuduAssetLayout -id $AssetLayout.Id -fields @($layoutFields)
        $AssetLayout = $LayoutObject.assetlayout

        
        $LayoutsCreated+=@{
            SourceList = $list
            layoutFields = $layoutFields    
            HuduLayout = $AssetLayout
        }


        if ($list.LinkedFiles.Count -gt 0){
            $RelationsToResolve += @{
                field_type   = "AssetTag"
                label        = $task.Name
                linked_files = $list.LinkedFiles
            }
        }
    }
     
    foreach ($layout in $LayoutsCreated | Where-Object {$_.id -and $_.id -ne $null -and $_.id -ne 0}) {
        Set-PrintandLog -message "setting $($(Set-HuduAssetLayout -id $layout.id -Active $true).asset_layout.name) as active and readying row-level attribution"         
        $listName = $LayoutsCreated.Sourcelist.ListName
        $siteName = $LayoutsCreated.Sourcelist.SiteName
        
        $RowIDX = 0
        foreach ($row in $LayoutsCreated.Sourcelist.Values) {
            $RowIDX = $RowIDX+1
            $newAsset = [PSCustomObject]@{
                RowIDX              = $RowIDX
                CompanyAttribution  = $null
                HuduAssetObject     = $null
                Layout              = $layout.HuduLayout
                Fields              = @{}
            }
            $fields = $row.fields
            Write-Host "Setting Attribution for- Row $RowIDX of $($LayoutsCreated.Sourcelist.Values.Count) from site '$siteName', list '$listName':"
            foreach ($key in $fields.PSObject.Properties.Name) {
                if ($Layout.layoutFields -contains $key) {
                    Write-Host "  $key = $($fields.$key)"
                    $newAsset.Fields.key = $fields.$key
                } else {
                    Write-Host "  $key = $($fields.$key) [skipped, not in layout]"
                }

            }
            switch ([int]$RunSummary.JobInfo.MigrationDest.Identifier) {
                0 { $newAsset.CompanyAttribution = $SingleCompanyChoice.id }
                1 { $newAsset.CompanyAttribution = $null }
                default {
                    $newAsset.CompanyAttribution = (
                        Select-ObjectFromList `
                            -message "Migrating Row from layout $($newAsset.Layout.name)... Which Company to attribute this to?" `
                            -objects $AllCompanies -allowNull $false
                    )
                }
            }
            Set-PrintAndLog -message  "Adding new asset under new hudu layout $($newAsset.Layout.Id) / $($newAsset.Layout.Name) for $($newAsset.CompanyAttribution.Id) / $($newAsset.CompanyAttribution.Name)"

            $newAsset.HuduAssetObject = $(New-HuduAsset -asset_layout_id $newAsset.Layout.Id -company_id $newAsset.CompanyAttribution.Id -fields $newAsset.Fields).asset
            $AssetsCreated += $newAsset
        }
    }

} else {
    Set-PrintAndLog -message "Processing Lists as Procedures and Procedure Tasks!" -Color Yellow
    $allProcedures =  Get-HuduProcedures
    $createdProcedures = [System.Collections.ArrayList]@()
    foreach ($list in $DiscoveredLists) {
        $newProcedure= [PSCustomObject]@{
            Name                = "$($list.SiteName)-$($list.ListName)"
            FoundProcedure      = $allProcedures | Where-Object {$_.name -eq $newProcedure.Name}
            CreatedProcedure    = $null
            Tasks               = $list.Fields.Values
            PreviewBlock        = Get-ProcedureTasksPreviewBlock -ProcedureTitle $newProcedure.Name -TaskList $list.Fields.Values
            CompanyAttribution  = $null
            FormattedTasks      = @()
            CreatedTasks        = @()
        Description         = "Procedure Migrated from Sharepoint$(if ($list.itemsUri) { " at $($list.itemsUri)"})"
        }
        Set-PrintAndLog -message  "$(if ($newProcedure.FoundProcedure) {'Updating Found'} else {'Creating New'}) procedure with $($newProcedure.Tasks.Count) tasks- $($newProcedure.Name)"
        switch ([int]$RunSummary.JobInfo.MigrationDest.Identifier) {
            0 { $newProcedure.CompanyAttribution = $SingleCompanyChoice.id }
            1 { $newProcedure.CompanyAttribution = $null }
            default {
                $newProcedure.CompanyAttribution = (
                    Select-ObjectFromList `
                        -message "Migrating $(if ($newProcedure.FoundProcedure) {'existing'} else {'new'}) Procedure: $newProcedure.PreviewBlock... Which company to migrate into?" `
                        -objects $Attribution_Options
                ).CompanyId
            }
        }



        Set-PrintAndLog -message  "Creating New Procedure $($newProcedure.Name) $(if ($newProcedure.CompanyAttribution) {"Attribution set to Company $($newProcedure.CompanyAttribution)"} else {'Global attribution Set'})"
        if ($null -ne $newProcedure.CompanyAttribution) {
            $newProcedure.CreatedProcedure = $(New-HuduProcedure -CompanyId $newProcedure.CompanyAttribution -Name $newProcedure.Name `
                                            -Description $newProcedure.description).procedure
        } else {
            $newProcedure.CreatedProcedure = $(New-HuduProcedure -Name $newProcedure.Name `
                                            -Description $newProcedure.description).procedure
        }
        foreach ($task in $newProcedure.Tasks) {
            $newTask=@{
                ProcedureId   = $($newProcedure.CreatedProcedure).id
                Name          = "$($task.Name)"
                Description   = "$($task.Name)$(if ($task.hint) {"- $($task.hint)"} else {''})$(if ($field.Choices) {  "Choices - $($field.Choices -join ', ' )"} else { '' })"
            }
            $newProcedure.FormattedTasks += $newTask


            $newProcedure.CreatedTasks += New-HuduProcedureTask @newTask
        }
        $createdProcedures += $newProcedure
    }

}