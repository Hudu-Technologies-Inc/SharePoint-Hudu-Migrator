##### Step 2B Select Dest Options
Set-PrintAndLog -message  "Getting All Companies and configuring destination options (Hudu-Side)" -Color Blue
$availableDestinationOptions = @(
[PSCustomObject]@{
    OptionMessage= "To a Single Specific Company in Hudu"
    Identifier = 0
},
[PSCustomObject]@{
    OptionMessage= "To Global/Central Knowledge Base in Hudu (generalized / non-company-specific)"
    Identifier = 1
}, 
[PSCustomObject]@{
    OptionMessage= "To Multiple Companies in Hudu - Let Me Choose for Each article ($($AllCompanies.count) available destination company choices)"
    Identifier = 2
},
[PSCustomObject]@{
    OptionMessage= "To Companies in Hudu - Match/Create one company per SharePoint site"
    Identifier = 3
})
$validDestinationOptions = @()
$AllCompanies = Get-HuduCompanies
$articleFeaturesAvailable = Get-HuduFeatureAvailability -Core_Feature articles
if ($AllCompanies.Count -eq 0) {
    Set-PrintAndLog -message  "Sorry, we didnt seem to see any Companies set up in Hudu... If you intend to attribute certain articles to certain companies, be sure to add your companies first! company attribution will be disabled otherwise." -Color Yellow
} elseif ($false -eq $articleFeaturesAvailable.companyKB) {
    Set-PrintAndLog -message  "Sorry the company knowledge base feature is not available in Hudu. Please switch this feature on in order to attribute articles to company. company attribution will be disabled otherwise." -Color Yellow
} else {
    Set-PrintAndLog -message  "Company KB is enabled and is a valid user option for destination." -Color Green
    $validDestinationOptions = $availableDestinationOptions | where-object {@(0,2,3) -contains $_.Identifier}
}
if ($true -eq $articleFeaturesAvailable.centralKB){
    Set-PrintAndLog -message  "Central KB is enabled and is a valid user option for destination." -Color Green        
    $validDestinationoptions += ($availableDestinationOptions | where-object {@(1) -contains $_.Identifier} | select-object -first 1)
} else {
    Set-PrintAndLog -message  "Central KB is not enabled and will not be available as a destination option." -Color Yellow
}
if ($validDestinationOptions.Count -eq 0 -or ($articleFeaturesAvailable.centralKB -eq $false -and $articleFeaturesAvailable.companyKB -eq $false)) {
    Set-PrintAndLog -message  "No valid destination options are available. Please enable either the Company KB (and create companies) or Central KB feature in Hudu." -Color Red
    exit 1
}


$RunSummary.JobInfo.MigrationDest=$(Select-ObjectFromList -Objects  -message "Configure Destination (Hudu-Side) Options- $($RunSummary.JobInfo.MigrationSource.OptionMessage) to where in Hudu?" -allowNull $false)


if ([int]$RunSummary.JobInfo.MigrationDest.Identifier -eq 0) {
    $SingleCompanyChoice=$(Select-ObjectFromList -Objects $AllCompanies -message "Which company to $($SourcePages.OptionMessage) articles to?")
    $Attribution_Options=[PSCustomObject]@{
        CompanyId            = $SingleCompanyChoice.Id
        CompanyName          = $SingleCompanyChoice.Name
        OptionMessage        = "Company Name: $($SingleCompanyChoice.Name), Company ID: $($SingleCompanyChoice.Id)"
        IsGlobalKB           = $false
}
    $RunSummary.JobInfo.MigrationDest.OptionMessage="$($RunSummary.JobInfo.MigrationDest.OptionMessage) (Company Name: $($SingleCompanyChoice.Name), Company ID: $($SingleCompanyChoice.Id))"
} elseif ([int]$RunSummary.JobInfo.MigrationDest.Identifier -eq 1) {
    $Attribution_Options+=[PSCustomObject]@{
        CompanyId            = 0
        CompanyName          = "Global KB"
        OptionMessage        = "No Company Attribution (Upload As Global/Central KnowledgeBase Article)"
        IsGlobalKB           = $true
    }    
} elseif ([int]$RunSummary.JobInfo.MigrationDest.Identifier -eq 2) {
    foreach ($company in $AllCompanies) {
        $Attribution_Options+=[PSCustomObject]@{
            CompanyId            = $company.Id
            CompanyName          = $company.Name
            OptionMessage        = "Company Name: $($company.Name), Company ID: $($company.Id)"
            IsGlobalKB           = $false
        }
    }
    $Attribution_Options+=[PSCustomObject]@{
        CompanyId            = 0
        CompanyName          = "Global KB"
        OptionMessage        = "No Company Attribution (Upload As Global/Central KnowledgeBase Article)"
        IsGlobalKB           = $true
    }
    $Attribution_Options+=[PSCustomObject]@{
        CompanyId            = -1
        CompanyName          = "None (SKIP FOR NOW)"
        OptionMessage        = "Skipped"
        IsGlobalKB           = $false
    }
} else {
    foreach ($company in $AllCompanies) {
        $Attribution_Options+=[PSCustomObject]@{
            CompanyId            = $company.Id
            CompanyName          = $company.Name
            OptionMessage        = "Company Name: $($company.Name), Company ID: $($company.Id)"
            IsGlobalKB           = $false
        }
    }
    $Attribution_Options+=[PSCustomObject]@{
        CompanyId            = 0
        CompanyName          = "Global KB"
        OptionMessage        = "No Company Attribution (Upload As Global/Central KnowledgeBase Article)"
        IsGlobalKB           = $true
    }
    $Attribution_Options+=[PSCustomObject]@{
        CompanyId            = -1
        CompanyName          = "None (SKIP FOR NOW)"
        OptionMessage        = "Skipped"
        IsGlobalKB           = $false
    }
}

$RunSummary.SetupInfo.LinkSourceArticles =[bool]($(Select-ObjectFromList -objects @("yes","no") -message "Would you like to include links to original SharePoint Documents in Hudu Articles") -eq "yes")
$RunSummary.SetupInfo.SourceFilesAsAttachments =[bool]($(Select-ObjectFromList -objects @("yes","no") -message "Would you like to include a copy of original Sharepoint Document as Attachments to Hudu Articles") -eq "yes")
if ($(Select-ObjectFromList -objects @("yes","no") -message "Would you like to Convert Excel Workbooks / Spreadsheets to Hudu Articles") -eq "no") {
    $RunSummary.SetupInfo.DisallowedForConvert.AddRange(@("xlsx","xls","ods","xlsm"))
}
if ($(Select-ObjectFromList -objects @("yes","no") -message "Would you like to Convert Powerpoints / Presentations to Hudu Articles") -eq "no"){
    $RunSummary.SetupInfo.DisallowedForConvert.AddRange(@("pptx","ppt","odp","pptm"))
}

if ( $RunSummary.SetupInfo.DisallowedForConvert.count -gt 0) 
    {Set-PrintAndLog -Message "$($RunSummary.SetupInfo.DisallowedForConvert -join ', ') will be disallowed during conversion."}
else 
    {Set-PrintAndLog -Message "All file conversions allowed per user."}
