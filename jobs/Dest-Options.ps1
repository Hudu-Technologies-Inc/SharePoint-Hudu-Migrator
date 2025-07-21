##### Step 2B Select Dest Options
Set-PrintAndLog -message  "Getting All Companies and configuring destination options (Hudu-Side)" -Color Blue
$AllCompanies = Get-HuduCompanies
if ($AllCompanies.Count -eq 0) {
    Set-PrintAndLog -message  "Sorry, we didnt seem to see any Companies set up in Hudu... If you intend to attribute certain articles to certain companies, be sure to add your companies first!" -Color Red
}
$RunSummary.JobInfo.MigrationDest=$(Select-ObjectFromList -Objects @(
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
    }) -message "Configure Destination (Hudu-Side) Options- $($RunSummary.JobInfo.MigrationSource.OptionMessage) to where in Hudu?" -allowNull $false)


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
if ($(Select-ObjectFromList -allowNull $false -objects @("yes","no") -message "Would you like to Convert Excel Workbooks / Spreadsheets to Hudu Articles") -eq "no") {
    $RunSummary.SetupInfo.DisallowedForConvert.AddRange(@("xlsx","xls","ods","xlsm"))
}
if ($(Select-ObjectFromList -allowNull $false -objects @("yes","no") -message "Would you like to Convert Powerpoints / Presentations to Hudu Articles") -eq "no"){
    $RunSummary.SetupInfo.DisallowedForConvert.AddRange(@("pptx","ppt","odp","pptm"))
}

if ( $RunSummary.SetupInfo.DisallowedForConvert.count -gt 0) 
    {Set-PrintAndLog -Message "$($RunSummary.SetupInfo.DisallowedForConvert -join ', ') will be disallowed during conversion."}
else 
    {Set-PrintAndLog -Message "All file conversions allowed per user."}
