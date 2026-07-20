##### Step 2D, Export structured SharePoint lists for later asset import

$structuredListNames = @($RunSummary.SetupInfo.StructuredListJsonNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
if ($structuredListNames.Count -eq 0) {
    Set-PrintAndLog -message "No structured SharePoint list JSON exports configured." -Color DarkGray
    return
}

if ($null -eq $manifestSet) {
    Set-PrintAndLog -message "No SharePoint manifest set is available; skipping structured list JSON export." -Color Yellow
    return
}

Set-PrintAndLog -message "Exporting structured SharePoint lists as per-company JSON: $($structuredListNames -join ', ')" -Color Cyan

$structuredListExport = Export-SharePointStructuredListJson `
    -ManifestSet $manifestSet `
    -ListNames $structuredListNames `
    -AttributionMap ($ClientAttributionResolver ?? $ClientAttributionMap) `
    -OutputDirectory $RunSummary.OutputJsonFiles.StructuredListJsonDir `
    -IndexPath $RunSummary.OutputJsonFiles.StructuredListJsonIndex

Set-PrintAndLog -message "Structured list JSON export complete: $($structuredListExport.Items) item(s) in $($structuredListExport.Bundles) bundle(s)." -Color Cyan
Set-PrintAndLog -message "Structured list JSON directory: $($structuredListExport.OutputDirectory)" -Color DarkMagenta
Set-PrintAndLog -message "Structured list JSON index: $($structuredListExport.IndexPath)" -Color DarkMagenta
