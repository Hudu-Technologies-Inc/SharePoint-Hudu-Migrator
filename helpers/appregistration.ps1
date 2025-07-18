function Set-AzureAppRegistration {
    param (
        [Parameter(Mandatory)]
        [string]$SAMDisplayName,
        [System.Collections.ArrayList]$delegatedPermissions=@("User.Read"),
        [System.Collections.ArrayList]$ApplicationPermissions=@("User.Read")

    )
    Write-Host "Starting App Registration create/update for $SAMDisplayName with $($delegatedPermissions.count) Delegated permissions and $($ApplicationPermissions.count) application permissions"

     # Alt URL for redirects
    $env:AZURE_CLI_ENCODING = "UTF-8"
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

    # install/import az module, set initial AZ context
    foreach ($module in @('Az')) {if (Get-Module -ListAvailable -Name $module) 
        { Write-Host "Importing module, $module"; Import-Module $module } else {Write-Host "Installing and importing module $module...... Please be patient, it can take a while."; Install-Module $module -Force -AllowClobber; Import-Module $module }
    }
    if (-not (Get-AzContext)) {
        try {
            Connect-AzAccount -UseDeviceAuthentication -ErrorAction Stop
        } catch {
            Write-Error "Failed to connect to Azure. Error: $_"
            exit 1
        }
    } else {
        Write-Host "AZContext already set. Skipping Sign-on."
    }
    # Checking for AZ CLI and installing
    Write-Host "Checking for AZ CLI..."
    $azCommand = Get-Command az -ErrorAction SilentlyContinue
    if (-not $azCommand) {
        Write-Host "AZ CLI command not found! If this script was ran as Administrator, we can install it, let's check."
        if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Error "You will need AZ CLI to run this script. If you want to install AZ CLI with this script, run again as admin, otherwise install AZ CLI Manually before running..."
            exit 1
        }
        try {
            Write-Host "Attempting WinGet install of AZ CLI via PowerShell module from PSGallery..."
            Write-Host "Setting up Nuget Provider for AZ CLI... $(Install-PackageProvider -Name NuGet -Force)"
            Write-Host "Setting up PSGallery for AZ CLI... $(Install-Module -Name Microsoft.WinGet.Client -Force -Repository PSGallery)"
            Write-Host "Ensuring Winget for AZ CLI... $(Repair-WinGetPackageManager)"
            Write-Host "Using Winget to install Azure CLI... $(winget install -e --id Microsoft.AzureCLI)"
        } catch {
            write-host "Couldnt install AZ CLI via Winget. Attempting install via MSI package."
            Write-Host "Downloading MSI Package to .\AzureCLI.msi... $(Invoke-WebRequest -Uri https://aka.ms/installazurecliwindowsx64 -OutFile .\AzureCLI.msi)"
            Write-Host "Installing AZ CLI via MSI Package in Background... $(Start-Process msiexec.exe -Wait -ArgumentList '/I AzureCLI.msi /quiet')"
            Write-Host "Clean up MSI Package file post-install... $(Remove-Item .\AzureCLI.msi -Force)"
        }
    } else { Write-Host "Azure CLI is available at: $($azCommand.Path)" }

    # Sign in with AZ CLI
    Write-Host "AZ CLI is present and available. Attempting login with AZ CLI if not already authenticated..."
    try {
        # Check if already logged in
        $account = az account show --output json | ConvertFrom-Json
        if ($account -and $account.id) {
            Write-Host "Already logged into Azure CLI as: $($account.user.name)"
        } else {
            throw "No active Azure account."
        }
    } catch {
        # If not logged in, log in
        Write-Host "Not logged into Azure CLI. Attempting to log in..."
        try {
            az login --use-device-code
            Write-Host "Successfully logged into Azure CLI."
        } catch {
            Write-Error "Failed to log in to Azure CLI. Exiting."
            exit 1
        }
    }

    # Step 1: List Available Tenant IDs (Optional)
    Write-Host "Step 1: Select Tenants"
    $Tenants = az account list --query "[].{Name:name, TenantId:tenantId, SubscriptionId:id}" --output json | ConvertFrom-Json
    Write-Host "Available Tenants:"
    Write-Host $Tenants
    $TenantId = Read-Host "Enter the Tenant ID to use (leave blank for current)"
    if ($TenantId -ne "") {
        az account set --tenant $TenantId
        Write-Host "Switched to Tenant: $TenantId"
    } else {
        $TenantId = az account show --query tenantId --output tsv
        Write-Host "Using current Tenant: $TenantId"
    }

    # Step 2, find or create app
    Write-Host "Step 2: Find or Create App"
    $AppId = az ad app list --all --display-name "$SAMDisplayName" --query "[?displayName=='$SAMDisplayName'].appId" --output tsv
    if ($AppId) {
        Write-Host "App '$SAMDisplayName' already exists with ID: $AppId"
    } else {
        Write-Host "Creating new App Registration: $SAMDisplayName"
        $AppId = az ad app create `
            --display-name "$SAMDisplayName" `
            --query appId --output tsv
        Write-Host "Created App with ID: $AppId"
    }

    # Step 3, find or create Service Principal
    Write-Host "Step 3: Find or Create Service Principal"
    $ServicePrincipalId = az ad sp list --all --query "[?appId=='$AppId'].id" --output tsv
    if (-not $ServicePrincipalId) {
        Write-Host "Creating Service Principal for App ID: $AppId"
        $ServicePrincipalId = az ad sp create --id "$AppId" --query id --output tsv
        Write-Host "Created Service Principal with ID: $ServicePrincipalId"
    } else {
        Write-Host "Service Principal already exists with ID: $ServicePrincipalId"
    }

    # Step 4: Get Service Principal Type(s) from user, desired permissions, scopes, roles
    Write-Host "Step 4: Get Service Principal Type(s) from user, desired permissions, scopes, roles"
    $GraphAppId = az ad sp list --filter "displayName eq 'Microsoft Graph'" --query "[].appId" --output tsv
    $PermissionsToAssign = @()
    $PermissionsToAssign = $DelegatedPermissions + $ApplicationPermissions

    # Step 5, Apply Desired AD Service Principal Permissions/Scopes
    Write-Host "Step 5, Apply Desired AD Service Principal Permissions/Scopes"
    $DelegatedScopes = @()
    $ApplicationRoles = @()
    foreach ($permission in $PermissionsToAssign) {
        Write-Host "Retrieving permission ID for '$permission'..."
        $PermissionId = az ad sp show --id $GraphAppId --query "appRoles[?value=='$permission'].id" --output tsv
        if (-not $PermissionId) {
            $PermissionId = az ad sp show --id $GraphAppId --query "oauth2PermissionScopes[?value=='$permission'].id" --output tsv
            if ($PermissionId) {
                Write-Host "Granting $permission ($PermissionId) as Scope (Delegated permission)..."
                $DelegatedScopes += "$PermissionId=Scope"
            } else {
                Write-Host "Warning: Permission '$permission' not found in Microsoft Graph API!"
            }
        } else {
            Write-Host "Granting $permission ($PermissionId) as Role (Application permission)..."
            $ApplicationRoles += "$PermissionId=Role"
        }
    }
    Write-Host "Step 5, Apply Desired AD Service Principal Permissions/Scopes"
    if ($DelegatedScopes.Count -gt 0) {
        Write-Host "Applying Delegated permissions..."
        foreach ($scope in $DelegatedScopes) {
            az ad app permission add --id $AppId --api $GraphAppId --api-permissions "$scope"
        }
    }
    if ($ApplicationRoles.Count -gt 0) {
        Write-Host "Applying Application permissions..."
        foreach ($role in $ApplicationRoles) {
            az ad app permission add --id $AppId --api $GraphAppId --api-permissions "$role"
        }
    }
    if ($DelegatedScopes.Count -gt 0) {
        Write-Host "Granting admin consent for Delegated permissions..."
        az ad app permission grant --id $AppId --api $GraphAppId --scope ($DelegatedScopes -replace "=Scope", "" -join " ")
    }
    if ($ApplicationRoles.Count -gt 0) {
        Write-Host "Granting admin consent for Application permissions..."
        az ad app permission admin-consent --id $AppId
    }
    # Print final summary in a structured way
    return @{
        AppId                   = $AppId
        TenantId                = $TenantId
        ServicePrincipalId      = $ServicePrincipalId
        AppSecret               = $(az ad app credential reset --id $AppId --query password --output tsv)
        DefaultAccount          = $($AzAccounts | Where-Object { $_.isDefault -eq $true })
        GraphAppId              = $GraphAppId
        Account                 = $account
        RegistrationUrl         = "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Authentication/appId/$AppId/isMSAApp~/false?Microsoft_AAD_IAM_legacyAADRedirect=true"

    }



}

function EnsureRegistration {
    param (
        [string]$clientId,
        [string]$tenantId
    )

    if ($clientId -and $tenantId) {
        return @{
            clientId = $clientId
            tenantId = $tenantId
        }
    }

    while ($true) {
        if ((Select-ObjectFromList -objects @("enter","new") -message "Would you like to enter your App Registration info or make a new Registration?") -eq "new") {
            $newAppRegResult = Set-AzureAppRegistration -SamDisplayName $(Get-SafeTitle $(Read-Host "Enter title for your new app registration")) `
                -delegatedPermissions @("Sites.Read.All", "Files.Read.All", "User.Read", "offline_access") `
                -ApplicationPermissions @("Files.Read.All", "Sites.Read.All")

            if ($newAppRegResult -and $newAppRegResult.TenantId -and $newAppRegResult.AppId) {
                $tenantId = $newAppRegResult.TenantId
                $clientId = $newAppRegResult.AppId
            } else {
                Write-Warning "App registration failed. Try again."
                continue
            }

            Write-Host "Now, make sure Device Code Flow is enabled in your app registration." -ForegroundColor Cyan
            Write-Host "`nTo do this:" -ForegroundColor Cyan
            Write-Host "1. Open your app registration in Azure Portal $($newAppRegResult.RegistrationUrl)" -ForegroundColor Cyan
            Write-Host "2. Go to Authentication > Advanced Settings" -ForegroundColor Cyan
            Write-Host "3. Under 'Allow public client flows', enable 'Allow device code flow'" -ForegroundColor Cyan
            Write-Host "`nOpening the Device Code Login URL to test your registration..."  -ForegroundColor Cyan
            Start-Process $($newAppRegResult.RegistrationUrl)
            Read-Host "Press Enter when Finished"
        } else {
            $tenantId = $tenantId ?? $(Read-Host "Enter your Microsoft Tenant ID")
            $clientId = $clientId ?? $(Read-Host "Enter your App Registration Client ID")

            if (-not $tenantId -or -not $clientId) {
                Write-Warning "Both Tenant ID and Client ID are required."
                continue
            }
        }

        if ($clientId -and $tenantId) { break }
    }

    return @{
        clientId = $clientId
        tenantId = $tenantId
    }
}