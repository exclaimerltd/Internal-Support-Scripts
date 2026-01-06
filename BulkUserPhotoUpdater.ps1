<#
.SYNOPSIS
    Bulk updates Microsoft 365 user photos using Microsoft Graph app-only authentication based on file names.

.DESCRIPTION
    This script automates updating user profile photos in Microsoft 365 by matching local image files
    to user aliases. Administrators store photos in a designated folder, naming each file with the
    corresponding user's alias (e.g., jsmith.jpg). The script will connect to Microsoft Graph
    using an Azure AD application with app-only permissions, verify required modules, and process
    all supported image files in the folder automatically (.jpg, .jpeg, .png). Missing or unmatched
    files are reported.

.NOTES
    Email: helpdesk@exclaimer.com
    Date: 6th January 2026
    Version: 1.1.0

.PRODUCTS
    Microsoft 365 / Office 365

.REQUIREMENTS
    - Global Administrator privileges to register an Azure AD app and grant permissions
    - PowerShell 7+ recommended
    - Internet connectivity
    - Microsoft.Graph.Users PowerShell module
    - Directory containing user photo files named with the user’s alias

.VERSION
    1.1.0
        - Automatically detects all supported image formats (.jpg, .jpeg, .png)
        - Resolves user aliases to full UserPrincipalName in Microsoft 365
        - Connects to Microsoft Graph using client credentials
        - Reads all photo files in the specified folder and updates corresponding users
        - Lists any missing or unmatched files
        - Logs successes and failures

.INSTRUCTIONS
    1. Register an Azure AD application in the Azure Portal:
        - Name the app (e.g., BulkUserPhotoUpdater)
        - Select account type: Accounts in this organizational directory only
        - Add Application permission: Microsoft Graph -> User.ReadWrite.All (Application)
        - Grant admin consent
        - Create a client secret and note the value
    2. Store all user photos in a folder, naming each file with the user's alias (e.g., jsmith.jpg)
    3. Set the `$Path` variable to the folder path
    4. Open PowerShell 7+ as Administrator
    5. Install the required module if not already present:
        `Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force`
    6. Update the script variables for Client ID, Tenant ID, and Client Secret
    7. Execute the script:
        `.\BulkUserPhotoUpdate.ps1`
    8. Review the log output for missing or failed photo updates
#>

# Clear the console for readability
Clear-Host

Write-Host "========================== Bulk User Photo Updater ==========================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Step 1: Azure AD App Registration" -ForegroundColor Yellow
Write-Host "This script will guide you through registering an Azure AD application needed for" `
           "app-only authentication to update user photos in Microsoft 365." -ForegroundColor White
Write-Host ""

Write-Host "Open your web browser and navigate to the following page:" -ForegroundColor Green
Write-Host "https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" -ForegroundColor Cyan
Write-Host ""

Write-Host "Instructions:" -ForegroundColor Yellow
Write-Host "1. Click 'New registration'." -ForegroundColor White
Write-Host "2. Give the app a name, e.g., 'BulkUserPhotoUpdater'." -ForegroundColor White
Write-Host "3. Supported account types: 'Accounts in this organizational directory only'." -ForegroundColor White
Write-Host "4. Redirect URI: Leave blank (not required for app-only authentication)." -ForegroundColor White
Write-Host "5. Click 'Register'." -ForegroundColor White
Write-Host ""

Write-Host "After registration, note down the following values (you will need them later):" -ForegroundColor Green
Write-Host "- Application (client) ID" -ForegroundColor Cyan
Write-Host "- Directory (tenant) ID" -ForegroundColor Cyan
Write-Host ""

Write-Host "Step 2: Assign Permissions" -ForegroundColor Yellow
Write-Host "1. In the App Registration, go to 'API Permissions' -> 'Add a permission' -> 'Microsoft Graph' -> 'Application permissions'." -ForegroundColor White
Write-Host "2. Select 'User.ReadWrite.All'." -ForegroundColor White
Write-Host "3. Click 'Add permissions'." -ForegroundColor White
Write-Host "4. Click 'Grant admin consent for <YourTenant>' (Global Administrator required)." -ForegroundColor White
Write-Host ""

Write-Host "Step 3: Create a Client Secret" -ForegroundColor Yellow
Write-Host "1. Go to 'Certificates & secrets' -> 'New client secret'." -ForegroundColor White
Write-Host "2. Provide a description and expiry period (1 or 2 years recommended)." -ForegroundColor White
Write-Host "3. Copy the secret value now. You will not be able to retrieve it again!" -ForegroundColor White
Write-Host ""

Write-Host "Step 4: Save these values for the script:" -ForegroundColor Green
Write-Host "- Application (client) ID" -ForegroundColor Cyan
Write-Host "- Directory (tenant) ID" -ForegroundColor Cyan
Write-Host "- Client Secret Value (of the secret you created)" -ForegroundColor Cyan
Write-Host ""

Write-Host "Once you have these, you can proceed to install the Microsoft.Graph.Users module and connect using app-only authentication." -ForegroundColor Magenta
Write-Host ""
Write-Host "Press Enter when ready to continue to module installation and connection..." -ForegroundColor Yellow
Read-Host

function EnsureMgUsersModule {
    Clear-Host
    Write-Host ""
    Write-Host "========== Microsoft Graph PowerShell Module Check ==========" -ForegroundColor Cyan

    $moduleName = "Microsoft.Graph.Users"

    $installedModule = Get-Module -ListAvailable -Name $moduleName

    if ($installedModule) {
        Write-Host "Module '$moduleName' is already installed." -ForegroundColor Green
        return
    }

    Write-Host "Module '$moduleName' is not installed." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "This script requires the Microsoft Graph Users module to continue." -ForegroundColor White
    Write-Host "It will be installed from the PowerShell Gallery." -ForegroundColor White
    Write-Host ""

    $confirmation = Read-Host "Do you want to install '$moduleName' now? (Y/N)"

    if ($confirmation -notin @("Y", "y")) {
        Write-Host "Module installation declined. Script cannot continue." -ForegroundColor Red
        Write-Host "Exiting..." -ForegroundColor Red
        exit 1
    }

    try {
        Write-Host "Installing '$moduleName'..." -ForegroundColor Cyan

        Install-Module -Name $moduleName `
                       -Repository PSGallery `
                       -Scope CurrentUser `
                       -Force `
                       -AllowClobber `
                       -ErrorAction Stop

        Write-Host "Module '$moduleName' installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to install '$moduleName'." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }
}

function ConnectGraphAppOnly {
    Write-Host ""
    Write-Host "========== Microsoft Graph App Authentication ==========" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "You will need the following values from Azure App Registration:" -ForegroundColor White
    Write-Host " - Client ID (Application ID)" -ForegroundColor White
    Write-Host " - Tenant ID (Directory ID)" -ForegroundColor White
    Write-Host " - Client Secret (secret VALUE, not the ID)" -ForegroundColor White
    Write-Host ""

    $clientId = Read-Host "Application (client) ID"
    $tenantId = Read-Host "Directory (tenant) ID"
    $clientSecret = Read-Host "Enter Client Secret (Value)" -AsSecureString

    Write-Host ""
    Write-Host "Connecting to Microsoft Graph using app-only authentication..." -ForegroundColor Cyan

    try {
        $credential = New-Object System.Management.Automation.PSCredential (
            $clientId,
            $clientSecret
        )

        Connect-MgGraph `
            -TenantId $tenantId `
            -ClientSecretCredential $credential `
            -NoWelcome `
            -ErrorAction Stop

        $context = Get-MgContext

        if ($context.AuthType -ne "AppOnly") {
            Write-Host "Authentication completed but is not app-only." -ForegroundColor Red
            Write-Host "AuthType detected: $($context.AuthType)" -ForegroundColor Red
            throw "Incorrect authentication type"
        }

        Write-Host "Successfully authenticated to Microsoft Graph (App-only)." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to authenticate to Microsoft Graph." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }
}

function UpdateUserPhotosByUpn {
    Write-Host ""
    Write-Host "========== Bulk User Photo Update ==========" -ForegroundColor Cyan

    $path = Read-Host "Enter the full path to the user photo directory"

    if (-not (Test-Path $path)) {
        Write-Host "The specified directory does not exist." -ForegroundColor Red
        return
    }

    # Supported image formats
    $validExtensions = @("jpg", "jpeg", "png")

    # Get all image files in the folder with valid extensions
    $photoFiles = Get-ChildItem -Path $path -File | Where-Object { $validExtensions -contains $_.Extension.TrimStart('.').ToLower() }

    if (-not $photoFiles) {
        Write-Host "No image files found in the specified directory." -ForegroundColor Yellow
        return
    }

    $failedUploads = @()
    $missingFiles = @()

    foreach ($file in $photoFiles) {
        $userPrefix = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        Write-Host ""
        Write-Host "Processing file '$($file.Name)' for user prefix '$userPrefix'..." -ForegroundColor White

        try {
            # Attempt to find the matching user by prefix
            $user = Get-MgUser -Filter "startsWith(userPrincipalName,'$userPrefix')" -Property Id,UserPrincipalName -Top 1 -ErrorAction Stop

            if ($null -eq $user) {
                Write-Host "No user found matching '$userPrefix'" -ForegroundColor Yellow
                $missingFiles += $userPrefix
                continue
            }

            $userUpn = $user.UserPrincipalName

            # Set the photo
            Set-MgUserPhotoContent -UserId $userUpn -InFile $file.FullName -ErrorAction Stop
            Write-Host "Photo updated successfully for $userUpn" -ForegroundColor Green
        }
        catch {
            if ($_.Exception.Message -match "ResourceNotFound") {
                Write-Host "No user found for '$userPrefix'" -ForegroundColor Yellow
                $missingFiles += $userPrefix
            }
            else {
                Write-Host "Failed to update photo for '$userPrefix'" -ForegroundColor Red
                Write-Host $_.Exception.Message -ForegroundColor DarkRed
                $failedUploads += $userPrefix
            }
        }
    }

    Write-Host ""
    Write-Host "========== Summary ==========" -ForegroundColor Cyan

    if ($missingFiles.Count -gt 0) {
        Write-Host "Users not found in Microsoft 365:" -ForegroundColor Yellow
        $missingFiles | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }
    }

    if ($failedUploads.Count -gt 0) {
        Write-Host "Photo uploads failed for the following users:" -ForegroundColor Red
        $failedUploads | ForEach-Object { Write-Host " - $_" -ForegroundColor Red }
    }

    if ($missingFiles.Count -eq 0 -and $failedUploads.Count -eq 0) {
        Write-Host "All user photos updated successfully." -ForegroundColor Green
    }
}


function DisconnectGraph {
    Clear-Host
    Write-Host ""
    Write-Host "========== Disconnecting from Microsoft Graph ==========" -ForegroundColor Cyan

    try {
        Disconnect-MgGraph -ErrorAction Stop
        Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to disconnect from Microsoft Graph." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}
function CheckGraphSession {
    Write-Host ""
    Write-Host "========== Microsoft Graph Session Check ==========" -ForegroundColor Cyan

    $context = Get-MgContext

    if (-not $context) {
        Write-Host "No active Microsoft Graph session found." -ForegroundColor Green
        return
    }

    Write-Host "An active Microsoft Graph session is still present." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "AuthType : $($context.AuthType)" -ForegroundColor White
    Write-Host "ClientId : $($context.ClientId)" -ForegroundColor White
    Write-Host "TenantId : $($context.TenantId)" -ForegroundColor White
}


EnsureMgUsersModule
ConnectGraphAppOnly
UpdateUserPhotosByUpn
DisconnectGraph
CheckGraphSession