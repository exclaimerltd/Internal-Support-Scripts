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
    1. Open PowerShell 7+ as Administrator
    2. Execute the script:
        `.\BulkUserPhotoUpdate.ps1`
    3. Register an Azure AD application in the Azure Portal, instrcutions by the Script
    4. Store all user photos in a folder, with each file named as the user's alias (e.g., jsmith.jpg)
    5. (Automated by the Script) Install the required module if not already present
    6. Script requests the Client ID, Tenant ID, and Client Secret
    7. Review the log output for missing or failed photo updates
#>

# Clear the console for readability
function ConfirmPowerShellVersion {
    Write-Host "`n========== PowerShell Version Check ==========" -ForegroundColor Cyan

    $requiredMajorVersion = 7
    $currentVersion = $PSVersionTable.PSVersion

    if ($currentVersion.Major -lt $requiredMajorVersion) {
        Write-Host ""
        Write-Host "Unsupported PowerShell version detected." -ForegroundColor Red
        Write-Host "Current version : $currentVersion" -ForegroundColor Yellow
        Write-Host "Required version: PowerShell 7 or later" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Please install PowerShell 7 from:" -ForegroundColor White
        Write-Host "https://learn.microsoft.com/powershell/scripting/install/installing-powershell" -ForegroundColor Cyan
        Write-Host ""

        Write-Host "Press Enter to exit the script." -ForegroundColor Yellow
        Read-Host

        exit 0
    }

    Write-Host "PowerShell version $currentVersion is supported." -ForegroundColor Green
}

function ShowBulkUserPhotoUpdaterAppRegistrationGuide {
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
    Write-Host "1. In the App Registration, go to 'Manage' -> 'API Permissions' -> 'Add a permission' -> 'Microsoft Graph' -> 'Application permissions'." -ForegroundColor White
    Write-Host "2. Find and expand 'User' then select 'User.ReadWrite.All'." -ForegroundColor White
    Write-Host "3. Click 'Add permissions'." -ForegroundColor White
    Write-Host "4. Click 'Grant admin consent for <YourTenant>' (Global Administrator required)." -ForegroundColor White
    Write-Host ""

    Write-Host "Step 3: Create a Client Secret" -ForegroundColor Yellow
    Write-Host "1. Go to 'Certificates & secrets' -> 'New client secret'." -ForegroundColor White
    Write-Host "2. Provide a description and expiry period (1 or 2 years recommended)." -ForegroundColor White
    Write-Host "3. Copy the secret VALUE, not the ID. You will not be able to retrieve it again." -ForegroundColor White
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
}

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
        Write-Host "Please wait....." -ForegroundColor Yellow
        Start-Sleep -Seconds 10
    }
    catch {
        Write-Host "Failed to authenticate to Microsoft Graph." -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }
}

function UpdateUserPhotosByUpn {
    Clear-Host
    Write-Host ""
    Write-Host "========== Bulk User Photo Update ==========" -ForegroundColor Cyan

    $path = Read-Host "Enter the full path to the user photo directory"

    if (-not (Test-Path $path)) {
        Write-Host "The specified directory does not exist." -ForegroundColor Red
        return
    }

    # Report file in the SAME folder as photos
    $timestamp  = Get-Date -Format "yyyy-MM-dd_HHmm"
    $reportPath = Join-Path $path "BulkUserPhotoUpdateReport_$timestamp.csv"
    $report     = @()

    # Supported image formats
    $validExtensions = @("jpg", "jpeg", "png")

    $photoFiles = Get-ChildItem -Path $path -File |
        Where-Object { $validExtensions -contains $_.Extension.TrimStart('.').ToLower() }

    if (-not $photoFiles) {
        Write-Host "No image files found in the specified directory." -ForegroundColor Yellow
        return
    }

    $failedUploads = @()
    $missingFiles  = @()

    foreach ($file in $photoFiles) {
        $userPrefix = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

        Write-Host ""
        Write-Host "Processing file '$($file.Name)' for user prefix '$($userPrefix)'..." -ForegroundColor White

        try {
            $user = Get-MgUser -Filter "startsWith(userPrincipalName,'$userPrefix')" `
                -Property Id,UserPrincipalName -Top 1 -ErrorAction Stop

            if (-not $user) {
                Write-Host "No user found matching '$($userPrefix)'" -ForegroundColor Yellow
                $missingFiles += $userPrefix

                $report += [pscustomobject]@{
                    FileName = $file.Name
                    User     = $userPrefix
                    Status   = "User not found"
                }
                continue
            }

            $userUpn    = $user.UserPrincipalName
            $maxRetries = 3
            $attempt    = 0
            $success    = $false

            do {
                $attempt++
                try {
                    Set-MgUserPhotoContent -UserId $userUpn -InFile $file.FullName -ErrorAction Stop
                    Write-Host "Photo updated successfully for $($userUpn)" -ForegroundColor Green

                    $report += [pscustomobject]@{
                        FileName = $file.Name
                        User     = $userUpn
                        Status   = "Success"
                    }

                    $success = $true
                }
                catch {
                    Write-Host "Attempt $($attempt) failed for $($userUpn): $($_.Exception.Message)" -ForegroundColor Red

                    if ($attempt -lt $maxRetries) {
                        Write-Host "Waiting 5 seconds before retry..." -ForegroundColor Cyan
                        Start-Sleep -Seconds 5
                    }
                    else {
                        Write-Host "Maximum retries reached for $($userUpn)" -ForegroundColor DarkRed
                        $failedUploads += $userPrefix

                        $report += [pscustomobject]@{
                            FileName = $file.Name
                            User     = $userUpn
                            Status   = "Failed after retries"
                        }

                        $userChoice = PromptRetryOrSignOut -Message ("Failed to update photo for $($userUpn).")
                        if ($userChoice -eq "Retry") {
                            $attempt = 0
                            Write-Host "Retrying $($userUpn) in 10 seconds..." -ForegroundColor Cyan
                            Start-Sleep -Seconds 10
                        }
                        elseif ($userChoice -eq "SignOut") {
                            return
                        }
                    }
                }
            } while (-not $success -and $attempt -lt $maxRetries)
        }
        catch {
            Write-Host "Unexpected error for $($userPrefix): $($_.Exception.Message)" -ForegroundColor DarkRed
            $failedUploads += $userPrefix

            $report += [pscustomobject]@{
                FileName = $file.Name
                User     = $userPrefix
                Status   = "Unexpected error"
            }
        }
    }

    # Write report to the SAME folder as photos
    $report | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "========== Summary ==========" -ForegroundColor Cyan
    Write-Host "Report saved to:" -ForegroundColor Green
    Write-Host $reportPath -ForegroundColor Cyan

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

function PromptRetryOrSignOut {
    param (
        [string]$Message = "Would you like to try again?"
    )

    while ($true) {
        Write-Host ""
        Write-Host $Message -ForegroundColor Yellow
        Write-Host "Options:"
        Write-Host "  1) Try Again"
        Write-Host "  2) Sign Out and Disconnect"
        $choice = Read-Host "Enter 1 or 2"

        switch ($choice) {
            "1" {
                UpdateUserPhotosByUpn
            }
            "2" {
                Write-Host "Signing out and disconnecting..." -ForegroundColor Cyan
                # Disconnect from Microsoft Graph if connected
                DisconnectGraph
                return "SignOut"
            }
            default {
                Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red
            }
        }
    }
}
ConfirmPowerShellVersion
ShowBulkUserPhotoUpdaterAppRegistrationGuide
EnsureMgUsersModule
ConnectGraphAppOnly
UpdateUserPhotosByUpn
PromptRetryOrSignOut