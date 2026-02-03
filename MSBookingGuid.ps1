<#
.SYNOPSIS
    Extracts Microsoft Bookings-compatible GUIDs for all Exchange Online user mailboxes.

.DESCRIPTION
    This script connects to Exchange Online and retrieves all user mailboxes, extracting each
    user's UserPrincipalName and ExchangeGuid. The ExchangeGuid is converted into a hyphen-less
    format required by Microsoft Bookings and exported to a CSV file for easy reuse.

    The script automatically checks for the Exchange Online Management PowerShell module,
    installs it if missing, and enforces elevated permissions only when installation is required.
    Output is saved to the user's Downloads folder by default, with a fallback to C:\Temp.

.NOTES
    Email: helpdesk@exclaimer.com
    Date: 2nd February 2026
    Version: 1.0.0

.PRODUCTS
    Microsoft 365 / Exchange Online / Microsoft Bookings

.REQUIREMENTS
    - Exchange Administrator or Global Administrator permissions
    - PowerShell 5.1 or later
    - Internet connectivity
    - ExchangeOnlineManagement PowerShell module
    - Interactive Microsoft 365 sign-in

.VERSION
    1.0.0
        - Automatically installs Exchange Online Management module if missing
        - Enforces elevation only when module installation is required
        - Connects interactively to Exchange Online
        - Retrieves all user mailboxes
        - Extracts UserPrincipalName and ExchangeGuid
        - Converts ExchangeGuid to Microsoft Bookings-compatible format
        - Exports results to a timestamped CSV file

.INSTRUCTIONS
    1. Open PowerShell (run as Administrator if module installation is required)
    2. Execute the script:
        `.\MSBookingGuid.ps1`
    3. Sign in when prompted
    4. Locate the generated CSV file in the Downloads folder or C:\Temp
    5. Use the exported GUIDs when configuring Microsoft Bookings
#>

Clear-Host
Write-Host "           -----------------------------------------------" -ForegroundColor Cyan
Write-Host "           |                 EXCLAIMER                   |" -ForegroundColor Yellow
Write-Host "           |     Microsoft Bookings GUID Extraction      |" -ForegroundColor Yellow
Write-Host "           -----------------------------------------------" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1


function ConnectExchangeOnlineModule {
    # Connect to Exchange Online
    Write-Host "You will be prompted to Sign in with Microsoft in order to continue." -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline
}
function Test-IsElevated {
    # Check if running with elevated permissions, if module installation required
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function CheckExchangeOnlineModule {
    # Check if module is or needs installing
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Write-Host "Exchange Online Management module is already installed." -ForegroundColor Green
        ConnectExchangeOnlineModule
        return
    }

    Write-Host "Exchange Online Management module not found." -ForegroundColor Yellow

    if (-not (Test-IsElevated)) {
        Write-Host ""
        Write-Host "Administrator permissions are required to install the Exchange Online Management module." -ForegroundColor Red
        Write-Host "Please restart PowerShell using 'Run as administrator' and re-run the script." -ForegroundColor Red
        Exit
    }

    Write-Host "PowerShell is running with elevated permissions." -ForegroundColor Green
    Write-Host ""

    $installChoice = Read-Host "Would you like to install it now? (Y/N)"
    if ($installChoice.ToUpper() -ne "Y") {
        Write-Host "Module installation skipped. Processing cancelled." -ForegroundColor Yellow
        Exit
    }

    try {
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false | Out-Null
        }

        $galleryTrusted = (Get-PSRepository -Name PSGallery).InstallationPolicy
        if ($galleryTrusted -ne "Trusted") {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }

        Install-Module ExchangeOnlineManagement -Force -AllowClobber
        Write-Host "Exchange Online Management module installed successfully." -ForegroundColor Green

        ConnectExchangeOnlineModule
    } catch {
        Write-Host "Failed to install the module: $($_.Exception.Message)" -ForegroundColor Red
        Exit
    }
}

CheckExchangeOnlineModule

function SetOutputPath {
    # Default path: Downloads folder
    $FilePath = [System.IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), 'Downloads')
    $csvFile = "MSBookingGuid_$(Get-Date -Format 'HHmmss').csv"

    # Check if the path exists, if not, use C:\Temp
    if (-not (Test-Path -Path $FilePath)) {
        $FilePath = "C:\Temp"
        if (-not (Test-Path -Path $FilePath)) {
            New-Item -Path $FilePath -ItemType Directory -Force | Out-Null
        }
    }
    # Define output CSV path on Desktop
    $Global:csvPath = Join-Path $FilePath $csvFile
}

function GetMailboxScopeFromGroup {
    Write-Host "`nIf you sync with Exclaimer only members of a specific Group, we recommend you run this script for only that same group." -ForegroundColor Yellow
    $choice = Read-Host "Do you want to limit the export to members of a specific group? (Y/N)"
    if ($choice.ToUpper() -ne "Y") {
        return $null
    }

    $groupEmail = Read-Host "Enter the group email address"
    Write-Host "Resolving group: $groupEmail" -ForegroundColor Cyan

    try {
        $rootGroup = Get-DistributionGroup -Identity $groupEmail -ErrorAction Stop
    } catch {
        Write-Host "Group not found. Exiting." -ForegroundColor Red
        Exit
    }

    $mailboxUpns = New-Object System.Collections.Generic.HashSet[string]

    # Get direct members of the entered group
    $members = Get-DistributionGroupMember -Identity $rootGroup.Identity -ResultSize Unlimited

    foreach ($member in $members) {

        # Direct user mailbox
        if ($member.RecipientType -eq "UserMailbox") {
            $mailboxUpns.Add($member.PrimarySmtpAddress.ToString()) | Out-Null
        }

        # Mail-enabled group
        elseif ($member.RecipientType -like "*Group*") {

            Write-Host "Processing nested group: $($member.PrimarySmtpAddress)" -ForegroundColor Yellow

            $nestedMembers = Get-DistributionGroupMember -Identity $member.Identity -ResultSize Unlimited

            foreach ($nested in $nestedMembers) {
                if ($nested.RecipientType -eq "UserMailbox") {
                    $mailboxUpns.Add($nested.PrimarySmtpAddress.ToString()) | Out-Null
                }
            }
        }
    }

    if ($mailboxUpns.Count -eq 0) {
        Write-Host "No user mailboxes found in group scope." -ForegroundColor Red
        Exit
    }

    Write-Host "Mailbox scope built. Total mailboxes: $($mailboxUpns.Count)" -ForegroundColor Green
    return $mailboxUpns
}

function GetMSBookingGuid {

    SetOutputPath

    $mailboxScope = GetMailboxScopeFromGroup
    $results = @()

    if ($mailboxScope) {
        Write-Host "Running export for scoped mailbox list." -ForegroundColor Cyan

        foreach ($upn in $mailboxScope) {
            $mbx = Get-EXOMailbox `
                -Identity $upn `
                -Properties ExchangeGuid `
                -ErrorAction SilentlyContinue

            if ($mbx) {
                $results += $mbx
            }
        }
    } else {
        Write-Host "Running export for all user mailboxes." -ForegroundColor Cyan

        $results = Get-EXOMailbox `
            -ResultSize Unlimited `
            -RecipientTypeDetails UserMailbox `
            -Properties ExchangeGuid
    }

    $results |
    Select-Object `
        @{ Name = "Email"; Expression = { $_.UserPrincipalName } },
        @{
            Name = "MSBookingGUID"
            Expression = {
                "$($_.ExchangeGuid.ToString('N'))@$($_.UserPrincipalName.Split('@')[1])"
            }
        } |
    Export-Csv -Path $Global:csvPath -NoTypeInformation -Encoding UTF8

    Write-Host "Export complete: $Global:csvPath" -ForegroundColor Green
}

GetMSBookingGuid

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "NEXT STEPS: Upload the CSV to Exclaimer" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host ""

Write-Host "1. Go to the Exclaimer portal and Sign in if prompted." -ForegroundColor White
Write-Host ""

Write-Host "2. Navigate to the 'Sender Management' page and 'User Details Upload' section." -ForegroundColor White
Write-Host ""

Write-Host "3. Select the CSV file generated by this script:" -ForegroundColor White
Write-Host "   $Global:csvPath" -ForegroundColor Cyan
Write-Host ""

Write-Host "4. When prompted for the upload type, select:" -ForegroundColor White
Write-Host "   UPDATE EXISTING" -ForegroundColor Yellow
Write-Host ""

Write-Host "   This ensures listed users are updated without overriding other user details." -ForegroundColor DarkGray
Write-Host ""

Write-Host "5. Complete the upload and review the results for any errors." -ForegroundColor White
Write-Host ""

Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "Upload preparation complete." -ForegroundColor Green
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host ""


Read-Host "`nPress ENTER to close"
