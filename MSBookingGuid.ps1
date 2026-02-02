# -----------------------------------------------
# Microsoft Bookings GUID Extraction Script
# -----------------------------------------------


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

#CheckExchangeOnlineModule

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
    $Global:csvPath = Join-Path $FilePath $CSVFile
}

function GetMSBookingGuid {
    SetOutputPath
    # Retrieve all user mailboxes, extract UPN + ExchangeGuid (no hyphens), export to CSV
Get-EXOMailbox -ResultSize Unlimited `
    -RecipientTypeDetails UserMailbox `
    -Properties ExchangeGuid,PrimarySmtpAddress `
    -PropertySets StatisticsSeed |
Select-Object `
    @{ Name = "Email"; Expression = { $_.UserPrincipalName } },
    @{ Name = "MSBookingGUID"; Expression = { $_.ExchangeGuid.ToString('N') } } |
Export-Csv -Path $Global:csvPath -NoTypeInformation -Encoding UTF8

# Completion message
Write-Host "Export complete: $Global:csvPath"
}
GetMSBookingGuid

Read-Host "Press ENTER to close"
