# .SYNOPSIS
#     Script to check what version of the Exclaimer Cloud Add-in is on each mailbox.
# 
# .DESCRIPTION
#     Will first check for required Module and prompt to install them if not present.
#     It will then prompt the user to Sign in with Microsoft.
# 
# .NOTES
#     Created: 25th January 2024
#     Updated: 7th October 2024
# 
# .PRODUCTS
#     Exclaimer Signature Management - Microsoft 365
# 
# .REQUIREMENTS
#     - Global Administrator access to the Microsoft Tenant
#     - ExchangeOnlineManagement module
# 
# .VERSION 
#     1.1.0

param(
    [string]$AddInID = "efc30400-2ac5-48b7-8c9b-c0fd5f266be2",
    [switch]$VerboseLogging = $false
)

# Preferred output path (Downloads → fallback C:\Temp)
$OutputPath = [System.IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), 'Downloads')

if (-not (Test-Path -Path $OutputPath)) {
    $OutputPath = "C:\Temp"
    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }
}

# File naming
$TimeStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$CsvFilePath = Join-Path $OutputPath "GetAddInForAllUsers_$TimeStamp.csv"
$TranscriptLogFile = Join-Path $OutputPath "Transcript_$TimeStamp.txt"

# Start transcript if enabled
if ($VerboseLogging) {
    Write-Output "Verbose Logging Enabled: $TranscriptLogFile"
    Start-Transcript -Path $TranscriptLogFile -Append -NoClobber
}

# Function to check and install required module
function Ensure-ModuleInstalled {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "`nThe required 'ExchangeOnlineManagement' Module is NOT installed" -ForegroundColor Red
        $installMsModule = Read-Host "Do you want to install the required 'ExchangeOnlineManagement' Module? Y/n"
        if ($installMsModule -eq "Y") {
            try {
                Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
                Write-Host "`n'ExchangeOnlineManagement' Module installed successfully." -ForegroundColor Green
            } catch {
                Write-Host "Failed to install module. Error: $_" -ForegroundColor Red
                Exit
            }
        } else {
            Write-Host "Cannot continue without 'ExchangeOnlineManagement'. Exiting..." -ForegroundColor Red
            Exit
        }
    } else {
        Write-Host "`nThe 'ExchangeOnlineManagement' Module is already installed" -ForegroundColor Green
    }
}

# Function to connect to Exchange Online
function Connect-ExchangeSession {
    try {
        $getsessions = Get-PSSession | Select-Object -Property State, Name
        $session = (@($getsessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0

        if (-not $session) {
            Write-Host "Starting a new session..." -ForegroundColor Green
            Connect-ExchangeOnline
        } else {
            Write-Host "Exchange Online session is already active." -ForegroundColor Green
        }
    } catch {
        Write-Host "Failed to connect to Exchange Online. Error: $_" -ForegroundColor Red
        Exit
    }
}

# Function to gather mailbox add-in versions
function Get-MailboxAddInVersions {
    $results = @()
    try {
        $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.AccountDisabled -eq $false }

        Write-Host "`nGathering information, please wait..........." -ForegroundColor Green
        Write-Host "`nNote:" -ForegroundColor Red
        Write-Host "Some mailboxes may not list an Add-in, you should try again after the specific mailbox has logged on to 'https://outlook.office.com/'...`n" -ForegroundColor Yellow

        foreach ($mailbox in $mailboxes) {
            $appIdentity = ($mailbox.UserPrincipalName -split "@")[0] + "\" + $AddInID
            $appVersion = Get-App -Identity $appIdentity -ErrorAction SilentlyContinue

            $results += [pscustomobject]@{
                Mailbox            = $mailbox.DisplayName
                AppVersion         = $appVersion.AppVersion
                Enabled            = $appVersion.Enabled
                Deployment_Method  = $appVersion.Scope
            }
        }

        # Export results
        $results | Export-Csv -Path $CsvFilePath -NoTypeInformation
        Write-Host "Output saved to $CsvFilePath" -ForegroundColor Green

    } catch {
        Write-Host "Failed to retrieve mailbox information. Error: $_" -ForegroundColor Red
    }
}

# Function to end the session
function End-ExchangeSession {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "Error disconnecting session. Error: $_" -ForegroundColor Red
    }
}

# Open output directory
function Open-OutputDir {
    Start-Process $OutputPath
}

# Main execution
Ensure-ModuleInstalled
Connect-ExchangeSession
Get-MailboxAddInVersions
End-ExchangeSession
Open-OutputDir

if ($VerboseLogging) {
    Stop-Transcript
}