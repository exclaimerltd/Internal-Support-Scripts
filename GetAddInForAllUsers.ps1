# .SYNOPSIS
#     Script to check what version of the Exclaimer Cloud Add-in is on each mailbox.
# 
# .DESCRIPTION
#     Will first check for required Module and prompt to install them if not present.
#     It will then prompt the user to Sign in with Microsoft.
# 
# .NOTES
#     Email: helpdesk@exclaimer.com
#     Created: 25th January 2024
#     Updated: 7th October 2024
# 
# .PRODUCTS
#     Exclaimer Signature Management - Microsoft 365
# 
# .REQUIREMENTS
#     - Global Administrator access to the Microsoft Tenant
#     - ExchangeOnlineManagement - https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps
# 
# .VERSION 
#     1.1.0
# 
# .INSTRUCTIONS
#     - Open PowerShell as Administrator
#     - Run: set-executionpolicy unrestricted
#     - Go to the directory where the Script is saved (i.e 'cd "C:\Users\ReplaceWithUserName\Downloads"')
#     - Run the Script (i.e '.\GetAddInForAllUsers.ps1')

# Script Parameters
param(
    [string]$OutputPath = "$PSScriptRoot\Exclaimer",
    [string]$AddInID = "efc30400-2ac5-48b7-8c9b-c0fd5f266be2",
    [switch]$VerboseLogging = $false
)

# Setting up logging if enabled
if ($VerboseLogging) {
    $LogFile = "$OutputPath\ScriptLog.txt"
    Write-Output "Verbose Logging Enabled: $LogFile"
    Start-Transcript -Path $LogFile -Append -NoClobber
}

# Function to check and install required module
function Ensure-ModuleInstalled {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "`nThe required 'ExchangeOnlineManagement' Module is NOT installed" -ForegroundColor Red
        $installMsModule = Read-Host ("Do you want to install the required 'ExchangeOnlineManagement' Module? Y/n")
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

# Ensure output directory exists
function Ensure-DirectoryExists {
    if (-not (Test-Path -Path $OutputPath)) {
        try {
            New-Item $OutputPath -ItemType Directory -Force | Out-Null
            Write-Host "Created directory: $OutputPath" -ForegroundColor Green
        } catch {
            Write-Host "Failed to create directory: $OutputPath. Error: $_" -ForegroundColor Red
            Exit
        }
    } else {
        Write-Host "Output directory already exists: $OutputPath" -ForegroundColor Green
    }
}

# Function to connect to Exchange Online
function Connect-ExchangeSession {
    try {
        #$session = Get-Module ExchangeOnlineManagement | Format-Table -Property Name,Version
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
            [array]$results += [pscustomobject]@{
                Mailbox    = $mailbox.DisplayName
                AppVersion = $appVersion.AppVersion
                Enabled    = $appVersion.Enabled
                Deployment_Method       = $appVersion.Scope
            }
        }

        # Output to CSV
        $results | Export-Csv -Path "$OutputPath\GetAddInForAllUsers.csv" -NoTypeInformation
        Write-Host "Output saved to $OutputPath\GetAddInForAllUsers.csv" -ForegroundColor Green
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

#Open Ouput directory
function open-OutputDir {
    Start "$OutputPath"
}

# Main Script Execution
Ensure-ModuleInstalled
Ensure-DirectoryExists
Connect-ExchangeSession
Get-MailboxAddInVersions
End-ExchangeSession
open-OutputDir

if ($VerboseLogging) { Stop-Transcript }