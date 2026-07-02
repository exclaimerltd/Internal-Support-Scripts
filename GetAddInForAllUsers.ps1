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
function Ensure_ModuleInstalled {
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
function Connect_ExchangeSession {
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
function Get_MailboxAddInVersions {
    $results = @()
    $errors  = @()

    try {
        $mailboxes = Get-Mailbox -ResultSize Unlimited |
                     Where-Object { $_.AccountDisabled -eq $false }

        $total   = $mailboxes.Count
        $counter = 0

        Write-Host "`nProcessing $total mailboxes..." -ForegroundColor Green

        foreach ($mailbox in $mailboxes) {
            $counter++
            Write-Progress -Activity "Checking Add-in versions" `
                           -Status "$counter of $total : $($mailbox.UserPrincipalName)" `
                           -PercentComplete (($counter / $total) * 100)

            # Use full UPN, not just the local-part
            $appIdentity = "$($mailbox.UserPrincipalName)\$AddInID"
            $appVersion  = $null
            $errorMsg    = $null

            try {
                $appVersion = Get-App -Identity $appIdentity -ErrorAction Stop
            } catch {
                $errorMsg = $_.Exception.Message
                $errors  += [pscustomobject]@{
                    Mailbox = $mailbox.DisplayName
                    UPN     = $mailbox.UserPrincipalName
                    Error   = $errorMsg
                }
            }

            $results += [pscustomobject]@{
                Mailbox           = $mailbox.DisplayName
                UPN               = $mailbox.UserPrincipalName  # added for traceability
                AppVersion        = $appVersion.AppVersion
                Enabled           = $appVersion.Enabled
                Deployment_Method = $appVersion.Type
                Error             = $errorMsg                   # visible in CSV
            }
        }

        Write-Progress -Activity "Checking Add-in versions" -Completed

        $results | Export-Csv -Path $CsvFilePath -NoTypeInformation
        Write-Host "Output saved to: $CsvFilePath" -ForegroundColor Green

        # Summary
        $successCount = ($results | Where-Object { $null -ne $_.AppVersion }).Count
        $emptyCount   = ($results | Where-Object { $null -eq $_.AppVersion -and $null -eq $_.Error }).Count
        $errorCount   = $errors.Count

        Write-Host "`n--- Summary ---" -ForegroundColor Cyan
        Write-Host "Total mailboxes : $total"
        Write-Host "With add-in     : $successCount" -ForegroundColor Green
        Write-Host "No add-in found : $emptyCount"   -ForegroundColor Yellow
        Write-Host "Errors          : $errorCount"   -ForegroundColor Red

        if ($errors.Count -gt 0) {
            $errorCsvPath = $CsvFilePath -replace '\.csv$', '_errors.csv'
            $errors | Export-Csv -Path $errorCsvPath -NoTypeInformation
            Write-Host "Error details   : $errorCsvPath" -ForegroundColor Red
        }

    } catch {
        Write-Host "Fatal error retrieving mailbox list. Error: $_" -ForegroundColor Red
    }
}

# Function to end the session
function End_ExchangeSession {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "Error disconnecting session. Error: $_" -ForegroundColor Red
    }
}

# Open output directory
function Open_OutputDir {
    Start-Process $OutputPath
}

# Main execution
Ensure_ModuleInstalled
Connect_ExchangeSession
Get_MailboxAddInVersions
End_ExchangeSession
Open_OutputDir

if ($VerboseLogging) {
    Stop-Transcript
}