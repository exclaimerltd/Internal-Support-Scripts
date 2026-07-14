# .SYNOPSIS
#     Script to enable the Exclaimer Cloud Add-in for all mailboxes where it is disabled.
#
# .DESCRIPTION
#     Will first check for required Module and prompt to install if not present.
#     It will then connect to Exchange Online and check the add-in status for every
#     active mailbox. Any mailbox where the add-in is present but disabled will have
#     it re-enabled. Results are exported to CSV.
#
# .NOTES
#     Created: 2026
#
# .PRODUCTS
#     Exclaimer Signature Management - Microsoft 365
#
# .REQUIREMENTS
#     - Global Administrator access to the Microsoft Tenant
#     - ExchangeOnlineManagement module
#
# .VERSION
#     1.0.0

param(
    [string]$AddInID = "efc30400-2ac5-48b7-8c9b-c0fd5f266be2",
    [switch]$VerboseLogging = $false
)

# Preferred output path (Downloads -> fallback C:\Temp)
$OutputPath = [System.IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), 'Downloads')

if (-not (Test-Path -Path $OutputPath)) {
    $OutputPath = "C:\Temp"
    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }
}

# File naming
$TimeStamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
$CsvFilePath    = Join-Path $OutputPath "EnableAddInForAllUsers_$TimeStamp.csv"
$TranscriptFile = Join-Path $OutputPath "Transcript_EnableAddIn_$TimeStamp.txt"

if ($VerboseLogging) {
    Write-Output "Verbose Logging Enabled: $TranscriptFile"
    Start-Transcript -Path $TranscriptFile -Append -NoClobber
}

function Ensure_ModuleInstalled {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "`nThe required 'ExchangeOnlineManagement' Module is NOT installed" -ForegroundColor Red
        $install = Read-Host "Do you want to install it? Y/n"
        if ($install -eq "Y") {
            try {
                Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
                Write-Host "'ExchangeOnlineManagement' installed successfully." -ForegroundColor Green
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

function Connect_ExchangeSession {
    try {
        $sessions = Get-PSSession | Select-Object -Property State, Name
        $active   = (@($sessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0

        if (-not $active) {
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

function Enable_AddInForAllMailboxes {
    $results = @()

    try {
        $mailboxes = Get-Mailbox -ResultSize Unlimited |
                     Where-Object { $_.AccountDisabled -eq $false }

        $total   = $mailboxes.Count
        $counter = 0

        Write-Host "`nProcessing $total mailboxes..." -ForegroundColor Green

        foreach ($mbx in $mailboxes) {
            $counter++
            Write-Progress -Activity "Checking and enabling add-in" `
                           -Status "$counter of $total : $($mbx.UserPrincipalName)" `
                           -PercentComplete (($counter / $total) * 100)

            $appIdentity = "$($mbx.UserPrincipalName)\$AddInID"
            $status      = $null
            $action      = $null
            $errorMsg    = $null

            try {
                $app    = Get-App -Identity $appIdentity -ErrorAction Stop
                $status = $app.Enabled

                if ($app.Enabled -eq $false) {
                    Enable-App -Identity $appIdentity -ErrorAction Stop
                    $action = "Enabled"
                    Write-Host "Re-enabled: $($mbx.UserPrincipalName)" -ForegroundColor Yellow
                } else {
                    $action = "AlreadyEnabled"
                }
            } catch {
                $errorMsg = $_.Exception.Message
                $action   = "Error"
            }

            $results += [pscustomobject]@{
                Mailbox  = $mbx.DisplayName
                UPN      = $mbx.UserPrincipalName
                WasEnabled = $status
                Action   = $action
                Error    = $errorMsg
            }
        }

        Write-Progress -Activity "Checking and enabling add-in" -Completed

        $results | Export-Csv -Path $CsvFilePath -NoTypeInformation
        Write-Host "`nOutput saved to: $CsvFilePath" -ForegroundColor Green

        # Summary
        $alreadyEnabled = ($results | Where-Object { $_.Action -eq "AlreadyEnabled" }).Count
        $reEnabled      = ($results | Where-Object { $_.Action -eq "Enabled" }).Count
        $errors         = ($results | Where-Object { $_.Action -eq "Error" }).Count

        Write-Host "`n--- Summary ---" -ForegroundColor Cyan
        Write-Host "Total mailboxes  : $total"
        Write-Host "Already enabled  : $alreadyEnabled" -ForegroundColor Green
        Write-Host "Re-enabled now   : $reEnabled"      -ForegroundColor Yellow
        Write-Host "Errors / absent  : $errors"         -ForegroundColor Red

    } catch {
        Write-Host "Fatal error retrieving mailbox list. Error: $_" -ForegroundColor Red
    }
}

function End_ExchangeSession {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "Error disconnecting. Error: $_" -ForegroundColor Red
    }
}

function Open_OutputDir {
    Start-Process $OutputPath
}

# Main execution
Ensure_ModuleInstalled
Connect_ExchangeSession
Enable_AddInForAllMailboxes
End_ExchangeSession
Open_OutputDir

if ($VerboseLogging) {
    Stop-Transcript
}