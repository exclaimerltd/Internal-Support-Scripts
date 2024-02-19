# .SYNOPSIS
#     Script to check what version of the Exclaimer Cloud Add-in is on each mailbox.
# 
# .DESCRIPTION
#     Will first check for required Module and prompt to install them if not present.
#     It will then prompt the user to Sign in with Microsoft
# 
# .NOTES
#     Email: helpdesk@exclaimer.com
#     Created: 25th January 2024
#     Update: 19th February 2024
# 
# .PRODUCTS
#     Exclaimer Signature Management - Microsoft 365
# 
# .REQUIREMENTS
#     - Global Administrator access to the Microsoft Tenant
#     - ExchangeOnlineManagement - https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps
# 
# .VERSION 
# 
# 	1.0.0
# 
# .INSTRUCTIONS
# 	- Open PowerShell as Administrator
# 	- Run: set-executionpolicy unrestricted
# 	- Go to directory where the Script is saved (i.e 'cd "C:\Users\ReplaceWithUserName\Downloads"')
# 	- Run the Script (i.e '.\GetAddInForAllUsers.ps1')




#Setting variables to use later
$Path = "$PSScriptRoot\Exclaimer"
$addin = "efc30400-2ac5-48b7-8c9b-c0fd5f266be2"

#Getting Exchange Online Module
function checkRequiredModules {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        # If module is aleady installed, continue 
        Write-Host "`nThe 'ExchangeOnlineManagement' Module is already installed" -ForegroundColor Green
    } 
    else {       
        Write-Host "The required 'ExchangeOnlineManagement' Module is NOT installed" -ForegroundColor Red
        $installMsModule = Read-Host ("Do you want to install the required 'ExchangeOnlineManagement' Module? Y/n")
            if ($installMsModule -eq "y") {
                # If the module is not installed, offer to start Module install
                Install-Module ExchangeOnlineManagement -Scope CurrentUser
            }
            Else {
                # If user does not accept to install the module, terminates
                Write-Host "We are unable to continue, now exiting" -ForegroundColor Red
                Exit
            }

    }
}

# Check if Path exists
function checkPath {
    if (Test-Path -Path $Path){
    # Check for directory "Exclaimer" exists or not in the same directory as the script is stored
    Write-Output ("Output Path exists")
    }
    Else {
    # Creates directory "Exclaimer" if it does not exist in the same directory as the script store
    New-Item $Path -ItemType Directory
    }
}


function connectExchangeOnlineManagement {
    #Connecting to Connect-ExchangeOnline
    Connect-ExchangeOnline
}

function findMailboxes {
    $result = @()
    # Get all mailboxes that are not in a disabled state (intended to filter service mailboxes, but may cause it to miss some shared mailboxes)
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.AccountDisabled -eq $false}
    Write-Host "`nGathering information, please wait..........." -ForegroundColor Green
    foreach ($mailbox in $mailboxes) {
        # Extract the first part of the email address
        $appIdentity = ($mailbox.UserPrincipalName -split "@")[0] + "\efc30400-2ac5-48b7-8c9b-c0fd5f266be2"

        # Get the app version
        $appVersion = Get-App -Identity $appIdentity -ErrorAction SilentlyContinue

        # Add results to array
        [array]$result += New-Object psobject -Property @{
            Mailbox = $mailbox.DisplayName
            AppVersion = $appVersion.AppVersion
            Enabled =$appVersion.Enabled
        }

    }
    # Output it all to a file named "GetAddInForAllUsers.txt"
    $result | Out-File "$Path\GetAddInForAllUsers.txt"

}

function openOutPath {
            # Tries to open the directory the output is saved in, and provides full path
            Write-Host "`nTrying to open $Path`n" -ForegroundColor Green
            Start-Process $Path           
            Write-Host "Please find the output file in folder $Path`n" -ForegroundColor Green
}

function endSession {            
            # Disconnecting from Exchange Online
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Host "Session Ended" -ForegroundColor Green
            Break
            Exit
}

#Calling each function
checkRequiredModules
checkPath
connectExchangeOnlineManagement
findMailboxes
openOutPath
endSession