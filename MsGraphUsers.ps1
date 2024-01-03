# .SYNOPSIS
#     Script to get and Output User data using Microsoft Graph.
# 
# .DESCRIPTION
#     Will first check for required Module and prompt to install them if not present.
#     It will then prompt the user to Sign in with Microsoft
#     Finally ask for details of search required and output that data into a file in "C:\Temp"
# 
# .NOTES
#     Email: helpdesk@exclaimer.com
#     Date: 3rd January 2024
# 
# .PRODUCTS
#     Exclaimer Signature Management - Microsoft 365
# 
# .REQUIREMENTS
#     - Global Administrator access to the Microsoft Tenant
#     - Requires path "C:\Temp"
#     - Microsoft.Graph.Authentication - https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/?view=graph-powershell-1.0
#     - Microsoft.Graph.Beta.Users - https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.beta.users/?view=graph-powershell-beta
# 
# .VERSION
# 
# 
# 	1.0.0
# 
# .INSTRUCTIONS
# 	- Open PowerShell as Administrator
# 	- Run: set-executionpolicy unrestricted
# 	- Go to directory where the Script is saved (i.e 'cd "C:\Users\ReplaceWithUserName\Downloads"')
# 	- Run the Script (i.e '.\MsGraphUsers.ps1')




#Setting variables to use later
$Path = "$PSScriptRoot\Exclaimer"

#Getting Exchange Online Module
function checkMicrosoftGraphUsersModule {
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) {
        Write-Host "`nThe Microsoft.Graph.Authentication Module is already installed" -ForegroundColor Green
    } 
    else {       
        Write-Host "The required 'Microsoft.Graph.Authentication' Module is NOT installed" -ForegroundColor Red
        $installMsGraph = Read-Host ("Do you want to install the required 'Microsoft.Graph.Authentication' Module? Y/n")
            if ($installMsGraph -eq "y") {
                Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
            }
            Else {
                Write-Host "We are unable to continue, now exiting" -ForegroundColor Red
                Exit
            }

    }
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Beta.Users) {
        Write-Host "`nThe Microsoft.Graph.Beta.Users Module is already installed`n" -ForegroundColor Green
    } 
    else {       
        Write-Host "The required 'Microsoft.Graph.Beta.Users' Module is NOT installed" -ForegroundColor Red
        $installMsGraph = Read-Host ("Do you want to install the required 'Microsoft.Graph.Beta.Users' Module? Y/n")
            if ($installMsGraph -eq "y") {
                Install-Module Microsoft.Graph.Beta.Users -Scope CurrentUser
            }
            Else {
                Write-Host "We are unable to continue, now exiting" -ForegroundColor Red
                Exit
            }
    }
}

#Connecting to MicrosoftGraph
function connect-MicrosoftGraph {
    Connect-MgGraph -Scopes "User.Read.All"
}


#Check "C:\Temp"
function checkTemp {
    if (Test-Path -Path $Path){
    Write-Output ("Output Path exists") | Out-File $OutFile
    }
    Else {
    New-Item $Path -ItemType Directory
    }
}

function findBy {

    try {
        [string[]]$options = 'Display Name', 'UPN', 'Email'

        Write-Output "Please choose what to search for:`n"

        1..$options.Length | foreach-object { Write-Output "$($_): $($options[$_-1])" }

        [ValidateScript({$_ -ge 1 -and $_ -le $options.Length})]
        [int]$number = Read-Host "`nPress a number to choose how to search"
            if($?) {
                Write-Output "You chose: $number`n"
            }
    } catch {
        Write-Host "Invalid entry"
        $tryAgain = Read-Host ("Do you want to search again? Y/n")
            if ($tryAgain -eq "y") {
                findBy
            }
            Else {            
                endSession
            }

    } finally {
        if ($number -eq 1) {
            $searchPropriety="DisplayName"         
            $searchPrompt = "Please enter the Display Name of the required user" 
            find-users
        }
        if ($number -eq 2) {     
            $searchPropriety="UserPrincipalName"   
            $searchPrompt = "Please enter the full UPN of the required user" 
            find-users
        }
        if ($number -eq 3) {
            $searchPropriety="ProxyAddresses"
            $searchPrompt = "Please enter an email address for the required user"
            find-users
        }
    }
}


function find-users{
    $DateTimeRun = (Get-Date -Format "dddd MM/dd/yyyy HH:mm '- UTC' K")            
    $searchText = Read-Host ("$searchPrompt")
    $getUsers = (Get-MgBetaUser | Where-Object {$_.$searchPropriety -like "*$searchText*"} | Select-Object DisplayName,Id) 
    Write-Host ("Starting search: $DateTimeRun") -ForegroundColor Green
    $found=$getUsers | Measure-Object
    Write-Host "Number of matches found:" $found.Count
    if (([string]::IsNullOrEmpty($getUsers))) {
    Write-Host "No matching users found..." -ForegroundColor Red
    tryAgain
    } else {
    foreach ($user in $getUsers) {
        $userId = $user.Id
        $userDN = $user.DisplayName
        $OutFile = "$Path\$userDN - $userId.txt"
        Write-Host "User:" $userDN
        userInfo
        }
    }    
    tryAgain
}


function userInfo {    
    $DateTimeRun = (Get-Date -Format "dddd MM/dd/yyyy HH:mm '- UTC' K")
    checkTemp
    Write-Output "$DateTimeRun" | Out-File $OutFile
    (Get-MgBetaUser -UserId $userId).PSObject.Properties | Where-Object {$_.Value -notlike "Microsoft*" -and $_.Name -notlike "Security*"} | Format-Table Name,Value | Out-File $OutFile -Append
    Write-Host "`nOutput is saved in $OutFile" -ForegroundColor Green
}

function tryAgain {
        $repeat = Read-Host ("Do you want to search again? Y/n")
        if ($repeat -eq "y") {
            findBy
        }
        Else {          
            Start-Process $Path
            endSession
        }
}

function openOutFile {
        $openOutFile = Read-Host ("Do you want to open the output file? Y/n")
            if ($openOutFile -eq "y") {
                Write-Host "Trying to open $OutFile`n" -ForegroundColor Green
                Start-Process $OutFile
            }
            Else {            
                Write-Host "`nWill not open the output file" -ForegroundColor Green
            }
}


checkMicrosoftGraphUsersModule
connect-MicrosoftGraph
findBy