# 
<#
.SYNOPSIS
    Sets the RequireSenderAuthenticationEnabled field to $false for all Office 365 Groups
.DESCRIPTION
    If the field is set to true it can cause the issue outlined here 
    (https://cloudsupport.exclaimer.com/hc/en-us/articles/207803529-Messages-sent-to-SharePoint-or-Office-365-groups-generate-an-NDR)
 
    This script aims to achieve the steps in that guide en masse
.NOTES
    Email: helpdesk@exclaimer.com
    Date: 30th August 2017
.PRODUCTS
    Exclaimer Cloud - Signature for Office 365
.REQUIREMENTS
    - Global Administrator Account     
#>
 
#>
function basic-auth-connect {
    $LiveCred = Get-Credential  
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session   
}

function modern-auth-mfa-connect {
    Import-Module ExchangeOnlineManagement
    $upn = Read-Host ("Enter the UPN for your Global Administrator")
    Connect-ExchangeOnline -UserPrincipalName $upn
}

function modern-auth-no-mfa-connect {
    Import-Module ExchangeOnlineManagement
    $LiveCred = Get-Credential
    Connect-ExchangeOnline -Credential $LiveCred
}
 
function o365_gather {
    # Check for O365 group
    $groups = Get-UnifiedGroup -filter * | ? {$_.RequireSenderAuthenticationEnabled -eq $true}
 
    If ($Ogroups -ne $null) {
    Write-Host ("Below are the Office 365 groups with the RequireSenderAuthenticationEnabled field set to true") -ForegroundColor Green
    Write-Host ("####") -ForegroundColor Green
    Write-Host ("")
    Write-Output $OGroups | Select Name,RequireSenderAuthenticationEnable | Format-Table
    Write-Host ("")
    Write-Host ("####") -ForegroundColor Green
    }
    Else {
        Write-Host ("There are no groups with RequireSenderAuthenticationEnabled set to true") -ForegroundColor Green
        Write-Host ("The script will now end") -ForegroundColor Green
        Return
        }
}
 
function o365_change {
    $change = Read-Host ("Do you want to change these groups to false? Y/n")
 
    If ($change -eq "y") {
        $Ogroups | Set-UnifiedGroup -RequireSenderAuthenticationEnabled $false
        Write-Host ("Group Changes Complete!")
    }
}

$authtype = Read-Host ("Do you have basic auth enabled? Y/n")

If ($authtype -eq "y") {
    basic-auth-connect
}
Else {
    $mfa = Read-Host ("Do you have MFA enabled? Y/n")
    if ($mfa -eq "y") {
        modern-auth-mfa-connect
    }
    Else {
        modern-auth-no-mfa-connect
    }
}

o365_gather
o365_change
 
Remove-PSSession $Session
Write-Host ("Script Complete!") -ForegroundColor Green