# 
<#
.SYNOPSIS
    Sets the ReportToOriginatorEnabled and ReportToOriginator fields to true for Distribution Groups
.DESCRIPTION
    If the field is set to false, it can cause a number of issues when emailing groups due to the emails sent
    not containing any sender envelope data or a return-path.  This causes knock on affects with spam and signature
    application.
 
    The script is a quicker corrective measure to the issue that setting it manually and can be used regularly to
    ensure all groups are set correctly.
.NOTES
    Email: helpdesk@exclaimer.com
    Date: 13th May 2017
    Updated: 1st October 2020
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    - Global Administrator Account
.VERSION
     1.0 - Set ReportToOriginatorEnabled to True
     2.0 - Updated to Support Modern Authentication
#>
 
<#function o365_connect {
    # below connects to Office 365
    $credential = Get-Credential
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -Credential $credential -ConnectionUri https://ps.outlook.com/powershell -Authentication Basic -AllowRedirection
    Import-PSSession $session
}#>

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
    # Check for O365 groups
    $groups = Get-DistributionGroup -Filter ('ReportToOriginatorEnabled -eq $False -and IsDirSynced -eq $False')
    $dirsync = Get-DistributionGroup -Filter ('ReportToOriginatorEnabled -eq $False -and IsDirSynced -eq $true')
 
    If ($groups -ne $null) {
        Write-Host ("Below are the Office 365 Groups currently set to False") -ForegroundColor Green
        Write-Output $groups | Select DisplayName,ReportToOriginatorEnabled | Format-Table
    }
    Else {
        If ($dirsync -ne $null) {
            Write-Host ("Below are the Office 365 groups sync'd from AD with the value of False") -ForegroundColor Green
            Write-Output $dirsync | Select DisplayName,ReportToOriginatorEnabled | Format-Table
            Exit
        }
        Else {
            Write-Host ("There are no groups with ReportToOriginatorEnabled set to False") -ForegroundColor Green
            Exit
        }
    }
}
 
function o365_change {
    $change = Read-Host ("Would you like to change these groups to True? Y/n")
 
    If ($change -eq "y") {
        $groups | Set-DistributionGroup -ReportToOriginatorEnabled $true
        Write-Host ("Group changes complete!") -ForegroundColor White
    }
     Write-Host ("Disconnecting from Microsoft 365.") -ForegroundColor Green
    Disconnect-ExchangeOnline | echo a
    Write-Host ("This script will now end") -ForegroundColor Green
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
