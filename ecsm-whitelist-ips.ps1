#
<#
.SYNOPSIS
    Whitelist Exclaimer IPs via a Transport Rule
.DESCRIPTION
    This script was created to resolve an issue with messages coming back from
    our service being incorrectly identified as spam by Microsoft.
    This workaround was provided as part of a Microsoft Premier Support Ticket
.NOTES
    Email: helpdesk@exclaimer.com
    Date: 15th August 2019
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    Global Administrator access
    Ability to connect to Office 365 via PowerShell
.HISTORY
    1.0 - Commit of original script
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
    
New-TransportRule -Name "Bypass Spam for Exclaimer IP" `
-SenderIpRanges `
104.40.229.156,52.169.0.179,`
191.237.4.149,104.209.35.28,`
104.210.80.79,13.70.157.244,`
51.140.37.132,51.141.5.228,`
52.172.222.27,52.172.38.8,`
52.233.37.155,52.242.32.10,`
51.5.241.184,51.4.231.63 `
-SetSCL "-1" `
-Enabled $true `
-SenderAddressLocation Header -Priority 0