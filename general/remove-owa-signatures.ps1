# Remove OWA Signatures from all users
#{
<#
.SYNOPSIS
    Deletes all signatures from OWA users
.USAGE
    Script connects to email environment and deletes the current signatures present as well as turning off the auto-add function
.NOTES
    Email: support@exclaimer.com
    Date: 18th June 2018
.REQUIREMENTS
    None
.VERSION
    1.0 - Removes all OWA signatures
    1.1 - Updated to support Modern Authentication
#>

# Function to connect to Office 365 in current Window
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

# Function exchange connect
function exchange_connect {
    add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
}

$exchange = Read-Host ("Do you use Microsoft 365? Y/n")

if ($exchange -eq "y"){
    exchange_connect
}
Else {
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
}

# Function to remove all mailboxes signature
function remove-owa {
    $mailboxes = Get-Mailbox -ResultSize unlimited  
    $mailboxes | ForEach-Object { Set-MailboxMessageConfiguration -identity $_.alias -SignatureHtml "" }      
}
 
# Function to turn off autoadd
function remove-autoadd {
    $mailboxes = Get-Mailbox -ResultSize unlimited
    $mailboxes | ForEach-Object { Set-MailboxMessageConfiguration -identity $_.alias -autoaddsignature $false }
}

$environment = Read-Host ("Are you on Office 365? Y/n")

If ($environment -eq "y") {
    o365_connect
    remove-autoadd
    remove-owa
} else {
    exchange_connect
    remove-autoadd
    remove-owa
}


