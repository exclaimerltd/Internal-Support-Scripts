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
#>

# Function to connect to Office 365 in current Window
function o365_connect {
    # below connects to Office 365
    $credential = Get-Credential
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -Credential $credential -ConnectionUri https://ps.outlook.com/powershell -Authentication Basic -AllowRedirection
    Import-PSSession $session
}

# Function exchange connect
function exchange_connect {
    add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
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


