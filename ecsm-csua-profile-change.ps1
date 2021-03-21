#
<#
.SYNOPSIS
    PowerShell script to change the Outlook Profile address to a new domain
.DESCRIPTION
    As part of a ticket for a US company, an issue was identified where the agent would authenticate with an old domain after the company had changed their primary
    domain to another.  The client-side feature utilises the AccountName string value in the 
    Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\{ProfileName}\9375CFF0413111d3B88A00104B2A6676\00000002 registry key
.NOTES
    Email: helpdesk@exclaimer.com
    Date: 27th September 2019
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    Cloud Signature Update Agent
.HISTORY

#>

# Variables
$regkey="HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002"
$regstr="Account Name"
$olddomain="UPDATE WITH OLD DOMAIN"
$newdomain="UPDATE WITH NEW DOMAIN"

$currentaccount = (Get-ItemProperty -path $regkey).$regstr
$newaccount = $currentaccount.Replace($olddomain,$newdomain)

Set-ItemProperty -Path $regkey -Name $regstr -Value $newaccount