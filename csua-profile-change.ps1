#
<#
.SYNOPSIS
    PowerShell script to change the Outlook Profile address to a new domain
.DESCRIPTION
    As part of a ticket for a US company, an issue was identified where the agent would authenticate with an old domain after the company had changed their primary
    domain to another.  The client-side feature utilises the AccountName string value in the 
    Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\{ProfileName}\9375CFF0413111d3B88A00104B2A6676\00000002 registry key
.NOTES
    Authored By: David Milward
    Email: support@exclaimer.com
    Date: 27th September 2019
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    Cloud Signature Update Agent
.HISTORY

#>

# Variables
$regkey="Computer\HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002"
$regstr="AccountName"
$olddomain="nccmedia.com"
$newdomain="ampersand.tv"

