#
<#
.SYNOPSIS
    Script to uninstall click once applications
.DESCRIPTION
    This script is designed to be deployed as part of a GPO
    to allow administrators to centrally manage the uninstall of the 
    Signature Manager Office 365 Edition and Exclaimer Cloud Click Once
    applications
.NOTES
    Email: support@exclaimer.com
    Date: 20th January 2019
.PRODUCTS
    Signature Manager Office 365 Edition
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    - Update the variable $clickonce with the name of the agent you wish to remove
        - Outlook Signature Update Agent or Cloud Signature Update Agent
#>

$clickonce = "REPLACE WITH NAME OF CLICK ONCE"

$InstalledApplicationNotMSI = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall | foreach-object {Get-ItemProperty $_.PsPath}
$UninstallString = $InstalledApplicationNotMSI | ? { $_.displayname -match "$clickonce" } | select UninstallString 

$wshell = new-object -com wscript.shell
$selectedUninstallString = $UninstallString.UninstallString
$wshell.run("cmd /c $selectedUninstallString")
Start-Sleep 5
$wshell.sendkeys("`"OK`"~")