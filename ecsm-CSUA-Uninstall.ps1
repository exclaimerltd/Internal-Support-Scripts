#
<#
.SYNOPSIS
    Script to uninstall all instances of the Cloud Signature Update Agent
.DESCRIPTION
    This script is designed to be deployed via management software to all machines to remove the agent.
    It can be run manually via PowerShell.
    When run in the user context, it may not remove the ALLUSER MSI installations as these are deployed as admin
.NOTES
    Date: 8th July 2020
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    - Added commands to remove MSI installations
    - Added commands to remove Click Once Installations
#>

# Click Once uninstall
Write-Output "Uninstalling Click Once Installs"
$app = "Exclaimer Cloud Signature Update Agent"

$InstalledApplicationNotMSI = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall | foreach-object {Get-ItemProperty $_.PsPath}
$UninstallString = $InstalledApplicationNotMSI | ? { $_.displayname -match "$app" } | select UninstallString 

$wshell = new-object -com wscript.shell
$selectedUninstallString = $UninstallString.UninstallString
$wshell.run("cmd /c $selectedUninstallString")
Start-Sleep 5
$wshell.sendkeys("`"OK`"~")

# MSI Uninstall
Write-Output "Uninstalling MSI Installs"
$MyApp = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "Exclaimer Cloud Signature Update Agent"}
$MyApp.Uninstall() | Out-Null
Write-Output "All Cloud Signature Update Agent Installations removed"
