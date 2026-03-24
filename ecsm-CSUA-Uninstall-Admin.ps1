<#
.SYNOPSIS
    Script to uninstall all instances of the Cloud Signature Update Agent
.DESCRIPTION
    This script is designed to be deployed via management software to all machines to remove the agent.
    It can be run manually via PowerShell.
    When run in the Admin/System context, it may not remove the ALLUSER MSI installations as these are deployed as admin
.NOTES
    Date: 8th July 2020
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    - Added commands to remove MSI installations
    - Added commands to remove Click Once Installations
#>

# Click Once uninstall
$app = "Exclaimer Cloud Signature Update Agent"

# -------------------------------
# Ensure the script is running with elevated permissions
# -------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host 'Elevated privileges are required. Relaunching as administrator...'
    Start-Sleep -Seconds 3
    exit 1
}

# -------------------------------
# Remove Exclaimer Agent Run keys (all users)
# -------------------------------

function RemoveUserRunKey {
    param($sid)

    $runPath = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Run"

    if (Test-Path $runPath) {
        Remove-ItemProperty -Path $runPath -Name "*Cloud Signature Update Agent" -ErrorAction SilentlyContinue
        Write-Host "Removed Run key for user hive $sid"
    }
}


$userHives = Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | 
    Where-Object { $_.PSChildName -match '^S-' -and $_.PSChildName.Length -ge 30 -and $_.PSChildName -notmatch '_Classes$' }

foreach ($hive in $userHives) {
    RemoveUserRunKey -sid $hive.PSChildName
}

$runKeyMachine = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
Remove-ItemProperty -Path $runKeyMachine -Name "Cloud Signature Update Agent" -ErrorAction SilentlyContinue
Write-Host "Removed HKLM Run key for Cloud Signature Update Agent"

$InstalledApplicationNotMSI = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall | foreach-object {Get-ItemProperty $_.PsPath}
$UninstallString = $InstalledApplicationNotMSI | Where-Object { $_.displayname -match "$app" } | Select-Object UninstallString 

if (!$UninstallString.UninstallString) {
Write-Output "No ClickOnce agent found"
}

if ($UninstallString.UninstallString) {
$wshell = new-object -com wscript.shell
$selectedUninstallString = $UninstallString.UninstallString
$wshell.run("cmd /c $selectedUninstallString")
Start-Sleep 5
$wshell.sendkeys("`"OK`"~")
Write-Output "ClickOnce agent removed"
}

# MSI Uninstall
$MyApp = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "Exclaimer Cloud Signature Update Agent"}
if (!$MyApp) {
Write-Output "No MSI installed agent found"
}

if ($MyApp) {
$MyApp.Uninstall() | Out-Null
Write-Output "MSI installed agent removed"
}
