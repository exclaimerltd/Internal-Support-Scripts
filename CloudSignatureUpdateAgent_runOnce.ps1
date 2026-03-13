# <#
# .SYNOPSIS
#     Removes legacy Cloud Signature Update Agent Run keys and creates a user-specific scheduled task to run the agent once per logon.
#
# .DESCRIPTION
#     This script iterates through all local user profiles, removes any existing Run key entries for the Exclaimer Cloud Signature Update Agent in both HKCU and HKLM, 
#     and creates a per-user scheduled task to execute the agent at logon with a limited runtime of 15 minutes. 
#     The task is uniquely named per user to avoid conflicts and ensures offline signatures continue to function while reducing persistent ASR alert triggers.
#
# .NOTES
#     Date: 13th March 2026
#     Version: 1.0.0
#
# .PRODUCTS
#     Exclaimer Cloud Signature Update Agent
#
# .REQUIREMENTS
#     - PowerShell 5.1+ or PowerShell Core
#     - Administrative privileges to remove HKLM Run keys
#     - Local user profiles present under C:\Users
#     - Script executed in SYSTEM context for Intune deployment
#
# .VERSION
#     1.0.0
#         - Cleans per-user and machine-level Run keys for Cloud Signature Update Agent
#         - Detects all local user profiles and validates agent installation path
#         - Creates a scheduled task for each user to run agent at logon
#         - Scheduled task automatically terminates after 15 minutes
#         - Avoids task name conflicts by appending username
#
# .INSTRUCTIONS
#     **Deployment via Intune (Recommended for testing and production rollout):**
#     1. Save the script as `CloudSignatureUpdateAgent_runOnce.ps1`.
#     2. Go to Microsoft Endpoint Manager portal → Devices → Scripts → Add → Windows 10 and later.
#     3. Upload the PowerShell script.
#     4. Configure settings:
#        - Run script using logged-on credentials: **No** (system context required)
#        - Enforce script signature check: **No**
#        - Run script in 64-bit PowerShell Host: **Yes**
#     5. Assign the script to a test device group first.
#     6. Monitor deployment under Devices → PowerShell scripts → Device status.
#     7. Verify on test endpoints:
#        - Run keys removed from HKCU and HKLM
#        - Scheduled task exists per user
#        - Task runs at logon and stops after 15 minutes
# >

# Ensure the script is running with elevated permissions
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



# Function to remove Run key from a user hive
function Remove-UserRunKey {
    param($sid)
    $runPath = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Run"
    if (Test-Path $runPath) {
        Remove-ItemProperty -Path $runPath -Name "*Cloud Signature Update Agent" -ErrorAction SilentlyContinue
        Write-Host "Removed Run key for user hive $sid"
    }
}

# Enumerate all user SIDs under HKU (ignore _Classes)
$userHives = Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | 
    Where-Object { $_.PSChildName -match '^S-' -and $_.PSChildName.Length -ge 30 -and $_.PSChildName -notmatch '_Classes$' }

foreach ($hive in $userHives) {
    Remove-UserRunKey -sid $hive.PSChildName
}

# Remove machine-level Run key entry
$runKeyMachine = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
Remove-ItemProperty -Path $runKeyMachine -Name "Cloud Signature Update Agent" -ErrorAction SilentlyContinue
Write-Host "Removed HKLM Run key for Cloud Signature Update Agent"

# -------------------------------
# Create Scheduled Task for each user
# -------------------------------

# Detect user profiles (skip system profiles)
$userProfiles = Get-ChildItem 'C:\Users' | Where-Object { $_.Name -notin @("Public","Default","Default User","DefaultAppPool") }

foreach ($userProfile in $userProfiles) {
    if (-not (Test-Path $userProfile.FullName)) { continue }

    $localAppData = Join-Path $userProfile.FullName "AppData\Local"
    $exePath = Join-Path $localAppData "Programs\Exclaimer Ltd\Cloud Signature Update Agent\Exclaimer.CloudSignatureAgent.exe"

    if (-not (Test-Path $exePath)) {
        Write-Host "Agent not found for profile $($userProfile.Name). Skipping task creation."
        continue
    }

    Write-Host "Found agent for profile $($userProfile.Name) at $exePath"

    # Use username in task name to avoid collisions
    $taskName = "ExclaimerSignatureAgent_LogonRun"

    # Remove existing task if it exists
    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Existing task '$taskName' removed."
    }

    # Define task
    $action = New-ScheduledTaskAction -Execute $exePath
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries

    # Run task as the detected user
    $principal = New-ScheduledTaskPrincipal `
        -UserId $userProfile.Name `
        -LogonType Interactive `
        -RunLevel Limited

    # Register the scheduled task
    Register-ScheduledTask `
        -TaskName $taskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Principal $principal `
        -Description "Runs Exclaimer Cloud Signature Update Agent at logon with limited runtime"

    Write-Host "Scheduled task '$taskName' created for user $($userProfile.Name)."
}

Write-Host "All Run keys cleaned and tasks created successfully."