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
#     1.1.0
#         - Cleans per-user and machine-level Run keys for Cloud Signature Update Agent
#         - Detects active user sessions via HKU Volatile Environment
#         - Validates Cloud Signature Update Agent installation path per user
#         - Creates a scheduled task for each detected user to run the agent at logon
#         - Scheduled task automatically terminates after 15 minutes
#         - Avoids task name conflicts by appending username
#         - Adds configurable overrideExistingTasks option:
#             • 0 = Skip task creation if task already exists (default, idempotent for Intune)
#             • 1 = Remove existing task and recreate it
#
# .INSTRUCTIONS
#     **Deployment via Intune (Recommended for testing and production rollout):**
#     1. Save the script as `CloudSignatureUpdateAgent_runOnce.ps1`.
#     2. Go to Microsoft Endpoint Manager portal → Devices → Scripts → Add → Windows 10 and later.
#     3. Upload the PowerShell script.
#     4. Configure settings:
#        - Run script using logged-on credentials: No (system context required)
#        - Enforce script signature check: No
#        - Run script in 64-bit PowerShell Host: Yes
#     5. Assign the script to a test device group first.
#     6. Monitor deployment under Devices → PowerShell scripts → Device status.
#     7. Verify on test endpoints:
#        - Run keys removed from HKCU and HKLM
#        - Scheduled task created per detected user
#        - Task runs at user logon and stops automatically after 15 minutes
#        - Script safely re-runs without recreating tasks unless overrideExistingTasks = 1
# >
# -------------------------------
# Choose to or not override existing tasks
# Set to "1" to remove existing tasks and create new ones, or "0" to skip task creation if a task with the same name already exists
# -------------------------------
$overrideExistingTasks = 0

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


# -------------------------------
# Create Scheduled Task for each active user session
# -------------------------------

$userHives = Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^S-1-5-21-' -and $_.PSChildName -notmatch '_Classes$' }

foreach ($hive in $userHives) {

    $hiveName = $hive.PSChildName
    $volEnvPath = "Registry::HKEY_USERS\$hiveName\Volatile Environment"

    if (-not (Test-Path $volEnvPath)) { continue }

    $envProps = Get-ItemProperty $volEnvPath

    $username = $envProps.USERNAME
    $userDomain = $envProps.USERDOMAIN
    $localAppData = $envProps.LOCALAPPDATA

    if (-not $username -or -not $localAppData) { continue }

    $exePath = Join-Path $localAppData "Programs\Exclaimer Ltd\Cloud Signature Update Agent\Exclaimer.CloudSignatureAgent.exe"

    if (-not (Test-Path $exePath)) {
        Write-Host "Agent not found for $userDomain\$username. Skipping task creation."
        continue
    }

    Write-Host "Found agent for $userDomain\$username at $exePath"

    $taskName = "ExclaimerSignatureAgent_LogonRun_$username"

    # Only check if task name exists
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if ($existingTask) {

        if ($overrideExistingTasks -eq 1) {
            Write-Host "Scheduled task '$taskName' already exists. Override enabled. Recreating."
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        }
        else {
            Write-Host "Scheduled task '$taskName' already exists. Skipping."
            continue
        }

    }

    $action = New-ScheduledTaskAction -Execute $exePath

    $trigger = New-ScheduledTaskTrigger `
        -AtLogOn `
        -User "$userDomain\$username"

    $settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries

    $principal = New-ScheduledTaskPrincipal `
        -UserId "$userDomain\$username" `
        -LogonType Interactive `
        -RunLevel Limited

    Register-ScheduledTask `
        -TaskName $taskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Principal $principal `
        -Description "Runs Exclaimer Cloud Signature Update Agent at logon with limited runtime"

    Write-Host "Scheduled task '$taskName' created for $userDomain\$username"
}

Write-Host "All Run keys cleaned and scheduled task validation completed."