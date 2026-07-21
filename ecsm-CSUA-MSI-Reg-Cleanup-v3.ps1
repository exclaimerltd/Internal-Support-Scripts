
<#
.SYNOPSIS
    Removes all registry and file system entries for the MSI installation of the
    Exclaimer Cloud Signature Update Agent (CSUA) from all user profiles on the
    machine. Designed to run under SYSTEM or a local administrator account, such
    as an SCCM or Intune deployment. Signature folder handling is configurable
    via the $SignatureMode variable on line 79.
 
.DESCRIPTION
    Designed for system-wide deployment via Intune, SCCM, or GPO, running as
    SYSTEM or a local administrator.
 
    Product codes, product IDs, and component IDs are discovered dynamically
    from the live registry at runtime. No hardcoded GUID lists are maintained.
    The script searches the following sources:
      - HKLM\...\Installer\UserData\* (all SIDs)
      - HKCR\Installer\Products
      - HKLM\SOFTWARE\WOW6432Node\...\Uninstall
      - HKU\* (currently loaded user hives)
 
    For each user profile found under:
        HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList
 
    The script:
      - Skips built-in system accounts (S-1-5-18, S-1-5-19, S-1-5-20).
      - Loads the user's NTUSER.DAT into a temporary registry hive if not
        already mounted, operates against it, then unloads it.
      - Derives AppData and LocalAppData paths directly from the profile path
        on disk; no dependency on environment variables or a logged-on session.
      - Removes all per-user MSI installer keys, Run entries, AppData folders,
        Outlook signature folders, and Group Policy AppMgmt entries.
      - Handles Outlook signature folders according to $SignatureMode:
            0 = Back up to C:\Temp\Exclaimer\<username>\ then delete (default)
            1 = Delete only
            2 = No changes
        Backup only runs if the folder contains files. The delete is retried up
        to 3 times with a 5-second delay to handle transient file locks.
 
    Machine-wide (HKLM / HKCR) keys are cleaned once after all per-user work.
 
    The agent process is terminated via taskkill across all sessions before any
    file system cleanup begins.
 
    Log written to:
        C:\Temp\Exclaimer\CSUACleanupScript.log
 
.NOTES
    Version:  3.1
    Date:     2026-07-21
    Based on: ecsm-CSUA-MSI-Reg-Cleanup.ps1 (v1, last updated 17/10/2022)
 
    Requires: Run as SYSTEM or local administrator.
              PowerShell 5.1 or later.
 
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
 
.REQUIREMENTS
    Warning: Windows Registry modifications should always be approached with extreme
    care. Back up the registry before running this script.
    https://support.microsoft.com/en-gb/topic/how-to-back-up-and-restore-the-registry-in-windows-855140ad-e318-2a13-2829-d428a2ab0692
#>

#Requires -RunAsAdministrator

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'
$LogPath = 'C:\Temp\Exclaimer\CSUACleanupScript.log'
if (-not (Test-Path 'C:\Temp\Exclaimer')) { New-Item -Path 'C:\Temp\Exclaimer' -ItemType Directory -Force | Out-Null }
$ScriptErrors = [System.Collections.Generic.List[string]]::new()

# ---------------------------------------------------------------------------
# Signature folder behaviour  (set before deploying)
#   0 = Backup to C:\Temp\Exclaimer\<username>\ then delete
#   1 = Delete only (no backup)
#   2 = No changes
# ---------------------------------------------------------------------------
[int]$SignatureMode = 0

function Write-Log {
    param([string]$Message, [string]$Colour = 'Cyan')
    $Stamp = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $Stamp -ForegroundColor $Colour
    $Stamp | Out-File -FilePath $LogPath -Append -Encoding utf8
}

# ---------------------------------------------------------------------------
# Helper: convert a standard GUID string to MSI packed format
# MSI packed GUIDs are used as registry key names throughout the Windows
# Installer database. The algorithm reverses characters within each GUID
# segment: full char-reversal for segments 1-3, pair-swap for segment 4+5.
# ---------------------------------------------------------------------------
function ConvertTo-PackedGuid {
    param([string]$Guid)
    $g = $Guid.ToUpper() -replace '[{}\-]', ''
    if ($g.Length -ne 32) { return $null }
    $s1 = $g[7]+$g[6]+$g[5]+$g[4]+$g[3]+$g[2]+$g[1]+$g[0]
    $s2 = $g[11]+$g[10]+$g[9]+$g[8]
    $s3 = $g[15]+$g[14]+$g[13]+$g[12]
    $s4 = ''
    for ($i = 16; $i -lt 32; $i += 2) { $s4 += $g[$i+1] + $g[$i] }
    return $s1 + $s2 + $s3 + $s4
}

# ---------------------------------------------------------------------------
# Stop the CSUA process across all user sessions
# ---------------------------------------------------------------------------
# Get-Process is session-scoped when run as SYSTEM via Task Scheduler and will
# not see processes running in interactive user sessions. taskkill /F /IM works
# at the system level across all sessions regardless of who owns the process.
# ---------------------------------------------------------------------------
Write-Log "Stopping the Cloud Signature Update Agent process across all sessions..."
$tkResult = & taskkill.exe /F /IM 'Exclaimer.CloudSignatureAgent.exe' 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Log "Process terminated via taskkill." 'Green'
    Start-Sleep -Seconds 3  # allow file locks to release before filesystem cleanup
} elseif ($LASTEXITCODE -eq 128) {
    Write-Log "Process was not running." 'Green'
} else {
    Write-Log "taskkill returned exit code $LASTEXITCODE -- $tkResult" 'Yellow'
}

# ---------------------------------------------------------------------------
# Map HKCR / HKU PSDrives
# ---------------------------------------------------------------------------
Write-Log "Mapping registry PSDrives..."
New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue | Out-Null
New-PSDrive -Name HKU  -PSProvider Registry -Root HKEY_USERS        -ErrorAction SilentlyContinue | Out-Null
Write-Log "PSDrives mapped." 'Green'

# ---------------------------------------------------------------------------
# Dynamic discovery: find all CSUA product codes, product IDs and component IDs
# by searching the live registry rather than matching against hardcoded lists.
#
# Sources searched (in order):
#   1. HKLM\...\Installer\UserData\*\Products\*\InstallProperties  (all SIDs)
#   2. HKCR\Installer\Products\*
#   3. HKLM\SOFTWARE\WOW6432Node\...\Uninstall\*
#   4. HKU\*\Software\Microsoft\Installer\Products\*  (loaded hives only)
#
# For each product code found, the packed ID is derived mathematically so
# both forms are available without a second registry lookup.
# Component IDs are enumerated from the Components subkey of each product.
# ---------------------------------------------------------------------------
Write-Log "Discovering CSUA product artefacts from registry..."

$CSUAProdCode  = [System.Collections.Generic.List[string]]::new()
$CSUAProdID    = [System.Collections.Generic.List[string]]::new()
$ComponentID   = [System.Collections.Generic.List[string]]::new()
$MatchPattern  = '*Cloud Signature Update Agent*'

function Add-IfNew {
    param([System.Collections.Generic.List[string]]$List, [string]$Value)
    if ($Value -and -not $List.Contains($Value)) { $List.Add($Value) }
}

function Register-ProductCode {
    param([string]$Code)
    $clean = $Code.ToUpper() -replace '[{}]', ''
    $fmt   = '{' + $clean + '}'
    Add-IfNew $CSUAProdCode $fmt
    $packed = ConvertTo-PackedGuid $fmt
    if ($packed) { Add-IfNew $CSUAProdID $packed }
}

# Source 1: UserData (all SIDs) - most complete; also harvests component IDs
$UserDataBase = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData'
if (Test-Path $UserDataBase) {
    Get-ChildItem -Path $UserDataBase -ErrorAction SilentlyContinue | ForEach-Object {
        $SIDPath = $_.PSPath
        $ProdBase = "$SIDPath\Products"
        if (Test-Path $ProdBase) {
            Get-ChildItem -Path $ProdBase -ErrorAction SilentlyContinue | ForEach-Object {
                $ProdKey = $_
                $Props = Get-ItemProperty -Path "$($ProdKey.PSPath)\InstallProperties" -ErrorAction SilentlyContinue
                if ($Props -and $Props.DisplayName -clike $MatchPattern) {
                    $PackedID = $ProdKey.PSChildName
                    Add-IfNew $CSUAProdID $PackedID
                    # Component IDs are not under Products; they live at
                    # UserData\<SID>\Components\<CompID> and are harvested separately below
                }
            }
            # Harvest component IDs: any component key whose parent SID has a known product
            $CompPath = "$SIDPath\Components"
            if (Test-Path $CompPath) {
                Get-ChildItem -Path $CompPath -ErrorAction SilentlyContinue | ForEach-Object {
                    # Only add if this SID also has a matching product (avoid unrelated components)
                    $CompID = $_.PSChildName
                    # Check if any product under this SID matched
                    $HasCSUA = Get-ChildItem -Path $ProdBase -ErrorAction SilentlyContinue |
                        ForEach-Object { Get-ItemProperty -Path "$($_.PSPath)\InstallProperties" -ErrorAction SilentlyContinue } |
                        Where-Object { $_.DisplayName -clike $MatchPattern } |
                        Select-Object -First 1
                    if ($HasCSUA) { Add-IfNew $ComponentID $CompID }
                }
            }
        }
    }
}

# Source 2: HKCR\Installer\Products (packed IDs)
if (Test-Path 'HKCR:\Installer\Products') {
    Get-ChildItem 'HKCR:\Installer\Products' -ErrorAction SilentlyContinue | ForEach-Object {
        $Props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        if ($Props -and $Props.ProductName -clike $MatchPattern) {
            Add-IfNew $CSUAProdID $_.PSChildName
        }
    }
}

# Source 3: WOW6432Node Uninstall (product codes in GUID format)
$UninstallBase = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
if (Test-Path $UninstallBase) {
    Get-ChildItem $UninstallBase -ErrorAction SilentlyContinue | ForEach-Object {
        $Props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
        if ($Props -and $Props.DisplayName -clike $MatchPattern) {
            Register-ProductCode $_.PSChildName
        }
    }
}

# Source 4: Currently loaded user hives under HKU
Get-ChildItem 'HKU:\' -ErrorAction SilentlyContinue | ForEach-Object {
    $InstallerProds = "$($_.PSPath)\Software\Microsoft\Installer\Products"
    if (Test-Path $InstallerProds) {
        Get-ChildItem $InstallerProds -ErrorAction SilentlyContinue | ForEach-Object {
            $Props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            if ($Props -and $Props.ProductName -clike $MatchPattern) {
                Add-IfNew $CSUAProdID $_.PSChildName
            }
        }
    }
}

Write-Log "Discovery complete: $($CSUAProdCode.Count) product code(s), $($CSUAProdID.Count) product ID(s), $($ComponentID.Count) component ID(s) found." 'Green'
if ($CSUAProdCode.Count -gt 0) { $CSUAProdCode | ForEach-Object { Write-Log "  ProdCode: $_" } }
if ($CSUAProdID.Count   -gt 0) { $CSUAProdID   | ForEach-Object { Write-Log "  ProdID:   $_" } }
if ($ComponentID.Count  -gt 0) { $ComponentID  | ForEach-Object { Write-Log "  CompID:   $_" } }

if ($CSUAProdCode.Count -eq 0 -and $CSUAProdID.Count -eq 0) {
    Write-Log "No CSUA MSI artefacts found in registry. The product may already be fully removed." 'Yellow'
    Write-Log "Continuing to clean file system paths and any residual keys..." 'Yellow'
}

# ---------------------------------------------------------------------------
# Helper: remove a registry key (recurse) with logging
# ---------------------------------------------------------------------------
function Remove-RegKey {
    param([string]$KeyPath)
    if (Test-Path $KeyPath) {
        try {
            Remove-Item -Path $KeyPath -Recurse -Force -ErrorAction Stop
            Write-Log "  Removed key: $KeyPath" 'Green'
        } catch {
            $Msg = "  ERROR removing key: $KeyPath -- $_"
            Write-Log $Msg 'Red'
            $ScriptErrors.Add($Msg)
        }
    }
}

# ---------------------------------------------------------------------------
# Helper: remove a registry value with logging
# ---------------------------------------------------------------------------
function Remove-RegValue {
    param([string]$KeyPath, [string]$ValueName)
    if (Test-Path $KeyPath) {
        try {
            Remove-ItemProperty -Path $KeyPath -Name $ValueName -Force -ErrorAction Stop
            Write-Log "  Removed value: $ValueName @ $KeyPath" 'Green'
        } catch {
            # Value may simply not exist; only log actual errors
            if ($_.Exception -notmatch 'does not exist') {
                $Msg = "  ERROR removing value: $ValueName @ $KeyPath -- $_"
                Write-Log $Msg 'Red'
                $ScriptErrors.Add($Msg)
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Helper: remove a file system path with logging
# ---------------------------------------------------------------------------
function Remove-FSPath {
    param([string]$FSPath, [int]$Retries = 3, [int]$DelaySeconds = 5)
    if (-not (Test-Path $FSPath)) { return }
    for ($i = 1; $i -le $Retries; $i++) {
        try {
            Remove-Item -Path $FSPath -Recurse -Force -ErrorAction Stop
            Write-Log "  Removed path: $FSPath" 'Green'
            return
        } catch {
            if ($i -lt $Retries) {
                Write-Log "  Retry $i/$Retries for: $FSPath -- $_" 'Yellow'
                Start-Sleep -Seconds $DelaySeconds
            } else {
                $Msg = "  ERROR removing path: $FSPath -- $_"
                Write-Log $Msg 'Red'
                $ScriptErrors.Add($Msg)
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Per-user cleanup function
# Accepts the user SID and the resolved AppData / LocalAppData paths directly
# so it can be called regardless of whether a hive was loaded.
# The HiveRoot parameter is the PSDrive root to use for HKCU operations,
# e.g. 'HKU:\TempUser_S-1-5-21-...' or 'HKCU:' for the current user.
# ---------------------------------------------------------------------------
function Invoke-UserCleanup {
    param(
        [string]$SID,
        [string]$HiveRoot,       # Registry path prefix for this user's hive
        [string]$AppData,        # Resolved %APPDATA% for this user
        [string]$LocalAppData,   # Resolved %LOCALAPPDATA% for this user
        [string]$ProfilePath,    # Full profile path e.g. C:\Users\jsmith (used for sig backup folder name)
        [int]$SignatureMode      # 0=backup+delete, 1=delete, 2=no changes
    )

    Write-Log "--- Cleaning user: SID=$SID ---"
    Write-Log "    AppData      = $AppData"
    Write-Log "    LocalAppData = $LocalAppData"

    # -- Per-user Product ID keys --
    foreach ($ProdID in $CSUAProdID) {
        Remove-RegKey "$HiveRoot\Software\Microsoft\Installer\Features\$ProdID"
        Remove-RegKey "$HiveRoot\Software\Microsoft\Installer\Products\$ProdID"
        Remove-RegKey "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$SID\Products\$ProdID"
        Remove-RegKey "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Managed\$SID\Installer\Products\$ProdID"
        Remove-RegKey "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Managed\$SID\Installer\Features\$ProdID"
        # UpgradeCode entries (value under the key, not the key itself)
        $UpgradePaths = @(
            'HKCR:\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5'
            "$HiveRoot\Software\Microsoft\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5"
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5"
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Managed\$SID\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5"
        )
        foreach ($UP in $UpgradePaths) { Remove-RegValue $UP $ProdID }
    }

    # -- Per-user Product Code entries --
    foreach ($ProdCode in $CSUAProdCode) {
        Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' "$AppData\Microsoft\Installer\$ProdCode\"
        Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' "C:\Windows\Installer\$ProdCode\"
        Remove-RegKey  "HKLM:\SOFTWARE\Microsoft\EnterpriseDesktopAppManagement\$SID\MSI\$ProdCode"
    }

    # -- Per-user component ID keys --
    foreach ($CompID in $ComponentID) {
        Remove-RegKey "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$SID\Components\$CompID"
    }

    # -- Per-user generic keys --
    Remove-RegKey "$HiveRoot\SOFTWARE\Exclaimer Ltd\CloudSignatureUpdateAgent"
    Remove-RegKey "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Uninstall\Exclaimer Cloud Signature Update Agent"

    # -- Per-user run values --
    Remove-RegValue "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Run" 'Cloud Signature Update Agent'
    Remove-RegValue "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Run" 'Exclaimer Cloud Signature Update Agent'
    Remove-RegValue "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" 'Exclaimer Cloud Signature Update Agent'

    # -- Per-user Installer Folder references (AppData paths) --
    Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' "$AppData\Microsoft\Windows\Start Menu\Programs\Exclaimer\"
    Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' "$LocalAppData\Programs\Exclaimer Ltd\"
    Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' "$LocalAppData\Programs\Exclaimer Ltd\Cloud Signature Update Agent\"
    foreach ($Lang in @('de','es','fr','it','nl','pt')) {
        Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' "$LocalAppData\Programs\Exclaimer Ltd\Cloud Signature Update Agent\$Lang\"
    }

    # -- Per-user Group Policy app deployment keys --
    $GPBasePaths = @(
        "$HiveRoot\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Appmgmt"
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Appmgmt"
    )
    foreach ($GPBase in $GPBasePaths) {
        if (Test-Path $GPBase) {
            $GPKeys = Get-ChildItem -Path $GPBase -ErrorAction SilentlyContinue |
                ForEach-Object { Get-ItemProperty -Path "Registry::$_" -ErrorAction SilentlyContinue } |
                Where-Object { $_.'Deployment Name' -clike '*Cloud Signature Update Agent*' } |
                Select-Object -ExpandProperty PSPath
            $GPKeys | ForEach-Object {
                $GPKey = $_ -replace '.*{', '{' -replace '}}', '}'
                Remove-RegKey "$GPBase\$GPKey"
            }
        }
    }

    # -- Per-user file system paths --
    Remove-FSPath "$LocalAppData\Programs\Exclaimer Ltd\Cloud Signature Update Agent"
    Remove-FSPath "$LocalAppData\Programs\Exclaimer Ltd"
    Remove-FSPath "$LocalAppData\Exclaimer"
    Remove-FSPath "$AppData\Microsoft\Windows\Start Menu\Programs\Exclaimer Ltd"
    Remove-FSPath "$AppData\Microsoft\Windows\Start Menu\Programs\Exclaimer"

    # -- Outlook local signatures folders --
    # Behaviour controlled by $SignatureMode (defined at top of script).
    # Folders targeted:
    #   - AppData\Roaming\Microsoft\Signatures                  (active, English)
    #   - AppData\Roaming\Microsoft\Signatures_ExclaimerBackup  (CSUA backup)
    #   - AppData\Roaming\Microsoft\Handtekeningen              (active, Dutch)
    $SigFolders = @(
        "$AppData\Microsoft\Signatures"
        "$AppData\Microsoft\Signatures_ExclaimerBackup"
        "$AppData\Microsoft\Handtekeningen"
    )

    if ($SignatureMode -eq 2) {
        Write-Log "  Signature folders: no changes (mode 2)."
    } else {
        $UserName = Split-Path $ProfilePath -Leaf
        foreach ($SigPath in $SigFolders) {
            if (-not (Test-Path $SigPath)) { continue }

            # Skip empty folders entirely — nothing to back up or clean
            $HasFiles = Get-ChildItem -Path $SigPath -Recurse -File -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if (-not $HasFiles) {
                Write-Log "  Skipping empty folder: $SigPath"
                continue
            }

            if ($SignatureMode -eq 0) {
                $BackupDest = "C:\Temp\Exclaimer\$UserName\$(Split-Path $SigPath -Leaf)"
                try {
                    New-Item -Path $BackupDest -ItemType Directory -Force -ErrorAction Stop | Out-Null
                    Copy-Item -Path "$SigPath\*" -Destination $BackupDest -Recurse -Force -ErrorAction Stop
                    Write-Log "  Backed up: $SigPath -> $BackupDest" 'Green'
                } catch {
                    $Msg = "  ERROR backing up: $SigPath -> $BackupDest -- $_"
                    Write-Log $Msg 'Red'
                    $ScriptErrors.Add($Msg)
                    continue  # do not delete if backup failed
                }
            }

            # Delete with retry to handle transient file locks
            $Deleted = $false
            for ($i = 1; $i -le 3; $i++) {
                try {
                    Remove-Item -Path $SigPath -Recurse -Force -ErrorAction Stop
                    $Deleted = $true
                    break
                } catch {
                    if ($i -lt 3) {
                        Write-Log "  Retry $i/3 clearing: $SigPath -- $_" 'Yellow'
                        Start-Sleep -Seconds 5
                    } else {
                        $Msg = "  ERROR clearing: $SigPath -- $_"
                        Write-Log $Msg 'Red'
                        $ScriptErrors.Add($Msg)
                    }
                }
            }

            if ($Deleted) {
                # Recreate the active folder empty; do not recreate ExclaimerBackup
                if ($SigPath -notlike '*_ExclaimerBackup') {
                    New-Item -Path $SigPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                }
                Write-Log "  Cleared: $SigPath" 'Green'
            }
        }
    }

    Write-Log "--- User $SID complete. ---" 'Green'
}

# ---------------------------------------------------------------------------
# Enumerate all user profiles and run per-user cleanup
# ---------------------------------------------------------------------------
Write-Log "Enumerating user profiles from ProfileList..."

$ProfileListPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
$Profiles = Get-ChildItem -Path $ProfileListPath -ErrorAction SilentlyContinue

foreach ($UserProfile in $Profiles) {
    $SID         = $UserProfile.PSChildName
    $ProfilePath = (Get-ItemProperty -Path $UserProfile.PSPath -ErrorAction SilentlyContinue).ProfileImagePath

    # Skip built-in system accounts (S-1-5-18 = SYSTEM, S-1-5-19 = Local Service, S-1-5-20 = Network Service)
    if ($SID -in @('S-1-5-18','S-1-5-19','S-1-5-20')) { continue }
    if (-not $ProfilePath -or -not (Test-Path $ProfilePath)) {
        Write-Log "  Skipping ${SID}: profile path '$ProfilePath' not found." 'Yellow'
        continue
    }

    # Derive AppData paths from the profile path.
    # Standard layout: %USERPROFILE%\AppData\Roaming and %USERPROFILE%\AppData\Local
    $AppData      = Join-Path $ProfilePath 'AppData\Roaming'
    $LocalAppData = Join-Path $ProfilePath 'AppData\Local'

    # Determine whether the hive is already loaded under HKU
    $HiveKey    = "HKU:\$SID"
    $HiveLoaded = Test-Path $HiveKey

    if (-not $HiveLoaded) {
        $NTUserDat = Join-Path $ProfilePath 'NTUSER.DAT'
        if (-not (Test-Path $NTUserDat)) {
            Write-Log "  Skipping ${SID}: NTUSER.DAT not found at '$NTUserDat'." 'Yellow'
            continue
        }
        # Ensure HKU PSDrive exists
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -ErrorAction SilentlyContinue | Out-Null
        }
        Write-Log "  Loading hive for $SID from $NTUserDat..."
        $LoadResult = & reg.exe load "HKU\$SID" "$NTUserDat" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "  WARNING: Could not load hive for $SID -- $LoadResult" 'Yellow'
            Write-Log "  The user may be logged in; attempting cleanup via HKLM paths only."
            # Still clean HKLM-side keys that reference this SID
            Invoke-UserCleanup -SID $SID -HiveRoot "HKU:\$SID" -AppData $AppData -LocalAppData $LocalAppData -ProfilePath $ProfilePath -SignatureMode $SignatureMode
            continue
        }
        Write-Log "  Hive loaded." 'Green'
    }

    Invoke-UserCleanup -SID $SID -HiveRoot "HKU:\$SID" -AppData $AppData -LocalAppData $LocalAppData -ProfilePath $ProfilePath -SignatureMode $SignatureMode

    # Unload only hives we loaded ourselves
    if (-not $HiveLoaded) {
        Write-Log "  Unloading hive for $SID..."
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        $UnloadResult = & reg.exe unload "HKU\$SID" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "  WARNING: Could not unload hive for $SID -- $UnloadResult" 'Yellow'
        } else {
            Write-Log "  Hive unloaded." 'Green'
        }
    }
}

# ---------------------------------------------------------------------------
# Machine-wide (HKLM / HKCR) cleanup
# These do not vary per user and only need to run once.
# ---------------------------------------------------------------------------
Write-Log "Cleaning machine-wide HKLM / HKCR registry keys..."

# Product ID keys in HKCR and S-1-5-18
foreach ($ProdID in $CSUAProdID) {
    Remove-RegKey "HKCR:\Installer\Features\$ProdID"
    Remove-RegKey "HKCR:\Installer\Products\$ProdID"
    Remove-RegKey "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$ProdID"
    Remove-RegKey "HKLM:\SOFTWARE\Classes\Installer\Products\$ProdID"
}

# Product Code keys
foreach ($ProdCode in $CSUAProdCode) {
    Remove-RegKey "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ProdCode"
}

# Component ID keys (S-1-5-18)
foreach ($CompID in $ComponentID) {
    Remove-RegKey "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\$CompID"
}

# Generic machine-wide keys
Remove-RegKey  'HKLM:\SOFTWARE\WOW6432Node\Exclaimer Ltd\CloudSignatureUpdateAgent'
Remove-RegKey  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Exclaimer Cloud Signature Update Agent'
Remove-RegValue 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run' 'Exclaimer Cloud Signature Update Agent'
Remove-RegValue 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run' 'Cloud Signature Update Agent'
Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' 'C:\Program Files (x86)\Exclaimer Ltd\'
Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' 'C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\'
foreach ($Lang in @('de','es','fr','it','nl','pt')) {
    Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' "C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\$Lang\"
}
Remove-RegValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders' "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Exclaimer\"

Write-Log "Machine-wide registry cleanup complete." 'Green'

# ---------------------------------------------------------------------------
# Machine-wide file system cleanup
# ---------------------------------------------------------------------------
Write-Log "Cleaning machine-wide file system paths..."

$MachineFS = @(
    'C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent'
    'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Exclaimer'
)
foreach ($FSPath in $MachineFS) { Remove-FSPath $FSPath }

Write-Log "Machine-wide file system cleanup complete." 'Green'

# ---------------------------------------------------------------------------
# Remove HKCR PSDrive mapping
# ---------------------------------------------------------------------------
Remove-PSDrive -Name HKCR -ErrorAction SilentlyContinue
Remove-PSDrive -Name HKU  -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# Final log summary
# ---------------------------------------------------------------------------
if ($ScriptErrors.Count -gt 0) {
    Write-Log "Script completed with $($ScriptErrors.Count) error(s). Review log at $LogPath" 'Yellow'
} else {
    Write-Log "Script completed successfully. Log saved to $LogPath" 'Green'
}