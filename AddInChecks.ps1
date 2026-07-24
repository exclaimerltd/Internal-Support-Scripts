# <#
# .SYNOPSIS
#     Gathers diagnostics and configuration data relevant to Exclaimer Add-In and signature deployment across Outlook clients.
#
# .DESCRIPTION
#     This script collects diagnostics including Outlook client versions, signature agent installations, WebView2 runtime presence, endpoint connectivity, local signatures, and cloud geolocation.
#     It prompts for a user's email to determine their domain, fetches relevant endpoints, tests connectivity, inspects Outlook installations, and produces an HTML report "AddInChecks.html" in the user’s Downloads folder.
#
# .NOTES
#     Date: 23rd September 2025
#     Version: 4.26.24
#
# .PRODUCTS
#     Exclaimer Signature Management - Microsoft 365
#
# .REQUIREMENTS
#     - PowerShell 5.1+ or PowerShell Core
#     - Internet connectivity
#     - Script must be run on a Windows machine
#     - Access to registry for Outlook configuration and installed apps
#     - Network ability to test endpoints on port 443
#
# .VERSION
#     1.1.5
#         - Collects Windows version details
#         - Collects Outlook installation and version details
#         - Checks for Exclaimer Cloud Add-in presence and version
#         - Detects deployment method (AppSource, Manifest, or User-installed)
#         - Verifies local Exclaimer Agent and WebView2 installation
#         - Tests network connectivity and service geolocation
#         - Inspects local signatures for Exclaimer integration
#         - Retrieves key organization-level Exchange settings affecting add-ins
#
# .INSTRUCTIONS
#     1. Open PowerShell (as Administrator recommended)
#     2. Set execution policy, e.g. `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`
#     3. Navigate to script folder, e.g. `cd c:\temp`
#     4. Execute: `.\AddInChecks.ps1`
# >
#  

function CheckPowerShellVersion {
    Write-Host "Checking PowerShell version..." -ForegroundColor Cyan
    if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) {
        Write-Host "This script must be run in Windows PowerShell 5.x." -ForegroundColor Red
        Write-Host "Please run this script in Windows PowerShell version 5.x." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press Enter to close"
        return $false
    }
    Write-Host "PowerShell version is compatible." -ForegroundColor Green
    return $true
}
# ------------------------------
# Output Setup
# ------------------------------

# Default path: Downloads folder
$Global:FilePath = [System.IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), 'Downloads')
$LogFile = "AddInChecks_$(Get-Date -Format 'HHmmss').html"

# Check if the path exists, if not, use C:\Temp
if (-not (Test-Path -Path $Global:FilePath)) {
    $Global:FilePath = "C:\Temp"
    if (-not (Test-Path -Path $Global:FilePath)) {
        New-Item -Path $Global:FilePath -ItemType Directory -Force | Out-Null
    }
}

# Final full path for the log file 
$DateTimeRun = Get-Date -Format "ddd dd MMMM yyyy, HH:MM 'UTC' K"
$FullLogFilePath = Join-Path $Global:FilePath $LogFile

# Example: write output to the log file
"Log started at $DateTimeRun" | Out-File -FilePath $FullLogFilePath -Encoding UTF8

# Start HTML structure and open <pre> for formatting
@"
<html>
<head>
    <meta charset='UTF-8'>
    <title>Exclaimer Diagnostics Report - Client-Side</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background-color: #f9f9f9; color: #333; padding: 20px; }
        .container { max-width: 1000px; margin: 0 auto; }
        h1 { color: #003366; }
        h2 { color: #2a52be; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 30px; }
        h3 { color: #2a52be; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 30px; }
        code { display:block; margin-top:5px; }
        .section { margin-bottom: 30px; }
        .success { color: green; font-weight: bold; }
        .fail { color: red; font-weight: bold; }
        .warning { color: orange; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        th { background-color: #eee; }
        a { color: #0078D4; text-decoration: none; } a:hover { text-decoration: underline; }
        .info-after-note { color:#0c5460; background-color:#d1ecf1; border:1px solid #bee5eb; border-left:4px solid #0c5460; padding:14px; border-radius:4px; font-weight:600; margin-top:10px; box-shadow:0 2px 4px rgba(0,0,0,0.1); }
        .info-after-error { color:#721c24; background-color:#f8d7da; border:1px solid #f5c6cb; border-left:4px solid #c82333; padding:14px; border-radius:4px; font-weight:600; margin-top:10px; box-shadow:0 2px 4px rgba(0,0,0,0.1); }
        .info-after-warning { color:#856404; background-color:#fff3cd; border:1px solid #ffeeba; border-left:4px solid #ffc107; padding:14px; border-radius:4px; font-weight:600; margin-top:10px; box-shadow:0 2px 4px rgba(0,0,0,0.1); }
        .info-after-success { color:#155724; background-color:#d4edda; border:1px solid #c3e6cb; border-left:4px solid #28a745; padding:14px; border-radius:4px; font-weight:600; margin-top:10px; box-shadow:0 2px 4px rgba(0,0,0,0.1); }
        .side-note { color: #555; font-size: 12px; margin-top: 5px; font-style: italic; }
        code { background-color: #f1f1f1; padding: 2px 4px; border-radius: 4px; font-weight: bold; color: #c7254e; display:inline; }
        .floating-button-error {
            position: fixed;
            top: 20px;
            right: 20px;
            background-color: #d40000;
            color: white;
            border: none;
            padding: 10px 15px;
            border-radius: 5px;
            cursor: pointer;
            z-index: 1000;
            font-size: 14px;
            font-weight: 600;
        }
        .floating-button-warning {
            position: fixed;
            top: 60px;
            right: 20px;
            background-color: #ffc107;
            color: white;
            border: none;
            padding: 10px 15px;
            border-radius: 5px;
            cursor: pointer;
            z-index: 1000;
            font-size: 14px;
            font-weight: 600;
        }
        .floating-button-warning:hover {
            background-color: #e0a800;
        }
        .floating-button-error:hover {
            background-color: #a80000;
        }
    </style>
</head>
<body>
<div class="container">
<h1>Exclaimer Diagnostics Report - Client-Side</h1>
<p><strong>Run Date:</strong> $DateTimeRun</p>
"@ | Set-Content -Path $FullLogFilePath -Encoding UTF8

Clear-Host
Write-Host "           -----------------------------------------------" -ForegroundColor Cyan
Write-Host "           |                 EXCLAIMER                   |" -ForegroundColor Yellow
Write-Host "           |       Diagnostics Script Collection         |" -ForegroundColor Yellow
Write-Host "           -----------------------------------------------" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1

function ConfirmElevationStatus {

    Write-Host "========== Script Permission Check ==========`n" -ForegroundColor Cyan

    # Get current user without domain
    $fullUser = whoami
    $currentUser = ($fullUser -split '\\')[-1]
    Write-Host "Current script runner: $currentUser`n" -ForegroundColor Cyan

    # Get installed ExchangeOnlineManagement module version
    $exchangeModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
                      Sort-Object Version -Descending |
                      Select-Object -First 1
    $exchangeModuleVersion = if ($exchangeModule) { $exchangeModule.Version.ToString() } else { "Not installed" }
    Write-Host "ExchangeOnlineManagement module version: $exchangeModuleVersion`n" -ForegroundColor Cyan

    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($isAdmin) {
        Write-Host "PowerShell is running with Administrator privileges." -ForegroundColor Green
    }
    else {
        Write-Host "PowerShell is NOT running with Administrator privileges." -ForegroundColor Yellow
        Write-Host "Some diagnostic checks will be skipped or provide incomplete results." -ForegroundColor Yellow
        Write-Host "`nRecommended action: Close this window and re-run the script as Administrator.`n" -ForegroundColor Cyan

        $choice = Read-Host "Do you want to continue anyway? (Y/N)"

        if ($choice.ToUpper() -ne "Y") {
            Write-Host "Script execution cancelled by user." -ForegroundColor Red

            # --- HTML Logging (cancelled run) ---
            Add-Content $FullLogFilePath '<div class="section">'
            Add-Content $FullLogFilePath '<h2>🔐 Script Permission Check</h2>'
            Add-Content $FullLogFilePath "<p title=`"The user running the script in PowerShell.`"><strong>User:</strong> $currentUser</p>"
            Add-Content $FullLogFilePath "<p><strong>ExchangeOnlineManagement Version:</strong> $exchangeModuleVersion</p>"
            Add-Content $FullLogFilePath '<p><strong>Status:</strong> Script stopped. User chose not to continue without Administrator privileges.</p>'
            Add-Content $FullLogFilePath '</div>'

            exit
        }

        Write-Host "Continuing without Administrator privileges..." -ForegroundColor Yellow
    }

    # --- HTML Logging (continued run) ---
    Add-Content $FullLogFilePath '<div class="section">'
    Add-Content $FullLogFilePath '<h2>🔐 Script Permission Check</h2>'
    Add-Content $FullLogFilePath '<table>'
    Add-Content $FullLogFilePath '<tr><th>Property</th><th>Value</th></tr>'

    if ($isAdmin) {
        Add-Content $FullLogFilePath "<tr><td><strong>Windows User</strong></td><td title=`"The user that running the script in PowerShell.`">$currentUser</td></tr>"
        Add-Content $FullLogFilePath '<tr><td><strong>Administrator privileges</strong></td><td>Yes</td></tr>'
        Add-Content $FullLogFilePath "<tr><td><strong>ExchangeOnlineManagement Version</strong></td><td title=`"The version of the ExchangeOnlineManagement module available on this machine.`">$exchangeModuleVersion</td></tr>"
        Add-Content $FullLogFilePath '<tr><td><strong>Impact</strong></td><td>All diagnostics can be collected.</td></tr>'
    }
    else {
        Add-Content $FullLogFilePath "<tr><td><strong>Windows User</strong></td><td title=`"The user that running the script in PowerShell.`">$currentUser</td></tr>"
        Add-Content $FullLogFilePath '<tr><td><strong>Administrator privileges</strong></td><td style="color:#F5A627;font-weight:bold;">No</td></tr>'
        Add-Content $FullLogFilePath "<tr><td><strong>ExchangeOnlineManagement Version</strong></td><td title=`"The version of the ExchangeOnlineManagement module available on this machine.`">$exchangeModuleVersion</td></tr>"
        Add-Content $FullLogFilePath '<tr><td><strong>Impact</strong></td><td>Some diagnostics (firewall, networking, system-level logs) were skipped or limited.</td></tr>'
        Add-Content $FullLogFilePath '<tr><td><strong>User decision</strong></td><td>User chose to continue without elevation.</td></tr>'
        Add-Content $FullLogFilePath '<tr><td><strong>Recommendation</strong></td><td>Re-run the script using Run as administrator if requested by support.</td></tr>'
    }

    Add-Content $FullLogFilePath '</table>'
    Add-Content $FullLogFilePath '</div>'

    # Optional: expose status to other functions
    $Global:IsElevatedSession = $isAdmin
}
ConfirmElevationStatus
function Get-ExclaimerUserInput {
    [CmdletBinding()]
    param ()

    while ($true) {
        # Initialize object
        $Global:userInput = [PSCustomObject]@{
            Purpose         = $null
            Email           = $null
            UsersAffected   = $null
            OutlookAffected = $null
            OSVersion       = $null
            OutlookVersion  = $null
            Network         = $null
        }

        # --- 1) Email (validated) ---
        while ($true) {
            $email = Read-Host "`nEnter the affect mailboxes email address (e.g. user@company.com)"
            if ($email -match '^[\w\.\-]+@([\w\-]+\.)+[\w\-]{2,}$') {
                $Global:userInput.Email = $email.Trim()
                break
            } else {
                Write-Host "Invalid email format. Try again." -ForegroundColor Red
            }
        }

        # --- 2) Purpose ---
        # Disabled prompt block without removing it
        if ($false) {
            do {
                Clear-Host
                Write-Host "`nPlease choose an option:" -ForegroundColor Cyan
                Write-Host "  1) Troubleshoot an issue"
                Write-Host "  2) View current configuration"
                $choice = Read-Host "`nEnter choice (1 or 2)"
            } while ($choice -notmatch '^[12]$')
        }

        # Always force purpose 1
        $choice = '1'
        $Global:userInput.Purpose = 'Troubleshooting'

        $Global:userInput.Purpose = if ($choice -eq '1') { 'Troubleshooting' } else { 'Configuration Overview' }

        # --- 3) If troubleshooting, ask follow-ups ---
        if ($Global:userInput.Purpose -eq 'Troubleshooting') {
            # Users affected
            do {
                Clear-Host
                Write-Host "`nHow many users are affected?" -ForegroundColor Cyan
                Write-Host "  1) All users"
                Write-Host "  2) Specify number"
                $uc = Read-Host "`nEnter choice (1 or 2)"
            } while ($uc -notmatch '^[12]$')

            if ($uc -eq '1') {
                $Global:userInput.UsersAffected = 'All Users'
            } else {
                do {
                    $num = Read-Host "Enter the approximate number of affected users (digits only)"
                } while ($num -notmatch '^\d+$')
                $Global:userInput.UsersAffected = [int]$num
            }

            # Outlook versions
            do {
                Clear-Host
                Write-Host "`nWhich Outlook version(s) are affected?" -ForegroundColor Cyan
                Write-Host "  1) Classic Outlook (Windows)"
                Write-Host "  2) New Outlook (Windows)"
                Write-Host "  3) Outlook on Web (OWA)"
                Write-Host "  4) Outlook iOS"
                Write-Host "  5) Outlook Android"
                Write-Host "  6) New Outlook on MacOS"
                Write-Host "  7) Outlook Web (MacOS)"
                Write-Host "  8) Multiple / All"
                $oChoice = Read-Host "`nEnter choice (1-8)"
            } while ($oChoice -notmatch '^[1-8]$')

            switch ($oChoice) {
                1 { $Global:userInput.OutlookAffected = 'Classic Outlook (Windows)' }
                2 { $Global:userInput.OutlookAffected = 'New Outlook (Windows)' }
                3 { $Global:userInput.OutlookAffected = 'Outlook Web (Windows)' }
                4 { $Global:userInput.OutlookAffected = 'Outlook iOS' }
                5 { $Global:userInput.OutlookAffected = 'Outlook Android' }
                6 { $Global:userInput.OutlookAffected = 'New Outlook on MacOS' }
                7 { $Global:userInput.OutlookAffected = 'Outlook Web (MacOS)' }
                8 { $Global:userInput.OutlookAffected = 'Multiple / All' }
            }

            # --- Version capture for non-Windows platforms ---
            Clear-Host
            switch ($oChoice) {
                3 {
                    # OWA — browser name/version
                    Write-Host "`nOWA — Browser details" -ForegroundColor Cyan
                    Write-Host "  Open the affected browser and go to its About page to find the version."
                    Write-Host "  e.g. Chrome: Menu > Help > About Google Chrome"
                    $Global:userInput.OSVersion     = Read-Host "`nBrowser name and version (e.g. Chrome 125.0.6422.112)"
                    $Global:userInput.OutlookVersion = Read-Host "Outlook Web build shown in OWA (top-right ? > About, or press Enter to skip)"
                }
                4 {
                    # iOS
                    Write-Host "`nOutlook iOS — device details" -ForegroundColor Cyan
                    Write-Host "  iOS version:     Settings > General > About > iOS Version"
                    Write-Host "  Outlook version: Settings > General > About (within Outlook app)"
                    $Global:userInput.OSVersion      = Read-Host "`niOS version (e.g. 18.4.1)"
                    $Global:userInput.OutlookVersion = Read-Host "Outlook for iOS version (e.g. 4.2559.0)"
                }
                5 {
                    # Android
                    Write-Host "`nOutlook Android — device details" -ForegroundColor Cyan
                    Write-Host "  Android version: Settings > About phone > Software information > Android version"
                    Write-Host "  Outlook version: Outlook app > Settings > Help & Feedback > About"
                    $Global:userInput.OSVersion      = Read-Host "`nAndroid version (e.g. 14)"
                    $Global:userInput.OutlookVersion = Read-Host "Outlook for Android version (e.g. 4.2559.0)"
                }
                6 {
                    # macOS
                    Write-Host "`nOutlook on macOS — device details" -ForegroundColor Cyan
                    Write-Host "  macOS version:   Apple menu > About This Mac"
                    Write-Host "  Outlook version: Outlook menu > About Microsoft Outlook"
                    $Global:userInput.OSVersion      = Read-Host "`nmacOS version (e.g. Sequoia 15.4.1)"
                    $Global:userInput.OutlookVersion = Read-Host "Outlook for Mac version (e.g. 16.96.2)"
                }
            }

            if ($Global:userInput.OSVersion)      { $Global:userInput.OSVersion      = $Global:userInput.OSVersion.Trim() }
            if ($Global:userInput.OutlookVersion) { $Global:userInput.OutlookVersion = $Global:userInput.OutlookVersion.Trim() }
            if ([string]::IsNullOrWhiteSpace($Global:userInput.OSVersion))      { $Global:userInput.OSVersion      = 'Not provided' }
            if ([string]::IsNullOrWhiteSpace($Global:userInput.OutlookVersion)) { $Global:userInput.OutlookVersion = 'Not provided' }

            # Network scope
            do {
                Clear-Host
                Write-Host "`nWhere does the issue occur?" -ForegroundColor Cyan
                Write-Host "  1) Internal network only"
                Write-Host "  2) External network only"
                Write-Host "  3) Both internal and external"
                $nChoice = Read-Host "`nEnter choice (1-3)"
            } while ($nChoice -notmatch '^[1-3]$')

            switch ($nChoice) {
                1 { $Global:userInput.Network = 'Internal network only' }
                2 { $Global:userInput.Network = 'External network only' }
                3 { $Global:userInput.Network = 'Both internal and external' }
            }
        }

        # --- 4) Show console summary ---
        Clear-Host
        Write-Host ""
        Write-Host "========================================" -ForegroundColor DarkGray
        Write-Host "            Summary captured" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor DarkGray

        Write-Host ("Purpose:          {0}" -f $Global:userInput.Purpose) -ForegroundColor Cyan
        Write-Host ("Email:            {0}" -f $Global:userInput.Email) -ForegroundColor Yellow

        if ($Global:userInput.Purpose -eq 'Troubleshooting') {
            Write-Host ("Users Affected:   {0}" -f $Global:userInput.UsersAffected) -ForegroundColor White
            Write-Host ("Outlook Affected: {0}" -f $Global:userInput.OutlookAffected) -ForegroundColor White
            if ($oChoice -in @('3','4','5','6','7','8')) {
                Write-Host ("OS / Browser:     {0}" -f $Global:userInput.OSVersion) -ForegroundColor White
                Write-Host ("Outlook Version:  {0}" -f $Global:userInput.OutlookVersion) -ForegroundColor White
            }
            Write-Host ("Network Scope:    {0}" -f $Global:userInput.Network) -ForegroundColor White
        }

        Write-Host ""
        do {
            $confirm = Read-Host "Is the information correct? (Y/N) [Y]"
            if ([string]::IsNullOrWhiteSpace($confirm)) { $confirm = 'Y' }
            $confirm = $confirm.Substring(0,1).ToUpper()
        } while ($confirm -notin @('Y','N'))

        if ($confirm -eq 'Y') {
            # Write to HTML log only once, after confirmation
            Add-Content $FullLogFilePath "<div class='section'>"
            Add-Content $FullLogFilePath "<h2>🧾 User Input Summary</h2>"
            Add-Content $FullLogFilePath "<table>"
            Add-Content $FullLogFilePath "<tr><td><strong>Purpose:</strong></td><td>$($Global:userInput.Purpose)</td></tr>"
            Add-Content $FullLogFilePath "<tr><td><strong>Email:</strong></td><td>$($Global:userInput.Email)</td></tr>"

            if ($Global:userInput.Purpose -eq 'Troubleshooting') {
                Add-Content $FullLogFilePath "<tr><td><strong>Users Affected:</strong></td><td>$($Global:userInput.UsersAffected)</td></tr>"
                Add-Content $FullLogFilePath "<tr><td><strong>Outlook Affected:</strong></td><td>$($Global:userInput.OutlookAffected)</td></tr>"
                if ($oChoice -in @('3','4','5','6','7','8')) {
                    Add-Content $FullLogFilePath "<tr><td><strong>OS / Browser Version:</strong></td><td>$($Global:userInput.OSVersion)</td></tr>"
                    Add-Content $FullLogFilePath "<tr><td><strong>Outlook Version:</strong></td><td>$($Global:userInput.OutlookVersion)</td></tr>"
                }
                Add-Content $FullLogFilePath "<tr><td><strong>Network Scope:</strong></td><td>$($Global:userInput.Network)</td></tr>"
            }

            Add-Content $FullLogFilePath "</table></div>"
            return $Global:userInput
        } else {
            Write-Host "`nLet's try again..." -ForegroundColor Yellow
            Start-Sleep -Seconds 1
            Clear-Host
        }
    }
}
function Get-Region {
    # Define log file path (adjust as needed)

    $email = $Global:userInput.Email.ToLower().Trim()

    # Validate email format (extra check just in case)
    if ($email -match '^[\w\.\-]+@([\w\-]+\.)+[\w\-]{2,4}$') {
        Write-Host "`nEmail address entered: $email" -ForegroundColor Green
        Add-Content $FullLogFilePath "<div class='section'><h2>🌐 Domain Name Checks</h2><p><strong>Email address:</strong> $email</p></div>"
    }
    else {
        Write-Host "Invalid email format detected. This should not happen because of prior validation." -ForegroundColor Red
        return
    }

    # Extract domain
    $domain = $email.Split("@")[1]

    # Host to test connectivity
    $hostToTest = "outlookclient.exclaimer.net"

    Write-Host "`nChecking connectivity to $hostToTest..." -ForegroundColor Cyan
    Add-Content $FullLogFilePath "<p>Checking connectivity to <strong>$hostToTest</strong>...</p>"

    # Test TCP connectivity on port 443
    $connectionTest = Test-NetConnection -ComputerName $hostToTest -Port 443 -InformationLevel Quiet

    if (-not $connectionTest) {
        Write-Host "Unable to connect to $hostToTest on port 443. Please check network connectivity." -ForegroundColor Red
        Add-Content $FullLogFilePath "<p class='fail'>❌ Unable to connect to $hostToTest on port 443.</p>"
        Add-Content $FullLogFilePath "<p class='info-after-error'>ℹ️ Check your Internet connection or network blocking (<a href='https://support.exclaimer.com/hc/en-gb/articles/7317900965149-Ports-and-URLs-used-by-the-Exclaimer-Outlook-Add-In' target='_blank'>see article</a>).</p>"

        $global:OutlookSignaturesEndpoint = $hostToTest
        return
    }

    Write-Host "Connectivity OK." -ForegroundColor Green
    Add-Content $FullLogFilePath "<p class='pass'>✅ Connectivity OK.</p>"

    Write-Host "Proceeding to fetch data for domain: '$domain'" -ForegroundColor Yellow
    Add-Content $FullLogFilePath "<p>Fetching cloud geolocation data for domain: <strong>$domain</strong></p>"

    $url = "https://$hostToTest/cloudgeolocation/$domain"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop

        if ($response.PSObject.Properties.Name -contains 'OutlookSignaturesEndpoint' -and
            -not [string]::IsNullOrEmpty($response.OutlookSignaturesEndpoint)) {

            $endpoint = $response.OutlookSignaturesEndpoint

            # Clean endpoint URL
            if ($endpoint.StartsWith("https://")) { $endpoint = $endpoint.Substring(8) }
            if ($endpoint.EndsWith("/")) { $endpoint = $endpoint.TrimEnd('/') }

            if (-not [string]::IsNullOrEmpty($endpoint)) {
                $global:OutlookSignaturesEndpoint = $endpoint
                Write-Host "`nOutlookSignaturesEndpoint found: '$endpoint'" -ForegroundColor Green
                Add-Content $FullLogFilePath "<p class='pass'>✅ OutlookSignaturesEndpoint found: <strong>$endpoint</strong></p>"
            }
            else {
                Write-Host "'OutlookSignaturesEndpoint' is empty after cleanup." -ForegroundColor Yellow
                Add-Content $FullLogFilePath "<p class='warn'>⚠️ 'OutlookSignaturesEndpoint' is empty after cleanup.</p>"
                Add-Content $FullLogFilePath "<p class='warn'>This may happen if your Exclaimer subscription is not synced with your Microsoft 365 tenant.</p>"
                $global:OutlookSignaturesEndpoint = $hostToTest
            }
        }
        else {
            Write-Host "'OutlookSignaturesEndpoint' not found in response." -ForegroundColor Yellow
            Add-Content $FullLogFilePath "<p class='warn'>⚠️ 'OutlookSignaturesEndpoint' not found in response.</p>"
            $global:OutlookSignaturesEndpoint = $hostToTest
        }
    }
    catch {
        Write-Host "No data found for domain '$domain'." -ForegroundColor Red
        Add-Content $FullLogFilePath "<p class='fail'>❌ No data found for domain '$domain'.</p>"
        Add-Content $FullLogFilePath "<p class='info-after-error'>ℹ️ This may happen if your Exclaimer subscription is not synced with your Microsoft 365 tenant (` +
            `<a href='https://support.exclaimer.com/hc/en-gb/articles/6389214769565-Synchronize-user-data' target='_blank'>see article</a>).</p>"
        $global:OutlookSignaturesEndpoint = $hostToTest
    }
}

# -------------------------------
# Endpoint Connectivity Tests
# -------------------------------
function CheckEndpoints {
    Write-Host "`n========== Endpoint Connectivity Tests ==========" -ForegroundColor Cyan
    Add-Content $FullLogFilePath "<div class='section'>"
    Add-Content $FullLogFilePath "<h2>📡 Endpoint Connectivity Tests</h2>"

    $endpoints = @(
        $global:OutlookSignaturesEndpoint,
        "login.microsoftonline.com",
        "secure.aadcdn.microsoftonline-p.com",
        "appsforoffice.microsoft.com",
        "outlookclient.exclaimer.net",
        "static2.sharepointonline.com",
        "pro.fontawesome.com"
    )

    $results = @()

    foreach ($endpoint in $endpoints) {
        $TimeTaken = Measure-Command {
            $Global:netTestResult = Test-NetConnection -ComputerName $endpoint -Port 443 -InformationLevel Quiet
        }

        $status = if ($Global:netTestResult) { "Success" } else { "Failed" }

        $results += [PSCustomObject]@{
            "Endpoint" = $endpoint
            "Status"   = $status
            "Time (s)" = "{0:N3}" -f $TimeTaken.TotalSeconds
        }
    }

    # Console output
    $results | Format-Table -AutoSize

    # HTML table output
    Add-Content $FullLogFilePath "<table>"
    Add-Content $FullLogFilePath "<thead><tr><th>Endpoint</th><th>Status</th><th>Time (s)</th></tr></thead>"
    Add-Content $FullLogFilePath "<tbody>"

    foreach ($r in $results) {
        $statusClass = if ($r.Status -eq "Success") { "success" } else { "fail" }
        $row = "<tr><td>$($r.Endpoint)</td><td class='$statusClass'>$($r.Status)</td><td>$($r.'Time (s)')</td></tr>"
        Add-Content $FullLogFilePath $row
    }

    Add-Content $FullLogFilePath "</tbody></table>"

    # Check for any failures and append warning message
    if ($results.Status -contains "Failed") {
        Add-Content $FullLogFilePath "<p class='warning'>❗ One or more endpoints failed to respond. Please check your internet connection or firewall settings and try again.</p>"
        Add-Content $FullLogFilePath "<p class='info-after-error'>ℹ️ Check your Internet connection, your network could also be blocking the connection (<a href='https://support.exclaimer.com/hc/en-gb/articles/7317900965149-Ports-and-URLs-used-by-the-Exclaimer-Outlook-Add-In' target='_blank'>see article</a>).</p>"
    }

    Add-Content $FullLogFilePath "</div>"
}
$userData = Get-ExclaimerUserInput
Get-Region -userInput $userData
CheckEndpoints
# -------------------------------
# Getting Windows version
# -------------------------------
function GetWindowsVersion {

    Write-Host "`n========== Microsoft Windows Version ==========" -ForegroundColor Cyan

    # --- HTML Section Header ---
    Add-Content $FullLogFilePath '<div class="section">'
    Add-Content $FullLogFilePath '<h2>💻 Microsoft Windows Version</h2>'
    Add-Content $FullLogFilePath '<table>'
    Add-Content $FullLogFilePath '<tr><th>Property</th><th>Value</th></tr>'

    # --- Collect OS Info ---
    $os      = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Caption, Version, BuildNumber, ProductType
    $caption = $os.Caption
    $version = $os.Version
    $build   = [int]$os.BuildNumber
    $isServer = $os.ProductType -ne 1

    # --- Support matrices ---
    $clientBuilds = @(
        [PSCustomObject]@{MinBuild = 28000; Note = 'Windows 11 26H2 build (Insider/preview).'},
        [PSCustomObject]@{MinBuild = 26200; Note = 'Windows 11 25H2 build (latest general availability).'},
        [PSCustomObject]@{MinBuild = 26100; Note = 'Windows 11 24H2 build (supported).'},
        [PSCustomObject]@{MinBuild = 22631; Note = 'Windows 11 23H2 build (Enterprise/extended until late 2026).'},
        [PSCustomObject]@{MinBuild = 19045; Note = 'Windows 10 22H2 build (supported until October 2025/extended).'}
    )

    $serverBuilds = @(
        [PSCustomObject]@{ MinBuild = 26100; CaptionMatch = 'Server 2025'; Note = 'Microsoft Windows Server 2025 Standard. Supported when Outlook is used in an RDS user session.' },
        [PSCustomObject]@{ MinBuild = 20348; CaptionMatch = 'Server 2022'; Note = 'Microsoft Windows Server 2022 Standard. Supported when Outlook is used in an RDS user session.' },
        [PSCustomObject]@{ MinBuild = 17763; CaptionMatch = 'Server 2019'; Note = 'Microsoft Windows Server 2019 Standard. Supported when Outlook is used in an RDS user session.' },
        [PSCustomObject]@{ MinBuild = 14393; CaptionMatch = 'Server 2016'; Note = 'Microsoft Windows Server 2016 Standard. Supported when Outlook is used in an RDS user session.' }
    )

    # --- Default state ---
    $supportStatus = '❌ Unsupported or legacy Windows version.'
    $supportNote   = 'Consider upgrading to a supported Windows client or server release.'

    # --- Determine support ---
    if ($isServer) {

        foreach ($entry in $serverBuilds) {
            if ($build -ge $entry.MinBuild -and $caption -like "*$($entry.CaptionMatch)*") {
                $supportStatus = '✅ Supported'
                $supportNote   = $entry.Note
                break
            }
        }

    }
    else {

        foreach ($entry in $clientBuilds) {
            if ($build -ge $entry.MinBuild) {
                $supportStatus = '✅ Supported'
                $supportNote   = $entry.Note
                break
            }
        }

    }

    # --- Console Output ---
    Write-Host "Windows Version: $caption ($version)" -ForegroundColor White
    Write-Host "Build Number:    $build" -ForegroundColor White
    Write-Host "OS Type:         $(if ($isServer) { 'Server' } else { 'Desktop' })" -ForegroundColor White
    Write-Host "Support Status:  $supportStatus" -ForegroundColor Yellow
    Write-Host "Note:            $supportNote" -ForegroundColor DarkGray

    # --- HTML Logging ---
    Add-Content $FullLogFilePath ("<tr><td><strong>Windows Version</strong></td><td>{0} ({1})</td></tr>" -f $caption, $version)
    Add-Content $FullLogFilePath ("<tr><td><strong>Build Number</strong></td><td>{0}</td></tr>" -f $build)
    Add-Content $FullLogFilePath ("<tr><td><strong>OS Type</strong></td><td>{0}</td></tr>" -f $(if ($isServer) { 'Server' } else { 'Desktop' }))
    Add-Content $FullLogFilePath ("<tr><td><strong>Support Status</strong></td><td>{0}</td></tr>" -f $supportStatus)
    Add-Content $FullLogFilePath ("<tr><td><strong>Notes</strong></td><td>{0}</td></tr>" -f $supportNote)

    Add-Content $FullLogFilePath '</table>'
    Add-Content $FullLogFilePath '</div>'
}
# Only run Windows-specific checks if affected platform is Classic or New Outlook
if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)','Outlook Web (Windows)','Multiple / All')) {
    GetWindowsVersion
}

function GetWindowsNetworkDetails {
    Write-Host "`n========== Network Connection Details ==========" -ForegroundColor Cyan

    # --- HTML Section Header ---
    Add-Content $FullLogFilePath '<div class="section">'
    Add-Content $FullLogFilePath '<h2>🌐 Network Connection Details</h2>'
    Add-Content $FullLogFilePath '<table>'
    # --- HTML Table Header ---
    Add-Content $FullLogFilePath '<tr><th>Interface</th><th>Network Name</th><th>Category</th><th>IPv4 Connectivity</th><th>IPv6 Connectivity</th><th title="Interface priority, lower is higher">Interface Metric</th></tr>'

    # --- Collect Network Profiles ---
    $netProfiles = Get-NetConnectionProfile

    if (-not $netProfiles) {
        Write-Host "No active network connections found." -ForegroundColor Yellow
        Add-Content $FullLogFilePath '<tr><td colspan="6">No active network connections detected.</td></tr>'
    }
    else {
        foreach ($netProfile in $netProfiles) {
            $interfaceAlias = $netProfile.InterfaceAlias
            $networkName    = if ($netProfile.Name) { $netProfile.Name } else { 'N/A' }
            $category       = $netProfile.NetworkCategory
            $ipv4           = $netProfile.IPv4Connectivity
            $ipv6           = $netProfile.IPv6Connectivity

            # --- Get InterfaceMetric (IPv4) ---
            $metricObj = Get-NetIPInterface |
                        Where-Object { $_.InterfaceAlias -eq $interfaceAlias -and $_.AddressFamily -eq 'IPv4' }
            $interfaceMetric = if ($metricObj) { $metricObj.InterfaceMetric } else { 'N/A' }

            # --- Console Output ---
            Write-Host "Interface:        $interfaceAlias" -ForegroundColor White
            Write-Host "Network Name:     $networkName" -ForegroundColor White
            Write-Host "Category:         $category" -ForegroundColor White
            Write-Host "IPv4 Connectivity $ipv4" -ForegroundColor DarkGray
            Write-Host "IPv6 Connectivity $ipv6" -ForegroundColor DarkGray
            Write-Host "Interface Metric: $interfaceMetric`n" -ForegroundColor Yellow

            # --- HTML Logging ---
            Add-Content $FullLogFilePath (
                "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td title=`"Interface priority, lower is higher`">{5}</td></tr>" -f `
                $interfaceAlias, $networkName, $category, $ipv4, $ipv6, $interfaceMetric
            )
        }
    }

    Add-Content $FullLogFilePath '</table>'
    Add-Content $FullLogFilePath '</div>'

    # ==========================================================
    # WinHTTP Proxy Configuration Check
    # ==========================================================

    Write-Host "`n========== WinHTTP Proxy Configuration (Info) ==========" -ForegroundColor Cyan

    $winHttpOutput = netsh winhttp show proxy

    $proxyConfigured = $true
    if ($winHttpOutput -match 'Direct access') {
        $proxyConfigured = $false
    }

    # --- Console Output ---
    $winHttpOutput | ForEach-Object {
        Write-Host $_ -ForegroundColor White
    }

    # --- Interpret State (Informational) ---
    if ($proxyConfigured) {
        $proxyState = ($winHttpOutput | Where-Object { $_ -match 'Proxy Server' }).Trim()
        $note = 'A WinHTTP proxy is configured. If Classic Outlook cloud apps encounter issues, verify that the proxy supports non-interactive authentication, ensure Microsoft 365 endpoints are reachable, compare browser PAC/auto-detect settings with WinHTTP, and check the proxy logs for any blocked or rewritten traffic.'
        $icon = 'ℹ️'
    }
    else {
        $proxyState = 'Direct access (no WinHTTP proxy configured)'
        $note = 'Direct access is normal when no proxy is configured or when using a transparent VPN such as Cloudflare WARP.'
        $icon = 'ℹ️'
    }

    # --- HTML Logging ---
    Add-Content $FullLogFilePath '<div class="section">'
    Add-Content $FullLogFilePath '<h2>🖥️ WinHTTP Proxy Configuration (Informational)</h2>'
    Add-Content $FullLogFilePath '<table>'
    Add-Content $FullLogFilePath '<tr><th>Status</th><th>Configuration</th><th>Notes</th></tr>'
    Add-Content $FullLogFilePath (
        "<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>" -f `
        $icon, $proxyState, $note
    )
    Add-Content $FullLogFilePath '</table>'
    Add-Content $FullLogFilePath '</div>'
}
if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)','Outlook Web (Windows)','Multiple / All')) {
    GetWindowsNetworkDetails
}

function InspectOutlookConfiguration {
    # -------------------------------
    # Local Function Scope Variables
    # -------------------------------

    $GuidToChannelMap = @{
        "64256afe-f5d9-4f86-8936-8840a6a4f5be" = "Monthly Enterprise Channel"
        "3a0e5e62-6ac6-4a3a-9864-e3b14b6e06b9" = "Semi-Annual Enterprise Channel (Broad)"
        "4a88f291-7f7b-4cbb-90bb-c6b1d75c2911" = "Semi-Annual Enterprise Channel (Targeted - Deprecated)"
        "32a6a3b2-c537-4cdb-9ec3-520e49f103f8" = "Beta Channel (Insider Fast)"
        "5d9b2b78-f6b9-4f8c-9f87-f9aa5f8d35b7" = "Current Channel (Preview)"
        "67c4f9b4-2ab4-4af9-a05d-c8b3a4f0c6a6" = "Monthly Enterprise Preview"
        "2479eec6-ec8d-44e4-9b7a-5a7a82db9821" = "Current Channel (Deprecated GUID)"
    }

    $RegistryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\ProductVersion"
    )

    # Compatibility Requirements Table
    # https://support.exclaimer.com/hc/en-gb/articles/4406058988945
    $minimumSupportedBuilds = @{
        "Subscription (Microsoft 365)"        = "18025.20000"
        "Retail (Perpetual or Subscription)"  = "18429.20132"
        "Volume Licensed (Perpetual)"         = "17932.20222"
    }

    # -------------------------------
    # Nested Functions
    # -------------------------------

    function Get-FriendlyUpdateChannel {
        param ($Properties)

        if ($Properties.PSObject.Properties.Name -contains "UpdateChannel") {
            $url = $Properties.UpdateChannel
            $channelId = ($url -split "/")[-1].ToLower()

            if ($GuidToChannelMap.ContainsKey($channelId)) {
                $channelName = $GuidToChannelMap[$channelId]
            } elseif ($channelId) {
                $channelName = $channelId
            } else {
                $channelName = "Unknown"
            }

            Write-Host "Update Channel: $channelName"
        }
        elseif ($Properties.PSObject.Properties.Name -contains "UpdateBranch") {
            Write-Host "Update Channel (UpdateBranch): $($Properties.UpdateBranch)"
        }
        elseif ($Properties.PSObject.Properties.Name -contains "CDNBaseUrl") {
            $cdnUrl = $Properties.CDNBaseUrl
            $channel = ($cdnUrl -split "/")[-1]
            Write-Host "Update Channel (CDNBaseUrl): $channel"
        }
    }

    function Get-OfficeConfiguration {
        foreach ($path in $RegistryPaths) {
            if (Test-Path $path) {
                $props = Get-ItemProperty -Path $path
                foreach ($prop in $props.PSObject.Properties) {
                    if ($prop.Name -match "Version|ProductReleaseIds" -and $prop.Name -ne "ClientXnoneVersion") {
                        Write-Host "$($prop.Name): $($prop.Value)"
                    }
                }
                if ($path -like "*ClickToRun*") {
                    Get-FriendlyUpdateChannel -Properties $props
                }
            }
        }
    }

    function Get-OfficeLicenseType {
        $c2rKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        if (Test-Path $c2rKey) {
            $cfg = Get-ItemProperty -Path $c2rKey -ErrorAction SilentlyContinue
            if ($null -ne $cfg.ProductReleaseIds) {
                $pr = $cfg.ProductReleaseIds
                if ($pr -match "O365|M365|Microsoft365|ProPlusRetail|BusinessRetail") {
                    return "Subscription (Microsoft 365)"
                }
                elseif ($pr -match "Volume" -and $pr -match "20(19|21|24)") {
                    return "Volume Licensed (Perpetual)"
                }
                elseif ($pr -match "Retail") {
                    # Assume Retail is Perpetual unless it matches subscription strings above
                    return "Retail (Perpetual or Subscription)"
                }
                else {
                    return "Unknown / Mixed: $pr"
                }

            } else {
                return "ProductReleaseIds not found"
            }
        } else {
            return "ClickToRun config key not found"
        }
    }

    function Get-NewOutlookPackage {
        return Get-AppxPackage -Name Microsoft.OutlookForWindows -ErrorAction SilentlyContinue
    }

    function IsNewOutlookAppInstalled {
        return [bool](Get-NewOutlookPackage)
    }

    function Get-NewOutlookVersion {
        $package = Get-NewOutlookPackage
        if ($package) { return $package.Version }
        return $null
    }

    function IsNewOutlookEnabled {
        $registryPaths = @(
            "HKCU:\Software\Microsoft\Office\Outlook\Settings",
            "HKCU:\Software\Microsoft\Office\Outlook\Profiles",
            "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General",
            "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
        )

        foreach ($path in $registryPaths) {
            if (Test-Path $path) {
                try {
                    $props = Get-ItemProperty -Path $path
                    if ($props.PSObject.Properties.Name -contains "UseNewOutlook" -and $props.UseNewOutlook -eq 1) {
                        return $true
                    }
                    foreach ($prop in $props.PSObject.Properties) {
                        if ($prop.Name -like "*NewExperienceEnabled*" -and $prop.Value -eq 1) {
                            return $true
                        }
                    }
                } catch {
                    continue
                }
            }
        }
        return $false
    }

    function IsClassicOutlookInstalled {
        $classicPaths = @(
            "${env:ProgramFiles}\Microsoft Office\root\Office16\Outlook.exe",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\Outlook.exe"
        )

        foreach ($path in $classicPaths) {
            if (Test-Path $path) {
                return $true
            }
        }

        $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
        return (Test-Path $registryPath)
    }

    # -------------------------------
    # Main Execution
    # -------------------------------

    Write-Host "`n========== Outlook Installations Detected ==========" -ForegroundColor Cyan
    Add-Content $FullLogFilePath "<div class='section'>"
    Add-Content $FullLogFilePath "<h2>✉️ Mail Client Checks</h2>"

    $classicInstalled     = IsClassicOutlookInstalled
    $newOutlookInstalled  = IsNewOutlookAppInstalled
    $newOutlookEnabled    = IsNewOutlookEnabled

    # Build Outlook installation summary table
    $installedSummary = "<table>"
    $installedSummary += "<tr><th>Client</th><th>Status</th></tr>"

    if ($classicInstalled -and $newOutlookInstalled) {
        Write-Host "Both Classic Outlook and New Outlook are installed."
        $installedSummary += "<tr><td>Classic + New Outlook</td><td><span class='success'>Installed</span></td></tr>"
    } elseif ($classicInstalled) {
        Write-Host "Only Classic Outlook is installed."
        $installedSummary += "<tr><td>Classic Outlook</td><td><span class='success'>Installed</span></td></tr>"
    } elseif ($newOutlookInstalled) {
        Write-Host "Only New Outlook is installed."
        $installedSummary += "<tr><td>New Outlook</td><td><span class='success'>Installed</span></td></tr>"
    } else {
        Write-Host "No Outlook installation detected."
        $installedSummary += "<tr><td>Classic / New Outlook</td><td><span class='fail'>Not Installed</span></td></tr>"
    }

    # Add toggle status for New Outlook
    if ($newOutlookInstalled) {
        if ($newOutlookEnabled) {
            Write-Host "New Outlook is installed, and the toggle is ON (New Outlook is Default)." -ForegroundColor Yellow
            $installedSummary += "<tr><td>New Outlook Toggle</td><td><span>ON (New Outlook is Default)</span></td></tr>"
        } else {
            Write-Host "New Outlook is installed, but the toggle is OFF (Classic Outlook is Default)." -ForegroundColor Yellow
            $installedSummary += "<tr><td>New Outlook Toggle</td><td><span>OFF (Classic Outlook is Default)</span></td></tr>"
        }
    }

$installedSummary += "</table>"
Add-Content $FullLogFilePath $installedSummary

    if ($newOutlookInstalled) {
        $newOutlookVersion = Get-NewOutlookVersion

        Write-Host "`n========== New Outlook Information ==========" -ForegroundColor Cyan

        Add-Content $FullLogFilePath "<h3>New Outlook</h3>"
        Add-Content $FullLogFilePath "<table>"
        Add-Content $FullLogFilePath "<tr><th>Property</th><th>Value</th></tr>"

        # Version row
        if ($newOutlookVersion) {
            Write-Host "New Outlook Version: $newOutlookVersion"
            Add-Content $FullLogFilePath "<tr><td>Version</td><td>$newOutlookVersion</td></tr>"
        } else {
            Write-Host "New Outlook version could not be determined." -ForegroundColor Yellow
            Add-Content $FullLogFilePath "<tr><td>Version</td><td><span class='warning'>Could not be determined</span></td></tr>"
        }

        Add-Content $FullLogFilePath "</table>"
    }

    if ($classicInstalled) {
        Write-Host "`n========== Classic Outlook Information ==========" -ForegroundColor Cyan
        Get-OfficeConfiguration

        # Capture Version from Registry
        $versionKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        $officeVersion = $null
        $officeBuild = $null

        if (Test-Path $versionKey) {
            $versionData = Get-ItemProperty -Path $versionKey -ErrorAction SilentlyContinue
            if ($versionData.VersionToReport) {
                # Example: "16.0.18025.20000" → extract build part
                $fullVersion = $versionData.VersionToReport
                $versionParts = $fullVersion -split '\.'
                if ($versionParts.Length -ge 4) {
                    $officeBuild = "$($versionParts[2]).$($versionParts[3])"  # e.g. 18025.20000
                    $officeVersion = "$($versionParts[0]).$($versionParts[1])" # e.g. 16.0
                }
                Write-Host "Office Version: $fullVersion"
                Write-Host "Build Number: $officeBuild"
            }
        }

        $licenseType = Get-OfficeLicenseType
        Write-Host "License Type: $licenseType"


        # Build Comparison Helper
        function Compare-Build {
            param (
                [string]$current,
                [string]$minimum
            )
            $currParts = $current -split '\.'
            $minParts  = $minimum -split '\.'

            for ($i = 0; $i -lt $currParts.Length; $i++) {
                if ([int]$currParts[$i] -gt [int]$minParts[$i]) { return $true }
                elseif ([int]$currParts[$i] -lt [int]$minParts[$i]) { return $false }
            }
            return $true  # Equal
        }

        # -----------------------------
        # Compatibility Check Output
        # -----------------------------

        if ($classicInstalled) {
            Add-Content $FullLogFilePath "<h3>Classic Outlook</h3>"
            $requirementsKB = "(<a href='https://support.exclaimer.com/hc/en-gb/articles/4406058988945-System-Requirements-for-Exclaimer#:~:text=365%20mailboxes%20only-,Windows,-Outlook%20on%20Windows' target='_blank'>Requirements</a>)"
            $buildSupport = ""

            # -------------------------------
            # ADDITIVE: derive bitness + version from build
            # -------------------------------
            $outlookBitness = "Unknown"
            $outlookVersion = "Unknown"

            # Build-to-Version mapping. Minimum supported: Version 2408 (Build 17928.x).
            # Builds below Version 2408 are treated as Not Supported regardless of build number.
            # Note: per-license minimum thresholds are enforced separately via $minimumSupportedBuilds.
            #   - Subscription (M365):       Version 2409 (18025.20000)
            #   - Retail (Perpetual/Sub):     Version 2501 (18429.20132)
            #   - Volume Licensed (Perpetual): Version 2408 (17932.20222)
            # Source: https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date
            # Last updated: July 2026. If the detected build is newer than the highest entry here,
            # the output table will flag this so the script can be refreshed against the MS article.
            $map = @{
                # M365 channel builds (Current Channel, Monthly Enterprise, Semi-Annual Enterprise).
                # Builds shared across channels appear once; the version label is channel-agnostic.
                # Source: https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date
                "20131.20154"="2606";"20131.20126"="2606";"20131.20112"="2606";"20131.20090"="2606";
                "20131.20152"="2606";"20131.20150"="2606";
                "20026.20254"="2605";"20026.20236"="2605";"20026.20182"="2605";"20026.20168"="2605";
                "20026.20140"="2605";"20026.20112"="2605";"20026.20076"="2605";"20026.20166"="2605";
                "19929.20264"="2604";"19929.20220"="2604";"19929.20172"="2604";"19929.20164"="2604";
                "19929.20162"="2604";"19929.20136"="2604";"19929.20106"="2604";"19929.20090"="2604";
                "19822.20288"="2603";"19822.20254"="2603";"19822.20240"="2603";"19822.20182"="2603";
                "19822.20168"="2603";"19822.20142"="2603";"19822.20114"="2603";"19822.20180"="2603";
                "19725.20346"="2602";"19725.20320"="2602";"19725.20190"="2602";"19725.20172"="2602";
                "19725.20152"="2602";"19725.20126"="2602";"19725.20244"="2602";"19725.20170"="2602";
                "19628.20214"="2601";"19628.20204"="2601";"19628.20166"="2601";"19628.20150"="2601";
                "19530.20282"="2512";"19530.20260"="2512";"19530.20226"="2512";"19530.20184"="2512";
                "19530.20144"="2512";"19530.20138"="2512";
                "19426.20314"="2511";"19426.20294"="2511";"19426.20260"="2511";"19426.20218"="2511";
                "19426.20186"="2511";"19426.20170"="2511";
                "19328.20306"="2510";"19328.20292"="2510";"19328.20266"="2510";"19328.20244"="2510";
                "19328.20232"="2510";"19328.20190"="2510";"19328.20178"="2510";"19328.20158"="2510";
                "19231.20300"="2509";"19231.20274"="2509";"19231.20246"="2509";"19231.20216"="2509";
                "19231.20194"="2509";"19231.20172"="2509";"19231.20156"="2509";
                "19127.20730"="2508";"19127.20678"="2508";"19127.20648"="2508";"19127.20646"="2508";
                "19127.20622"="2508";"19127.20570"="2508";"19127.20532"="2508";"19127.20484"="2508";
                "19127.20402"="2508";"19127.20384"="2508";"19127.20358"="2508";"19127.20314"="2508";
                "19127.20302"="2508";"19127.20264"="2508";"19127.20240"="2508";"19127.20222"="2508";
                "19029.20300"="2507";"19029.20274"="2507";"19029.20244"="2507";"19029.20208"="2507";
                "19029.20184"="2507";"19029.20156"="2507";"19029.20136"="2507";
                "18925.20268"="2506";"18925.20242"="2506";"18925.20216"="2506";"18925.20184"="2506";
                "18925.20168"="2506";"18925.20158"="2506";"18925.20138"="2506";
                "18827.20244"="2505";"18827.20230"="2505";"18827.20202"="2505";"18827.20176"="2505";
                "18827.20164"="2505";"18827.20150"="2505";"18827.20140"="2505";"18827.20128"="2505";
                "18730.20260"="2504";"18730.20240"="2504";"18730.20226"="2504";"18730.20220"="2504";
                "18730.20186"="2504";"18730.20168"="2504";"18730.20142"="2504";
                "18623.20316"="2503";"18623.20302"="2503";"18623.20298"="2503";"18623.20266"="2503";
                "18623.20208"="2503";"18623.20178"="2503";"18623.20156"="2503";
                "18526.20714"="2502";"18526.20696"="2502";"18526.20672"="2502";"18526.20660"="2502";
                "18526.20634"="2502";"18526.20604"="2502";"18526.20546"="2502";"18526.20472"="2502";
                "18526.20438"="2502";"18526.20416"="2502";"18526.20336"="2502";"18526.20286"="2502";
                "18526.20264"="2502";"18526.20168"="2502";"18526.20144"="2502";
                "18429.20240"="2501";"18429.20216"="2501";"18429.20200"="2501";"18429.20158"="2501";
                "18429.20132"="2501";
                "18324.20272"="2412";"18324.20240"="2412";"18324.20194"="2412";"18324.20190"="2412";
                "18324.20168"="2412";
                "18227.20240"="2411";"18227.20222"="2411";"18227.20162"="2411";"18227.20152"="2411";
                "18129.20242"="2410";"18129.20200"="2410";"18129.20158"="2410";
                "18025.20242"="2409";"18025.20214"="2409";"18025.20160"="2409";"18025.20140"="2409";
                "18025.20104"="2409";"18025.20096"="2409";
                "17928.20776"="2408";"17928.20762"="2408";"17928.20742"="2408";"17928.20730"="2408";
                "17928.20708"="2408";"17928.20700"="2408";"17928.20654"="2408";"17928.20604"="2408";
                "17928.20588"="2408";"17928.20572"="2408";"17928.20538"="2408";"17928.20512"="2408";
                "17928.20468"="2408";"17928.20440"="2408";"17928.20392"="2408";"17928.20336"="2408";
                "17928.20286"="2408";"17928.20216"="2408";"17928.20156"="2408";"17928.20114"="2408";
                # Office LTSC 2024 / Office 2024 Volume Licensed builds (all Version 2408).
                # Source: https://learn.microsoft.com/en-us/officeupdates/update-history-office-2024
                # Minimum supported for VL perpetual is 17932.20222 (Jan 14 2025), enforced via $minimumSupportedBuilds.
                "17932.20790"="2408";"17932.20776"="2408";"17932.20742"="2408";"17932.20700"="2408";
                "17932.20670"="2408";"17932.20638"="2408";"17932.20620"="2408";"17932.20602"="2408";
                "17932.20574"="2408";"17932.20540"="2408";"17932.20496"="2408";"17932.20428"="2408";
                "17932.20408"="2408";"17932.20396"="2408";"17932.20360"="2408";"17932.20328"="2408";
                "17932.20286"="2408";"17932.20252"="2408";"17932.20222"="2408";"17932.20190"="2408";
                "17932.20162"="2408";"17932.20130"="2408";
            }
            # Highest build in map - used to detect builds newer than this script's mapping.
            # Update this value whenever the $map is refreshed.
            # For VL/LTSC 2024, the highest known build is 17932.20790 (May 2026).
            # For M365 channels, the highest known build is 20131.20154 (Version 2606, July 2026).
            $mapHighestBuild = "20131.20154"  # Version 2606, July 14 2026
            $mapHighestBuildVL = "17932.20790"  # Office LTSC 2024 VL, May 14 2026

            # Minimum supported build for version map display (Version 2408).
            # Anything below Build 17928.x is shown as Not Supported in the Outlook Version column.
            # Per-license thresholds (e.g. VL requires 17932.20222) are enforced separately via $minimumSupportedBuilds.
            $minimumSupportedMapBuild = "17928.00000"
            # Build 16.0.19725 is the first build with baseline security mode support.
            # Below this build, the add-in requires EWS to be enabled.
            # At or above this build, the add-in uses baseline security mode and EWS is no longer required.
            # Source: https://learn.microsoft.com/en-us/microsoft-365/baseline-security-mode/baseline-security-mode-settings
            $ewsBaselineBuild = "19725.20000"
            $Global:buildRequiresEws = -not (Compare-Build -current $officeBuild -minimum $ewsBaselineBuild)
            $buildRequiresEws = $Global:buildRequiresEws
            $outlookVersionNote = $null
            if ($officeBuild) {
                $outlookVersion = $map[$officeBuild]

                # Determine whether this is a VL/LTSC build (17932.x range) or a standard channel build
                $isVLBuild = $officeBuild -match "^17932\."

                if (-not $outlookVersion) {
                    # Build not in map - determine why
                    if ($isVLBuild -and (Compare-Build -current $officeBuild -minimum $mapHighestBuildVL)) {
                        # VL build is newer than the highest VL entry in $map
                        $outlookVersion = "Unknown (VL build newer than script mapping)"
                        $outlookVersionNote = "newer"
                    } elseif (-not $isVLBuild -and (Compare-Build -current $officeBuild -minimum $mapHighestBuild)) {
                        # Standard channel build is newer than the highest entry in $map
                        $outlookVersion = "Unknown (build newer than script mapping)"
                        $outlookVersionNote = "newer"
                    } elseif (-not (Compare-Build -current $officeBuild -minimum $minimumSupportedMapBuild)) {
                        # Build is below the minimum supported threshold
                        $outlookVersion = "Not Supported"
                        $outlookVersionNote = "unsupported"
                    } else {
                        # Build falls within the supported range but is not in the map (intermediate patch)
                        $outlookVersion = "Unknown (build not in mapping table)"
                        $outlookVersionNote = "unknown"
                    }
                } elseif (-not (Compare-Build -current $officeBuild -minimum $minimumSupportedMapBuild)) {
                    # Build is in map but is below minimum supported threshold (should not normally happen
                    # after the map cleanup, but guard against it)
                    $outlookVersionNote = "unsupported"
                }
            }

            $ctrPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
            if (Test-Path $ctrPath) {
                $ctrProps = Get-ItemProperty $ctrPath
                if ($ctrProps.PSObject.Properties.Name -contains "Platform") {
                    $outlookBitness = if ($ctrProps.Platform -eq "x64") { "64-bit" } else { "32-bit" }
                }
            }
            # -------------------------------

            if ($officeBuild -and $minimumSupportedBuilds.ContainsKey($licenseType)) {
                $requiredBuild = $minimumSupportedBuilds[$licenseType]

                if (Compare-Build -current $officeBuild -minimum $requiredBuild) {
                    Write-Host "`nOutlook $officeVersion (Build $officeBuild) license type '$licenseType' is SUPPORTED." -ForegroundColor Green
                    $buildSupport = "<span class='success'>Supported $requirementsKB</span>"
                } else {
                    Write-Host "`n! Outlook $officeVersion (Build $officeBuild) license type '$licenseType' is NOT SUPPORTED." -ForegroundColor Red
                    Write-Host "  -> Minimum required build for this license: $requiredBuild" -ForegroundColor Gray
                    $buildSupport = "<span class='fail'>Not Supported (Required: $requiredBuild) $requirementsKB</span>"
                }

            } elseif (-not $minimumSupportedBuilds.ContainsKey($licenseType)) {
                Write-Host "`n! License type '$licenseType' is unknown or not mapped. Cannot validate support." -ForegroundColor Yellow
                $buildSupport = "<span class='warning'>Unknown license type $requirementsKB</span>"
            } else {
                Write-Host "`n! Office build number not detected. Cannot validate version compatibility." -ForegroundColor Yellow
                $buildSupport = "<span class='warning'>Build not detected $requirementsKB</span>"
            }

            # Build the Outlook Version cell - annotate when the build is newer than the map
            $outlookVersionCell = switch ($outlookVersionNote) {
                "newer"       { "<span class='warning' title='This build was released after the script mapping was last updated. The version shown is an estimate - check the Microsoft update history page to confirm.'>$outlookVersion ⚠️</span>" }
                "unsupported" { "<span class='fail'>$outlookVersion</span>" }
                "unknown"     { "<span class='warning'>$outlookVersion</span>" }
                default       { $outlookVersion }
            }

            # Write HTML table with version info
            $classicOutlookTable = @"
<table>
    <tr>
        <th>Office Version</th>
        <th>Build</th>
        <th>Outlook Version</th>
        <th>Bitness</th>
        <th>License Type</th>
        <th>Compatibility</th>
    </tr>
    <tr>
        <td>$officeVersion</td>
        <td>$officeBuild</td>
        <td>$outlookVersionCell</td>
        <td>$outlookBitness</td>
        <td>$licenseType</td>
        <td>$buildSupport</td>
    </tr>
</table>
"@
        Add-Content $FullLogFilePath $classicOutlookTable

# Show an EWS deprecation warning when the build is below the baseline security mode threshold
if ($Global:buildRequiresEws) {
    Add-Content $FullLogFilePath @"
<div class="info-after-warning">
    <strong>⚠️ EWS Deprecation Warning — Action Required</strong><br>
    This Outlook build (<code>$officeBuild</code>) is below Build 16.0.19725 and depends on
    <a href="https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online" target="_blank">Exchange Web Services (EWS)</a>
    for Cloud Add-in functionality.<br><br>
    Microsoft is retiring EWS in Exchange Online — phased disablement begins <strong>October 1, 2026</strong>,
    with permanent retirement on <strong>April 1, 2027</strong>. Once EWS is disabled or retired,
    Cloud Add-ins will stop functioning on this build.<br><br>
    Outlook builds at Build 19725.20000 or above use
    <a href="https://learn.microsoft.com/en-us/microsoft-365/baseline-security-mode/baseline-security-mode-settings?view=o365-worldwide#exchange-web-services-requirements" target="_blank">baseline security mode</a>
    and no longer depend on EWS. We recommend updating Outlook ahead of October 2026.
</div>
"@
}
            # Show a banner when the build is newer than the script's mapping table
            if ($outlookVersionNote -eq "newer") {
                Write-Host "`n========== ℹ️  BUILD NEWER THAN SCRIPT MAPPING ==========" -ForegroundColor Yellow
                Write-Host "Build $officeBuild is newer than the highest build in this script's version mapping ($mapHighestBuild)." -ForegroundColor Yellow
                Write-Host "The Outlook version shown may be inaccurate. Verify at:" -ForegroundColor Yellow
                Write-Host "https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date" -ForegroundColor Cyan

                Add-Content $FullLogFilePath @"
<div class="info-after-warning">
    <strong>ℹ️ Build newer than script version mapping</strong><br>
    Build <code>$officeBuild</code> is newer than the highest build recorded in this script (<code>$mapHighestBuild</code>). The Outlook version displayed above may be inaccurate.<br>
    Verify the exact version at: <a href="https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date" target="_blank">Microsoft 365 Apps update history</a>
</div>
"@
            }

            # Show a banner when the build is below the minimum supported threshold
            if ($outlookVersionNote -eq "unsupported") {
                Write-Host "`n========== ❌  OUTLOOK VERSION NOT SUPPORTED ==========" -ForegroundColor Red
                Write-Host "Build $officeBuild is below the minimum supported Build 17928.x." -ForegroundColor Red
                Write-Host "Exclaimer Add-In is not supported on this version of Outlook." -ForegroundColor Red

                Add-Content $FullLogFilePath @"
<div class="info-after-error">
    <strong>❌ Outlook version not supported</strong><br>
    Build <code>$officeBuild</code> is below the minimum supported Build 17928.x.<br>
    The Exclaimer Add-In requires at least Version 2408 (Volume Licensed) or higher depending on license type. Please update Microsoft 365 Apps to a supported release.<br>
    See: <a href="https://support.exclaimer.com/hc/en-gb/articles/4406058988945" target="_blank">Exclaimer System Requirements</a>
</div>
"@
            }
            $officeBitnessDocUrl = 'https://support.microsoft.com/en-gb/office/choose-between-the-64-bit-or-32-bit-version-of-office-2dee7807-8f95-4d0c-b5fe-6c6f49b8d261?blocks+of+information+or+graphics.=&utm_source=chatgpt.com#:~:text=You%27re%20using%20add%2Dins%20with%20Outlook%2C%20Excel%2C%20or%20other%20Office%20or%20Microsoft%20365%20apps'
            if ($outlookBitness -eq '32-bit') {
                Add-Content $FullLogFilePath @"
<div class="side-note">
While 32-bit applications can work with add-ins, they can use up a system's available virtual address space. With 64-bit apps, you have up to 128&nbsp;TB of virtual address space which the app and any add-ins running the same process can share. With 32-bit apps, you might get as little as 2&nbsp;GB of virtual address space which in many cases is not enough and can cause the app to stop responding or crash. <a href="$officeBitnessDocUrl">Microsoft Support</a>
</div>
"@
            }

            # Check for known issue with build 19822.20114 (Version 2603)
            if ($officeBuild -eq "19822.20114") {
                Write-Host "`n========== ⚠️  KNOWN ISSUE DETECTED ==========" -ForegroundColor Yellow
                Write-Host "After updating to Version 2603 (Build 19822.20114), XML manifest addins missing from ribbon after version 2603 update." -ForegroundColor Red
                Write-Host "`nSTATUS: For more information, see: https://support.microsoft.com/en-us/office/outlook-on-premise-exchange-web-add-ins-stopped-loading-after-updating-to-version-2603-fe1a0622-f190-4bb8-9fdb-e541591af5be" -ForegroundColor Yellow
                Write-Host "The Outlook Team is addressing this issue with a change from the service. The change is expected to be available by end of day 4/22/26. To pick up the change, restart Outlook. It can take up to four hours for service changes to be picked up by Outlook." -ForegroundColor White
                
                Add-Content $FullLogFilePath @"
<div class="info-after-warning">
    <h3 style="color: #f57c00; margin-top: 0;">⚠️ Known Issue Detected</h3>
    <p><strong>Issue:</strong> After updating to Version 2603 (Build 19822.20114), XML manifest addins missing from ribbon after version 2603 update.</p>
    <p><strong>STATUS:</strong> <span style="color: green; font-weight: bold;"><a href="https://support.microsoft.com/en-us/office/outlook-on-premise-exchange-web-add-ins-stopped-loading-after-updating-to-version-2603-fe1a0622-f190-4bb8-9fdb-e541591af5be" target="_blank">Learn more about this issue</a></span></p>
    <p>MICROSOFT: The Outlook Team is addressing this issue with a change from the service. The change is expected to be available by end of day 4/22/26. To pick up the change, restart Outlook. It can take up to four hours for service changes to be picked up by Outlook.</p>
</div>
"@
            }
        }

    }
    
}
if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)','Outlook Web (Windows)','Multiple / All')) {
    InspectOutlookConfiguration
}

# --- Check Classic Outlook Encoding Configuration ---
# Build function to search user hives
function InspectClassicOutlookEncoding {
    param([string]$userEmail)
    
    # Extract username from email (UPN - part before @)
    $emailUsername = $userEmail -split '@' | Select-Object -First 1
    
    # Search for user hives in HKEY_USERS
    $userHives = @()
    try {
        $hives = Get-ChildItem 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.PSChildName -match '^S-' -and 
                $_.PSChildName.Length -ge 30 -and 
                $_.PSChildName -notmatch '_Classes$'
            }
        
        foreach ($hive in $hives) {
            $hiveName = $hive.PSChildName
            $volEnvPath = "Registry::HKEY_USERS\$hiveName\Volatile Environment"
            
            try {
                $volEnv = Get-Item $volEnvPath -ErrorAction SilentlyContinue
                if ($volEnv) {
                    $registeredUsername = $volEnv.GetValue('USERNAME')
                    if ($registeredUsername) {
                        $userHives += @{
                            Hive = $hiveName
                            Username = $registeredUsername
                        }
                    }
                }
            }
            catch {
                continue
            }
        }
    }
    catch {
        # Silently handle errors
    }
    
    $matchedHive = $null
    
    # Try to find matching hive by username
    foreach ($hiveInfo in $userHives) {
        if ($hiveInfo.Username -eq $emailUsername) {
            $matchedHive = $hiveInfo.Hive
            break
        }
    }
    
    # If no match found, use the only available hive or ask the user to choose
    if (-not $matchedHive -and $userHives.Count -gt 0) {
        if ($userHives.Count -eq 1) {
            $matchedHive = $userHives[0].Hive
            Write-Host "Only one user hive found for '$($userHives[0].Username)'. Using it automatically." -ForegroundColor Yellow
        }
        else {
            Write-Host "Could not automatically match username '$emailUsername' to a hive." -ForegroundColor Yellow
            Write-Host "Found the following user hives:`n" -ForegroundColor Cyan
            
            for ($i = 0; $i -lt $userHives.Count; $i++) {
                Write-Host "[$($i+1)] $($userHives[$i].Username)`n" -ForegroundColor Yellow
            }
            
            $selection = Read-Host "Enter the number of the correct hive (or press Enter to skip)"
            
            if ($selection -match '^\d+$' -and [int]$selection -gt 0 -and [int]$selection -le $userHives.Count) {
                $matchedHive = $userHives[[int]$selection - 1].Hive
            }
        }
    }
    
    # Get encoding value if hive was identified
    if ($matchedHive) {
        $encodingPath = "Registry::HKEY_USERS\$matchedHive\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\MSHTML\International"
        
        try {
            $encodingReg = Get-Item $encodingPath -ErrorAction SilentlyContinue
            if ($encodingReg) {
                $codePage = $encodingReg.GetValue('Default_CodePageOut')
                
                if ($null -ne $codePage) {
                    # Console output
                    if ($codePage -eq 65001) {
                        Write-Host "Outlook encoding is configured to UTF-8" -ForegroundColor Green
                    }
                    else {
                        Write-Host "Outlook encoding code page: $codePage" -ForegroundColor White
                    }
                    
                    # HTML output
                    Add-Content $FullLogFilePath "<h3>Classic Outlook Encoding Configuration</h3>"
                    Add-Content $FullLogFilePath "<table><tr><th>Setting</th><th>Value</th><th>Notes</th></tr>"
                    
                    $htmlNote = $null
                    if ($codePage -eq 65001) {
                        Add-Content $FullLogFilePath "<tr><td>Preferred Encoding</td><td><span class='success'>65001 (UTF-8)</span></td><td><span class='success'>OK</span></td></tr>"
                        $htmlNote = "<div class='info-after-success'>✅ Encoding is UTF-8; no action required.</div>"
                    }
                    else {
                        Add-Content $FullLogFilePath "<tr><td>Preferred Encoding</td><td><span class='warning'>$codePage</span></td><td><span class='warning'>Please review</span></td></tr>"
                        $htmlNote = "<div class='info-after-warning'>
                                ⚠️ Detected code page $codePage. Please review 
                                <a href='https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers' target='_blank'>Code Page Identifiers</a>.
                               <br> If you experience issues with foreign characters displaying incorrectly in emails, see 
                                <a href='https://support.exclaimer.com/hc/en-gb/articles/6622042157085-Foreign-characters-are-not-displayed-correctly' target='_blank'>support article</a>.
                                </div>"
                    }
                    
                    Add-Content $FullLogFilePath "</table>"
                    if ($htmlNote) { Add-Content $FullLogFilePath $htmlNote }
                }
                else {
                    Add-Content $FullLogFilePath "<h3>Classic Outlook Encoding Configuration</h3>"
                    Add-Content $FullLogFilePath "<p class='warning'>⚠️ Outlook encoding configuration not found. Preferred Encoding registry value is not set.</p>"
                }
            }
            else {
                Add-Content $FullLogFilePath "<h3>Classic Outlook Encoding Configuration</h3>"
                Add-Content $FullLogFilePath "<p class='warning'>⚠️ Outlook encoding registry path not found for this user.</p>"
            }
        }
        catch {
            Add-Content $FullLogFilePath "<h3>Classic Outlook Encoding Configuration</h3>"
            Add-Content $FullLogFilePath "<p class='warning'>⚠️ Error accessing Outlook encoding configuration: $_</p>"
        }
    }
    elseif ($userHives.Count -eq 0) {
        Add-Content $FullLogFilePath "<h3>Classic Outlook Encoding Configuration</h3>"
        Add-Content $FullLogFilePath "<p class='warning'>⚠️ No user hives found in HKEY_USERS.</p>"
    }
    else {
        Add-Content $FullLogFilePath "<h3>Classic Outlook Encoding Configuration</h3>"
        Add-Content $FullLogFilePath "<p class='warning'>⚠️ User hive selection was skipped. Encoding configuration could not be retrieved.</p>"
    }
}
if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)','Outlook Web (Windows)','Multiple / All')) {
    # Call the function with the user's email
    InspectClassicOutlookEncoding -userEmail $Global:userInput.Email
}


function InspectWordFileBlocking {
    Write-Host "`n--- Word File Block Settings Check (Web Pages) ---" -ForegroundColor Yellow
    # Add-Content $FullLogFilePath "<h3>Word File Block Settings (Web Pages)</h3>"

    $registryChecks = @(
        @{
            Path = "HKCU:\Software\Policies\Microsoft\Office\16.0\Word\Security\FileBlock"
            Scope = "Policy (GPO)"
        },
        @{
            Path = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\FileBlock"
            Scope = "User"
        }
    )

$webPagesBlocked = $false
$webPagesUnknown = $false
$tableRows = ''

foreach ($check in $registryChecks) {
    if (Test-Path $check.Path) {
        $key = Get-ItemProperty -Path $check.Path -ErrorAction SilentlyContinue
        if ($key.HtmlFiles -eq 1 -or $key.HtmlFiles -eq 2) {
            Write-Host "Web Pages BLOCKED" -ForegroundColor Red
            $tableRows += '<tr><td>' + $check.Path + '</td><td class="fail">Blocked</td></tr>'
            $webPagesBlocked = $true
        }
        elseif ($key.HtmlFiles -eq 0) {
            Write-Host "Web Pages allowed" -ForegroundColor Green
            $tableRows += '<tr><td>' + $check.Path + '</td><td class="success">success</td></tr>'
        }
        else {
            $tableRows += '<tr><td>' + $check.Path + '</td><td class="success">No Key Found</td></tr>'
            $webPagesUnknown = $false
        }
    }
}

# Build the complete section (table + messages)
$webSection = '<div class="section"><h3>Word File Block Settings (Web Pages)</h3>'

if ($tableRows) {
    $webSection += '<table>' +
        '<tr><th>Registry Path</th><th>Status</th></tr>' +
        $tableRows +
        '</table>'
}

# Add status messages within the same section
if ($webPagesBlocked) {
    Write-Host "`nWARNING: Word is blocking Web Pages. HTML based signatures cannot be inserted into the Outlook email body." -ForegroundColor Red
    Write-Host "This setting must be disabled to allow signature injection." -ForegroundColor Red
    $webSection += '<div class="info-after-error">' +
        '<strong>Impact:</strong> Web Page file types (.htm/.html) are currently blocked in Microsoft Word. ' +
        'This prevents Exclaimer signatures from being inserted into Outlook.' +
        '<div style="margin-top:10px; font-weight:normal;">' +
        '<strong>Required action:</strong>' +
        '<ol style="margin-top:6px;">' +
        '<li>Open Microsoft Word</li>' +
        '<li>Go to <strong>File</strong> &gt; <strong>Options</strong> &gt; <strong>Trust Center</strong></li>' +
        '<li>Select <strong>Trust Center Settings</strong> &gt; <strong>File Block Settings</strong></li>' +
        '<li>Locate <strong>Web Pages</strong> and ensure the checkbox is <strong>unchecked</strong></li>' +
        '<li>Restart Outlook after making changes</li>' +
        '</ol>' +
        '<strong>Important:</strong> If the setting is greyed out, it is enforced by Group Policy and must be reviewed by your IT administrator.' +
        '</div></div>'
}
elseif ($webPagesUnknown) {
    Write-Host "`nWARNING: Unable to determine Word Web Page File Block setting." -ForegroundColor Yellow
    $webSection += '<div class="info-after-warning">' +
        '<strong>Manual verification required:</strong> The script was unable to determine the current Web Page File Block setting in Microsoft Word.' +
        '<div style="margin-top:10px; font-weight:normal;">' +
        '<strong>Please check:</strong>' +
        '<ol style="margin-top:6px;">' +
        '<li>Open Microsoft Word</li>' +
        '<li>Go to <strong>File</strong> &gt; <strong>Options</strong> &gt; <strong>Trust Center</strong></li>' +
        '<li>Select <strong>Trust Center Settings</strong> &gt; <strong>File Block Settings</strong></li>' +
        '<li>Confirm whether <strong>Web Pages</strong> (.htm/.html) is checked or <strong>unchecked</strong></li>' +
        '</ol>' +
        'If the option is greyed out, it is likely enforced by Group Policy.' +
        '<br><br>' +
        '<strong>When replying to your ticket please confrim:</strong>' +
        '<ul style="margin-top:6px;">' +
        '<li>Whether it is checked or unchecked</li>' +
        '<li>Whether it is editable or greyed out</li>' +
        '</ul>' +
        'We will review your response and advise on next steps.' +
        '</div></div>'
}
else {
    Write-Host "No blocking detected for Web Pages." -ForegroundColor Green
    $webSection += '<div class="info-after-success">✅ <strong>No issues detected:</strong> No Word File Block restrictions were found for Web Pages.</div>'
}
# Close the section and output
$webSection += '</div>'
Add-Content -Path $FullLogFilePath -Value $webSection
}
if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)','Outlook Web (Windows)','Multiple / All')) {
    InspectWordFileBlocking
}

function InspectExclaimerCloudSignatureAgent {
    # Define registry paths to search
    $registryPaths = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
    )

    Write-Host "`n========== Exclaimer Cloud Signature Update Agent for Windows ==========" -ForegroundColor Cyan
    $foundApps = @()

    foreach ($path in $registryPaths) {
        try {
            $apps = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
                Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            } | Where-Object {
                $_.DisplayName -like "*Cloud Signature Update Agent*"
            } | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, HelpLink, URLUpdateInfo,
            @{ Name = 'InstallType'; Expression = {
                    if ($_.URLUpdateInfo -like "*Exclaimer.CloudSignatureAgent.application*") {
                        'Click-Once'
                    }
                    else {
                        'MSI'
                    }
                }
            }

            if ($apps) {
                $foundApps += $apps
            }
        }
        catch {
            # Ignore errors
        }
    }

    if ($foundApps.Count -gt 0) {
        # Console output
        $foundApps | Select-Object DisplayName, DisplayVersion, InstallType | Format-Table -AutoSize

        # HTML output
        Add-Content $FullLogFilePath '<div class="section">'
        Add-Content $FullLogFilePath "<h3>Exclaimer Cloud Signature Update Agent for Windows</h3>"
        Add-Content $FullLogFilePath "<table><tr><th>Display Name</th><th>Version</th><th>Install Type</th></tr>"

        foreach ($app in $foundApps) {
            $displayName = [System.Web.HttpUtility]::HtmlEncode($app.DisplayName)
            $version = [System.Web.HttpUtility]::HtmlEncode($app.DisplayVersion)
            $installType = [System.Web.HttpUtility]::HtmlEncode($app.InstallType)
            $row = "<tr><td>$displayName</td><td>$version</td><td>$installType</td></tr>"
            Add-Content $FullLogFilePath $row
        }

        Add-Content $FullLogFilePath "</table>"
        Add-Content $FullLogFilePath "</div>"
    }
    else {
        Write-Host "The Exclaimer Cloud Signature Update Agent is not installed." -ForegroundColor Yellow
        Add-Content $FullLogFilePath '<div class="section">'
        Add-Content $FullLogFilePath "<h3>Exclaimer Cloud Signature Update Agent for Windows</h3>"
        Add-Content $FullLogFilePath '<table><tr><th>Status</th><th>Details</th></tr><tr><td class="success">Not Installed</td><td>✅ The Exclaimer Cloud Signature Update Agent is not installed.</td></tr></table>'
        Add-Content $FullLogFilePath "</div>"
    }

    # Remove the extra </div> since we now close properly
    # Add-Content $FullLogFilePath "</div>"

            # Checking for existing local signatures
        $baseSignaturePath = [System.IO.Path]::Combine($env:APPDATA, "Microsoft")
        $possibleFolders = @("Signatures", "Handtekeningen")
        $signaturePath = $null

        foreach ($folder in $possibleFolders) {
            $fullPath = Join-Path $baseSignaturePath $folder
            if (Test-Path $fullPath) {
                $signaturePath = $fullPath
                break
            }
        }

        if ($signaturePath) {
            $htmFiles = Get-ChildItem -Path $signaturePath -Filter *.htm -File -ErrorAction SilentlyContinue

            if ($htmFiles.Count -gt 0) {
                # Console output
                Write-Host "`n--- Local Outlook Signatures Found ---" -ForegroundColor Yellow
                $signatureData = $htmFiles | ForEach-Object {
                    $content = Get-Content -Path $_.FullName -Raw -ErrorAction SilentlyContinue
                    $hasRemialSans = ($content -match "remialcxesans")
                    $exclaimerUsed = if ($hasRemialSans) { "Yes" } else { "No" }

                    # Output to console
                    [PSCustomObject]@{
                        Name         = $_.BaseName
                        DateModified = $_.LastWriteTime
                        Exclaimer    = $exclaimerUsed
                    }
                }

                $signatureData | Format-Table -AutoSize

                # HTML output
                Add-Content $FullLogFilePath '<div class="section">'
                Add-Content $FullLogFilePath "<h3>Local Outlook Signatures</h3>"
                Add-Content $FullLogFilePath "<table><tr><th>Name</th><th>Date Modified</th><th>Exclaimer Signature</th></tr>"

                foreach ($sig in $signatureData) {
                    $sigRow = "<tr><td>$($sig.Name)</td><td>$($sig.DateModified)</td><td>$($sig.Exclaimer)</td></tr>"
                    Add-Content $FullLogFilePath $sigRow
                }

                Add-Content $FullLogFilePath "</table>"
                Add-Content $FullLogFilePath "</div>"

            } else {
                Write-Host "`nNo .htm signature files found in $signaturePath" -ForegroundColor DarkGray
                Add-Content $FullLogFilePath '<div class="section">'
                Add-Content $FullLogFilePath "<h3>Local Outlook Signatures</h3>"
                Add-Content $FullLogFilePath "<table><tr><th>Status</th><th>Details</th></tr><tr><td class='success'>No Signatures Found</td><td>✅ No local signature files found in $signaturePath</td></tr></table>"
                Add-Content $FullLogFilePath "</div>"
            }
        }
    }
if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)','Outlook Web (Windows)','Multiple / All')) {
    InspectExclaimerCloudSignatureAgent
}


function InspectWebView2Runtime {
Write-Host "`n========== Microsoft Edge WebView2 Runtime ==========" -ForegroundColor Cyan
    # Define registry paths to search
    $registryPaths = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
    )

    $webviewApps = @()

    foreach ($path in $registryPaths) {
        try {
            $apps = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
                Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            } | Where-Object {
                $_.DisplayName -like "*WebView2*"
            } | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate,
            @{ Name = 'InstallType'; Expression = { 'MSI' } }

            if ($apps) {
                $webviewApps += $apps
            }
        }
        catch {
            # Silently ignore errors
        }
    }

    if ($webviewApps.Count -gt 0) {
        # Console output
        $webviewApps | Select-Object DisplayName, DisplayVersion, InstallType | Format-Table -AutoSize

        # HTML output
        Add-Content $FullLogFilePath "<h3>Microsoft Edge WebView2 Runtime</h3>"
        Add-Content $FullLogFilePath "<table><tr><th>Display Name</th><th>Version</th><th>Install Type</th></tr>"

        foreach ($app in $webviewApps) {
            $displayName = [System.Net.WebUtility]::HtmlEncode($app.DisplayName)
            $version = [System.Net.WebUtility]::HtmlEncode($app.DisplayVersion)
            $installType = [System.Net.WebUtility]::HtmlEncode($app.InstallType)
            $row = "<tr><td>$displayName</td><td>$version</td><td>$installType</td></tr>"
            Add-Content $FullLogFilePath $row
        }

        Add-Content $FullLogFilePath "</table>"
    }
    else {
        Write-Host "Microsoft Edge WebView2 Runtime is not installed." -ForegroundColor Yellow
        Add-Content $FullLogFilePath "<p class='warning'>Microsoft Edge WebView2 Runtime is not installed.</p>"
    }
}
if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)','Outlook Web (Windows)','Multiple / All')) {
    InspectWebView2Runtime
}

# -------------------------------------------------------------------
# 📨 EXCLAIMER ADD-IN DETAILS COLLECTION (User or Admin)
# -------------------------------------------------------------------
    Write-Host ""
    Write-Host "=== Exclaimer Add-in Information ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "ℹ️ This includes deployment information and the current State of the Add-in for the user reporting issues." -ForegroundColor Cyan
    Write-Host "`nIf you do not run the next step as a Global Admin, we may need to ask you to run some PowerShell commands manually to collect the required information." -ForegroundColor Red
    Write-Host "`nRecommended action: continue as a Microsoft 365 Global Administrator to collect full details.`n" -ForegroundColor Cyan

# --- Function: Check for Exchange Online Module ---
function CheckExchangeOnlineModule {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Write-Host "✅ Exchange Online Management module is already installed." -ForegroundColor Green
        return $true
    } else {
        Write-Host "⚙️  Exchange Online Management module not found." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "ℹ️  The installation requires the NuGet provider and PowerShell Gallery access." -ForegroundColor Cyan
        Write-Host "    You may see prompts asking to install NuGet or trust the PowerShell Gallery — please answer 'Y' when prompted." -ForegroundColor Cyan
        Write-Host ""

        $installChoice = Read-Host "Would you like to install it now? (Y/N)"
        if ($installChoice.ToUpper() -eq "Y") {
            try {
                Write-Host "`n📦 Preparing to install prerequisites..." -ForegroundColor Cyan

                # --- Ensure NuGet provider is installed ---
                if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                    Write-Host "🔧 Installing NuGet provider..." -ForegroundColor Cyan
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false | Out-Null
                }

                # --- Ensure PowerShell Gallery is trusted ---
                $galleryTrusted = (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue).InstallationPolicy
                if ($galleryTrusted -ne 'Trusted') {
                    Write-Host "🔒 Trusting PowerShell Gallery repository..." -ForegroundColor Cyan
                    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
                }

                # --- Install pinned version to avoid known sign-in issue on later versions ---
                Write-Host "📦 Installing Exchange Online Management module v3.9.0..." -ForegroundColor Cyan
                Install-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -Force -Scope CurrentUser -AllowClobber

                Write-Host "📥 Importing Exchange Online Management module..." -ForegroundColor Cyan
                Import-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -Force

                Write-Host "✅ Installation completed successfully!" -ForegroundColor Green
                return $true

            } catch {
                Write-Host "❌ Failed to install the module: $($_.Exception.Message)" -ForegroundColor Red
                Add-Content $FullLogFilePath "<p class='warning'>Exchange Online Management module installation failed: $([System.Web.HttpUtility]::HtmlEncode($_.Exception.Message))</p>"
                return $false
            }
        } else {
            Write-Host "⚠️ Skipping module installation. Admin access required for automated mailbox queries." -ForegroundColor Yellow
            Add-Content $FullLogFilePath "<p class='warning'>User skipped Exchange Online module installation. Manual Add-in version collection required.</p>"
            return $false
        }
    }
}
# --- Function: Connect to Exchange Online ---
function ConnectExchangeOnlineSession {
    Write-Host "`n🔗 Connecting to Exchange Online..." -ForegroundColor Cyan
    Write-Host "   You will be prompted to Sign in with Microsoft in order to continue." -ForegroundColor Yellow

    Add-Content $FullLogFilePath '<div class="section">'
    Add-Content $FullLogFilePath '<h2>🔗 Exchange Online Connection</h2>'
    Add-Content $FullLogFilePath '<table>'
    Add-Content $FullLogFilePath '<tr><th>Property</th><th>Value</th></tr>'

    try {
        Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        Start-Sleep -Seconds 3
        Connect-ExchangeOnline -ErrorAction Stop

        Write-Host "✅ Connected successfully!" -ForegroundColor Green

        Add-Content $FullLogFilePath '<tr><td><strong>Status</strong></td><td><span class="success">✅ Connected successfully</span></td></tr>'
        Add-Content $FullLogFilePath '</table>'
        Add-Content $FullLogFilePath '</div>'

        return $true

    } catch {
        $errorMessage = $_.Exception.Message
        $errorType    = $_.Exception.GetType().FullName
        $errorStack   = $_.ScriptStackTrace

        Write-Host "❌ Connection failed: $errorMessage" -ForegroundColor Red

        Add-Content $FullLogFilePath '<tr><td><strong>Status</strong></td><td><span class="fail">❌ Connection failed</span></td></tr>'
        Add-Content $FullLogFilePath "<tr><td><strong>Error Type</strong></td><td>$([System.Web.HttpUtility]::HtmlEncode($errorType))</td></tr>"
        Add-Content $FullLogFilePath '</table>'
        Add-Content $FullLogFilePath '<div class="info-after-error"><strong>❌ Exchange Online connection failed.</strong> The error details are below:</div>'        
        Add-Content $FullLogFilePath '</div>'
        Add-Content $FullLogFilePath '<pre style="background-color:#f1f1f1; padding:12px; border-radius:4px; border-left:4px solid #c7254e; overflow-x:auto; font-size:13px;">'
        Add-Content $FullLogFilePath "$([System.Web.HttpUtility]::HtmlEncode($errorMessage))"
        Add-Content $FullLogFilePath ""
        Add-Content $FullLogFilePath "Stack Trace:"
        Add-Content $FullLogFilePath "$([System.Web.HttpUtility]::HtmlEncode($errorStack))"
        Add-Content $FullLogFilePath '</pre>'

        return $false
    }
}

function InpectEXOconfiguration { 
# --- Proceed only if module available ---
if (CheckExchangeOnlineModule) {
        # --- HTML Logging (safe formatting) ---
        Add-Content $FullLogFilePath '<h2>🔐 Exclaimer Exchange Online Information (EXO Admin)</h2>'
    if (ConnectExchangeOnlineSession) {

        
        # --- Organization-level Settings ---
        Write-Host "`nCollecting organization configuration related to Outlook Add-ins..." -ForegroundColor Cyan

try {
            $orgConfig = Get-OrganizationConfig | Select-Object `
                ReleaseTrack,
                OAuth2ClientProfileEnabled,
                OutlookMobileGCCRestrictionsEnabled,
                AppsForOfficeEnabled,
                EwsApplicationAccessPolicy,
                EwsEnabled,
                EwsAllowOutlook

            # Initialize flags
            $addGCCSideNote = $false
            $addAppsSideNote = $false
            $addEwsSideNote = $false
            $addEwsAllowOutlookSideNote = $false

            Add-Content $FullLogFilePath '<div class="section">'
            Add-Content $FullLogFilePath '<h3>⚙️ Organization Configuration - Add-in Compatibility</h3>'
            Add-Content $FullLogFilePath '<table><tr><th>Setting</th><th>Value</th><th>Impact</th></tr>'

            foreach ($prop in $orgConfig.PSObject.Properties) {
                $name  = [System.Web.HttpUtility]::HtmlEncode($prop.Name)
                $rawValue = $prop.Value
                $value = if ($null -eq $rawValue) { 'N/A' } else { [System.Web.HttpUtility]::HtmlEncode([string]$rawValue) }
                $impact = ''

                switch ($prop.Name) {
                    'ReleaseTrack' {
                        switch ($rawValue) {
                            $null               { $impact = '✅ Standard Release' }
                            'FirstRelease'      { $impact = '⚠️ Targeted release for everyone' }
                            'StagedRollout'     { $impact = '⚠️ Targeted release for select users' }
                            default             { $impact = '❌ Review manually.' }
                        }
                    }
                    'OAuth2ClientProfileEnabled' {
                        $impact = if (-not $rawValue) {
                            '❌ Add-ins cannot authenticate properly (modern auth disabled).'
                        } else {
                            '✅ Required for modern add-ins (OK).'
                        }
                    }
                    'OutlookMobileGCCRestrictionsEnabled' {
                        $impact = if ($rawValue) {
                            $addGCCSideNote = $true
                            '❌ Cloud add-ins not supported on Outlook Mobile.'
                        } else {
                            '✅ Mobile add-ins supported.'
                        }
                    }
                    'AppsForOfficeEnabled' {
                        $impact = if (-not $rawValue) {
                            $addAppsSideNote = $true
                            '❌ Add-ins disabled organization-wide.'
                        } else {
                            '✅ Add-ins allowed/enabled.'
                        }
                    }
                    'EwsApplicationAccessPolicy' {
                        if ($Global:buildRequiresEws) {
                            if ([string]::IsNullOrEmpty($rawValue) -or $rawValue -eq 'EnforceNone') {
                                $impact = '✅ No EWS restrictions detected.'
                            } elseif ($rawValue -eq 'EnforceAllowList') {
                                $impact = '⚠️ Only specific apps can use EWS.'
                            } elseif ($rawValue -eq 'EnforceBlockList') {
                                $impact = '⚠️ Some apps are blocked from EWS.'
                            } else {
                                $impact = "⚠️ Unrecognized policy value ($value). Review manually."
                            }
                        } else {
                            $impact = 'ℹ️ Not required. This Outlook build uses baseline security mode and does not depend on EWS.'
                        }
                    }
                    'EwsEnabled' {
                        if ($Global:buildRequiresEws) {
                            $impact = if ($rawValue -eq $false) {
                                $addEwsSideNote = $true
                                '❌ "EwsEnabled" is disabled at org level. This build requires EWS — add-ins will fail.'
                            } elseif ($rawValue -eq $true) {
                                '✅ "EwsEnabled" is enabled. Required for this Outlook build.'
                            } elseif ($null -eq $rawValue) {
                                $addEwsSideNoteWarning = $true
                                '⚠️ "EwsEnabled" state is NULL (not explicitly set).'
                            } else {
                                $addEwsSideNote = $true
                                '⚠️ Unable to determine "EwsEnabled" state. Review manually.'
                            }
                        } else {
                            $impact = 'ℹ️ Not required. This Outlook build uses baseline security mode and does not depend on EWS.'
                        }
                    }
                    'EwsAllowOutlook' {
                        if ($Global:buildRequiresEws) {
                            $impact = if ($rawValue -eq $false) {
                                $addEwsAllowOutlookSideNote = $true
                                '❌ "EwsAllowOutlook" is disabled at org level. This build requires EWS — Outlook add-ins will fail.'
                            } elseif ($rawValue -eq $true) {
                                '✅ "EwsAllowOutlook" is enabled. Required for this Outlook build.'
                            } elseif ($null -eq $rawValue) {
                                $addEwsAllowOutlookSideNoteWarning = $true
                                '⚠️ "EwsAllowOutlook" state is NULL (not explicitly set).'
                            } else {
                                $addEwsAllowOutlookSideNote = $true
                                '⚠️ Unable to determine "EwsAllowOutlook" state. Review manually.'
                            }
                        } else {
                            $impact = 'ℹ️ Not required. This Outlook build uses baseline security mode and does not depend on EWS.'
                        }
                    }
                    Default {
                        $impact = 'Review manually.'
                    }
                }

                Add-Content $FullLogFilePath ("<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>" -f $name, $value, [System.Web.HttpUtility]::HtmlEncode($impact))
            }

            Add-Content $FullLogFilePath '</table></div>'

            $anyEwsWarning = $addEwsSideNoteWarning -or $addEwsSideNote -or $addEwsAllowOutlookSideNoteWarning -or $addEwsAllowOutlookSideNote

            if ($addGCCSideNote) {
                $sideNote = '<div class="info-after-error"><span><b>ℹ️ ''OutlookMobileGCCRestrictionsEnabled'' is ''true'':</b><br>Run this command in PowerShell to set OutlookMobileGCCRestrictionsEnabled to ''false'': <code>Set-OrganizationConfig -OutlookMobileGCCRestrictionsEnabled $false</code></span></div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            if ($addAppsSideNote) {
                $sideNote = '<div class="info-after-error"><span><b>ℹ️ ''AppsForOfficeEnabled'' is disabled:</b><br>Run this command in PowerShell to enable Apps for Office: <code>Set-OrganizationConfig -AppsForOfficeEnabled $true</code></span></div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            if ($anyEwsWarning) {
                $sideNote = '<div class="info-after-note">' +
                    '<span>If you have reopened PowerShell, you may need to run first: <code>Connect-ExchangeOnline</code></span><br><br>' +
                    '<span>Once this is completed, please re-run the full script again to verify the changes made.</span>' +
                    '</div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            if ($addEwsSideNote) {
                $sideNote = '<div class="info-after-error"><span><b>❌ ''EwsEnabled'' is disabled:</b><br>' +
                    'This Outlook build requires EWS for add-in functionality. EWS is disabled at org level — add-ins will fail to initialise.<br><br>' +
                    'Run this command in PowerShell to enable EWS at the organization level:<br>' +
                    '<code>Set-OrganizationConfig -EwsEnabled $true</code><br><br>' +
                    'Note: Mailbox-level EWS settings can still override this organization setting.' +
                    '</span></div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            if ($addEwsAllowOutlookSideNote) {
                $sideNote = '<div class="info-after-error"><span><b>❌ ''EwsAllowOutlook'' is disabled:</b><br>' +
                    'This Outlook build requires EWS for add-in functionality. EwsAllowOutlook is disabled at org level — Outlook is blocked from using EWS and add-ins will fail.<br><br>' +
                    'Run this command to explicitly allow Outlook EWS access at the organization level:<br>' +
                    '<code>Set-OrganizationConfig -EwsAllowOutlook $true</code><br><br>' +
                    'Note: Mailbox-level EWS settings can still override this organization setting.' +
                    '</span></div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }
        }
        
    catch {
        $errorMessage = $_.Exception.Message
        $errorType    = $_.Exception.GetType().FullName

        Write-Host "⚠️ Could not retrieve OrganizationConfig values." -ForegroundColor Yellow
        Write-Host "   Error: $errorMessage" -ForegroundColor Red

        # Determine likely cause for a more actionable message
        $likelyCause = ''
        $remediation = ''

        if ($errorMessage -match 'Access is denied|Unauthorized|insufficient access|not authorized|permissions') {
            $likelyCause = 'The signed-in account does not have sufficient permissions to retrieve organization configuration. Global Administrator or Exchange Administrator role is required.'
            $remediation = 'Sign in with a Global Administrator or Exchange Administrator account and re-run the script.'
        } elseif ($errorMessage -match 'sign.?in|authentication|token|credential|AADSTS|MFA|multi.?factor') {
            $likelyCause = 'Authentication failed or was cancelled. The session may have timed out or MFA was not completed.'
            $remediation = 'Re-run the script and complete the sign-in prompt fully, including any MFA challenge.'
        } elseif ($errorMessage -match 'not connected|pipeline|Connect-ExchangeOnline|no active session') {
            $likelyCause = 'No active Exchange Online session was found. The connection may have dropped or was never established.'
            $remediation = 'Re-run the script to establish a new Exchange Online session.'
        } else {
            $likelyCause = 'An unexpected error occurred while retrieving organization configuration.'
            $remediation = 'Check the error details below and ensure the account has the correct permissions and an active Exchange Online session.'
        }

        Add-Content $FullLogFilePath '<div class="section">'
        Add-Content $FullLogFilePath '<h3>⚙️ Organization Configuration - Add-in Compatibility</h3>'
        Add-Content $FullLogFilePath '<table>'
        Add-Content $FullLogFilePath '<tr><th>Property</th><th>Value</th></tr>'
        Add-Content $FullLogFilePath '<tr><td><strong>Status</strong></td><td><span class="fail">❌ Failed to retrieve organization configuration</span></td></tr>'
        Add-Content $FullLogFilePath "<tr><td><strong>Likely Cause</strong></td><td>$([System.Web.HttpUtility]::HtmlEncode($likelyCause))</td></tr>"
        Add-Content $FullLogFilePath "<tr><td><strong>Error Type</strong></td><td>$([System.Web.HttpUtility]::HtmlEncode($errorType))</td></tr>"
        Add-Content $FullLogFilePath '</table>'
        Add-Content $FullLogFilePath "<div class='info-after-error'><strong>Recommended action:</strong> $([System.Web.HttpUtility]::HtmlEncode($remediation))</div>"
        Add-Content $FullLogFilePath '<pre style="background-color:#f1f1f1; padding:12px; border-radius:4px; border-left:4px solid #c7254e; overflow-x:auto; font-size:13px;">'
        Add-Content $FullLogFilePath "$([System.Web.HttpUtility]::HtmlEncode($errorMessage))"
        Add-Content $FullLogFilePath '</pre>'
        Add-Content $FullLogFilePath '</div>'
    }

        Write-Host "`n🎯 Querying Exclaimer Add-in deployment..." -ForegroundColor Cyan
        $ProdID = "efc30400-2ac5-48b7-8c9b-c0fd5f266be2"
        $PreviewID = "a8d42ca1-6f1f-43b5-84e1-9ff40e967ccc"

        $user = $Global:userInput.Email
        $ProdResult = $null
        $PreviewResult = $null

        try {
            $ProdResult = Get-App -Identity "$user\$ProdID" -ErrorAction SilentlyContinue |
                Select-Object DisplayName, Enabled, AppVersion, Scope, Type
        } catch {}

        try {
            $PreviewResult = Get-App -Identity "$user\$PreviewID" -ErrorAction SilentlyContinue |
                Select-Object DisplayName, Enabled, AppVersion, Scope, Type
        } catch {}

        if ($ProdResult -or $PreviewResult) {
            Write-Host "`n✅ Exclaimer Add-in found:" -ForegroundColor Green

            if ($ProdResult) {
                Write-Host "`n--- Production Add-in ---" -ForegroundColor Cyan
                $ProdResult | Format-Table -AutoSize
            }
            if ($PreviewResult) {
                Write-Host "`n--- Preview Add-in ---" -ForegroundColor Cyan
                $PreviewResult | Format-Table -AutoSize
            }

            # --- HTML Logging (safe formatting) ---
            Add-Content $FullLogFilePath '<h3>🧩 Exclaimer Add-in Information (Admin)</h3>'
            Add-Content $FullLogFilePath '<table><tr><th>Type</th><th>Display Name</th><th>Version</th><th>Enabled</th><th>Scope</th><th title="Deployment method, see table below">Type</th></tr>'

            if ($ProdResult) {
                $enabledColor = if ($ProdResult.Enabled -ne $true) { ' style="color:red;font-weight:bold;"' } else { '' }

                $typeColor = switch ($ProdResult.Type) {
                    'PrivateCatalog' { ' style="color:orange;font-weight:bold;"' }
                    'Marketplace'   { ' style="color:red;font-weight:bold;"' }
                    default         { '' }
                }

                Add-Content $FullLogFilePath ('<tr><td>Production</td><td>{0}</td><td>{1}</td><td{5}>{2}</td><td>{3}</td><td{6} title="Deployment method, see table below">{4}</td></tr>' -f `
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.DisplayName),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.AppVersion),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.Enabled),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.Scope),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.Type),
                        $enabledColor,
                        $typeColor)
            }

            if ($PreviewResult) {
                $enabledColor = if ($PreviewResult.Enabled -ne $true) { ' style="color:red;font-weight:bold;"' } else { '' }

                $typeColor = switch ($PreviewResult.Type) {
                    'PrivateCatalog' { ' style="color:orange;font-weight:bold;"' }
                    'Marketplace'   { ' style="color:red;font-weight:bold;"' }
                    default         { '' }
                }

            Add-Content $FullLogFilePath ('<tr><td>Preview</td><td>{0}</td><td>{1}</td><td{5}>{2}</td><td>{3}</td><td{6} title="Deployment method, see table below">{4}</td></tr>' -f `
                    [System.Web.HttpUtility]::HtmlEncode($PreviewResult.DisplayName),
                    [System.Web.HttpUtility]::HtmlEncode($PreviewResult.AppVersion),
                    [System.Web.HttpUtility]::HtmlEncode($PreviewResult.Enabled),
                    [System.Web.HttpUtility]::HtmlEncode($PreviewResult.Scope),
                    [System.Web.HttpUtility]::HtmlEncode($PreviewResult.Type),
                    $enabledColor,
                    $typeColor)
            }

            Add-Content $FullLogFilePath '</table>'

            # Check for unexpected Type values in Production and Preview results
            $unexpectedTypes = @()
            if ($ProdResult -and ($ProdResult.Type -in 'Marketplace','PrivateCatalog')) {
                $unexpectedTypes += $ProdResult.Type
            }
            if ($PreviewResult -and ($PreviewResult.Type -in 'Marketplace','PrivateCatalog')) {
                $unexpectedTypes += $PreviewResult.Type
            }

            if ($unexpectedTypes.Count -gt 0) {
                Add-Content $FullLogFilePath '<div class="info-after-error"><strong>Warning: Unexpected deployment type detected.</strong> Please check the deployment "Version" method and Add-in "Type".</div>'
            }

            # Add attention note if either is not enabled
            $attentionMessages = @()

            if ($ProdResult -and $ProdResult.Enabled -ne $true) {
                $identity = "$user\$ProdID"
                $enableCommand = "Enable-App -Identity `"$identity`""
                $attentionMessages += "<span><b>ℹ️ Production Add-in is Disabled:</b><br>Run this command in PowerShell to re-enable it:</span> <code>$enableCommand</code>"
            }

            if ($PreviewResult -and $PreviewResult.Enabled -ne $true) {
                $identity = "$user\$PreviewID"
                $enableCommand = "Enable-App -Identity `"$identity`""
                $attentionMessages += "<span><b>ℹ️ Preview Add-in is Disabled:</b><br>Run this command in PowerShell to re-enable it:</span><code>$enableCommand</code>"
            }

            if ($attentionMessages.Count -gt 0) {
                $fullMessage = '<div class="info-after-error">' + ($attentionMessages -join "<br><br>") + '</div>'
                Add-Content -Path $FullLogFilePath -Value $fullMessage

                $sideNote = '<p class="side-note">If you have both Production and Preview versions deployed, only one needs to enabled.</p><p class="side-note">If you have reopened PowerShell, then you may need to run the command below before enabling the Add-in.</p><code>Connect-ExchangeOnline</code><p class="side-note">When an Add-in is disabled for a user, it should not appear or function in Outlook. We have observed cases where it may still load in Outlook on the web, but this is not expected behaviour. If this occurs, it may need to be raised with Microsoft for further review.</p>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            # --- Add explanatory table for deployment methods ---
            Add-Content $FullLogFilePath '<h4>Deployment Method Reference</h4>'
            Add-Content $FullLogFilePath '<table>'
            Add-Content $FullLogFilePath '<tr><th>Type</th><th>Deployment</th><th>Description</th></tr>'

            Add-Content $FullLogFilePath '<tr><td>MarketplacePrivateCatalog</td><td>AppSource (Private/Public)</td><td>Deployed centrally by an administrator using Microsoft AppSource or Centralized Deployment. Updates managed automatically by Microsoft.</td></tr>'
            Add-Content $FullLogFilePath '<tr><td>PrivateCatalog</td><td>Manifest (Custom XML)</td><td>Deployed manually by an admin using an uploaded manifest file. Typically used for preview or testing deployments.</td></tr>'
            Add-Content $FullLogFilePath '<tr><td>Marketplace</td><td>User Installed</td><td>Installed directly by an individual user through Outlook "Get Add-ins" store. Managed at the user level.</td></tr>'

            Add-Content $FullLogFilePath '</table>'

        }
        else {
            Write-Host "`n⚠️ No Exclaimer Add-ins found for this user." -ForegroundColor Yellow

            Add-Content $FullLogFilePath '<h3>Exclaimer Add-in Information (Admin)</h3>'
            Add-Content $FullLogFilePath '<table>'
            Add-Content $FullLogFilePath '<tr><th>Status</th><th>Details</th></tr>'

            Add-Content $FullLogFilePath (
                '<tr><td class="warning">None</td><td>No Exclaimer Add-ins were found for {0}. This is expected for Shared Mailboxes, but not for User Mailboxes.</td></tr>' -f
                [System.Web.HttpUtility]::HtmlEncode($user)
            )

            Add-Content $FullLogFilePath '</table>'
        }

        # --- Getting Mailbox Details ---
        Write-Host "`nCollecting mailbox configuration related to Outlook Add-ins..." -ForegroundColor Cyan
        try {
            $casMailbox = Get-CASMailbox -Identity $user -ErrorAction Stop | Select-Object EwsEnabled, EwsAllowOutlook

            $mailbox = Get-Mailbox -Identity $user -ErrorAction Stop |
                Select-Object Name,
                            UserPrincipalName,
                            PrimarySmtpAddress,
                            AccountDisabled,
                            IsShared,
                            HiddenFromAddressListsEnabled

            # Append EwsEnabled as a synthetic property
            $mailbox | Add-Member -NotePropertyName EwsEnabled -NotePropertyValue $casMailbox.EwsEnabled
            $mailbox | Add-Member -NotePropertyName EwsAllowOutlook -NotePropertyValue $casMailbox.EwsAllowOutlook

            # Separator row
            Add-Content $FullLogFilePath '<h3>Mailbox configuration (Admin)</h3>'
            Add-Content $FullLogFilePath '<table>'
            Add-Content $FullLogFilePath '<tr><th colspan="6">📬 Mailbox Details</th></tr>'
            Add-Content $FullLogFilePath '<tr><th>Property</th><th colspan="2">Value</th><th colspan="3">Notes</th></tr>'

            foreach ($prop in $mailbox.PSObject.Properties) {
                $name  = [System.Web.HttpUtility]::HtmlEncode($prop.Name)
                $raw   = $prop.Value
                $value = if ($null -eq $raw) { 'N/A' } else { [System.Web.HttpUtility]::HtmlEncode([string]$raw) }
                if (($prop.Name -eq 'EwsEnabled' -or $prop.Name -eq 'EwsAllowOutlook') -and $raw -eq $false) {
                    $value = '<span style="color:red;font-weight:bold;">{0}</span>' -f $value
                }
                $notes = ''

                switch ($prop.Name) {
                    'UserPrincipalName' {
                        $notes = 'Authentication identity used for modern authentication.'
                    }
                    'PrimarySmtpAddress' {
                        $notes = 'Primary email address of the mailbox.'
                    }
                    'IsShared' {
                        $notes = if ($raw) {
                            'Is a Shared mailbox.'
                        } else {
                            'Not a Shared Mailbox.'
                        }
                    }
                    'AccountDisabled' {
                        $notes = if ($raw) {
                            'Associated user account is disabled. Expected for shared mailboxes.'
                        } else {
                            'Associated user account is enabled.'
                        }
                    }
                    'HiddenFromAddressListsEnabled' {
                        $notes = if ($raw) {
                            'Mailbox is hidden. Can cause Classic Outlook Add-in to apply the user signature by default.'
                        } else {
                            'Mailbox is visible.'
                        }
                    }
                    'EwsEnabled' {
                        if ($Global:buildRequiresEws) {
                            $notes = if ($raw -eq $false) {
                                '❌ "EwsEnabled" is disabled for this mailbox. This build requires EWS — add-ins will not function.'
                            } elseif ($raw -eq $true) {
                                '✅ "EwsEnabled" is enabled for this mailbox. Required for this Outlook build.'
                            } elseif ($null -eq $raw) {
                                '⚠️ "EwsEnabled" is NULL (inheriting org setting). This build requires EWS — verify org setting is enabled.'
                            } else {
                                '⚠️ Unable to determine "EwsEnabled" mailbox state. Review manually.'
                            }
                        } else {
                            $notes = 'ℹ️ Not required. This Outlook build uses baseline security mode and does not depend on EWS.'
                        }
                    }
                    'EwsAllowOutlook' {
                        if ($Global:buildRequiresEws) {
                            $notes = if ($raw -eq $false) {
                                '❌ "EwsAllowOutlook" is disabled for this mailbox. This build requires EWS — Outlook add-ins will not function.'
                            } elseif ($raw -eq $true) {
                                '✅ "EwsAllowOutlook" is enabled for this mailbox. Required for this Outlook build.'
                            } elseif ($null -eq $raw) {
                                '⚠️ "EwsAllowOutlook" is NULL (inheriting org setting). This build requires EWS — verify org setting is enabled.'
                            } else {
                                '⚠️ Unable to determine "EwsAllowOutlook" mailbox state. Review manually.'
                            }
                        } else {
                            $notes = 'ℹ️ Not required. This Outlook build uses baseline security mode and does not depend on EWS.'
                        }
                    }
                    Default {
                        $notes = 'Informational.'
                    }
                }

                Add-Content $FullLogFilePath (
                    "<tr><td>{0}</td><td colspan='2'>{1}</td><td colspan='3'>{2}</td></tr>" -f
                    $name,
                    $value,
                    [System.Web.HttpUtility]::HtmlEncode($notes)
                )
            }

            Add-Content $FullLogFilePath '</table>'

            # EWS mailbox-level error notices
            if ($Global:buildRequiresEws -and $mailbox.EwsEnabled -eq $false) {
                $sideNote = '<div class="info-after-error"><span><b>❌ ''EwsEnabled'' is disabled on this mailbox:</b><br>' +
                    'Services relying on EWS (including certain Outlook Classic add-in scenarios) may not function for this mailbox.<br><br>' +
                    'Recommended action:<br>' +
                    ('<code>Set-CASMailbox -Identity "{0}" -EwsEnabled $true</code>' -f [System.Web.HttpUtility]::HtmlEncode($user)) +
                    '</span></div>' +
                    '<div class="info-after-note">' +
                        '<span>If you have reopened PowerShell, you may need to run first: ' +
                        '<code>Connect-ExchangeOnline</code></span><br><br>' +
                        '<span>Once this is completed, please re-run the full script again to verify the changes made.</span>' +
                    '</div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            if ($Global:buildRequiresEws -and $mailbox.EwsAllowOutlook -eq $false) {
                $sideNote = '<div class="info-after-error"><span><b>❌ ''EwsAllowOutlook'' is disabled on this mailbox:</b><br>' +
                    'Outlook is blocked from using EWS against this mailbox, which can break add-in scenarios that depend on EWS.<br><br>' +
                    'Recommended action:<br>' +
                    ('<code>Set-CASMailbox -Identity "{0}" -EwsAllowOutlook $true</code>' -f [System.Web.HttpUtility]::HtmlEncode($user)) +
                    '</span></div>' +
                    '<div class="info-after-note">' +
                        '<span>If you have reopened PowerShell, you may need to run first: ' +
                        '<code>Connect-ExchangeOnline</code></span><br><br>' +
                        '<span>Once this is completed, please re-run the full script again to verify the changes made.</span>' +
                    '</div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            if ($Global:buildRequiresEws -and $null -eq $mailbox.EwsEnabled) {
                $sideNote = '<div class="info-after-warning"><span><b>ℹ️ ''EwsEnabled'' is not explicitly set on this mailbox:</b><br>' +
                    'EWS is used by <b>Outlook Classic (Windows)</b> for certain add-in scenarios.<br><br>' +
                    'When not explicitly set at mailbox level, behaviour falls back to the organization-level setting and could impact add-in functionality.<br><br>' +
                    'Recommended action:<br>' +
                    ('<code>Set-CASMailbox -Identity "{0}" -EwsEnabled $true</code>' -f [System.Web.HttpUtility]::HtmlEncode($user)) + '<br><br>' +
                    'Note: Mailbox-level EWS settings override the organization setting when explicitly configured.' +
                    '</span></div>' +
                    '<div class="info-after-note">' +
                        '<span>If you have reopened PowerShell, you may need to run first: <code>Connect-ExchangeOnline</code></span><br><br>' +
                        '<span>Once this is completed, please re-run the full script again to verify the changes made.</span>' +
                    '</div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            if ($Global:buildRequiresEws -and $null -eq $mailbox.EwsAllowOutlook) {
                $sideNote = '<div class="info-after-warning"><span><b>ℹ️ ''EwsAllowOutlook'' is not explicitly set on this mailbox:</b><br>' +
                    'This setting controls whether Outlook clients can access EWS, particularly in <b>Outlook Classic (Windows)</b>.<br><br>' +
                    'When not explicitly set at mailbox level, behaviour falls back to the organization-level setting and could impact add-in functionality.<br><br>' +
                    'Recommended action:<br>' +
                    ('<code>Set-CASMailbox -Identity "{0}" -EwsAllowOutlook $true</code>' -f [System.Web.HttpUtility]::HtmlEncode($user)) + '<br><br>' +
                    'Note: Mailbox-level EWS settings override the organization setting when explicitly configured.' +
                    '</span></div>' +
                    '<div class="info-after-note">' +
                        '<span>If you have reopened PowerShell, you may need to run first: <code>Connect-ExchangeOnline</code></span><br><br>' +
                        '<span>Once this is completed, please re-run the full script again to verify the changes made.</span>' +
                    '</div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            # Bulk remediation suggestion if either EWS setting is not explicitly TRUE (only relevant for builds that require EWS)
            if ($Global:buildRequiresEws -and ($mailbox.EwsEnabled -eq $false -or $mailbox.EwsAllowOutlook -eq $false)) {
                $sideNote = '<div class="info-after-warning"><span><b>ℹ️ Bulk remediation across all mailboxes:</b><br>' +
                    'To ensure both <code>EwsEnabled</code> and <code>EwsAllowOutlook</code> are explicitly set to <b>TRUE</b> for every mailbox in the tenant, the Admin can run:<br><br>' +
                    '<code>Get-Mailbox -ResultSize Unlimited | Set-CASMailbox -EwsEnabled $true -EwsAllowOutlook $true</code><br><br>' +
                    'Note: This applies the change to all mailboxes — review scope before running in production.' +
                    '</span></div>'
                Add-Content -Path $FullLogFilePath -Value $sideNote
            }

            # Compare UPN and Primary SMTP
                $upn  = [string]$mailbox.UserPrincipalName
                $smtp = [string]$mailbox.PrimarySmtpAddress

                $upn  = $upn.Trim()  -replace "`r|`n",""
                $smtp = $smtp.Trim() -replace "`r|`n",""

                if ($upn -and $smtp -and ($upn.ToLower() -ne $smtp.ToLower())) {

                    $articleUrl = 'https://learn.microsoft.com/en-us/windows-server/identity/ad-fs/operations/configuring-alternate-login-id'

                    Add-Content $FullLogFilePath (
                        "<div class='info-after-warning'>
                            ⚠️ UPN and Primary SMTP address do not match.<br><br>
                            UPN: <code>{0}</code><br>
                            SMTP: <code>{1}</code><br><br>
                            This can contribute to modern authentication token mismatches and login hint errors in Outlook.
                            Review alternate login ID guidance here:<br>
                            <a href='{2}' target='_blank'>{2}</a>
                        </div>" -f
                        [System.Web.HttpUtility]::HtmlEncode($upn),
                        [System.Web.HttpUtility]::HtmlEncode($smtp),
                        $articleUrl
                    )
                }
                else {
                    Add-Content $FullLogFilePath "<div class='info-after-success'>✔ UPN and Primary SMTP address match.</div>"
                }
        }
    catch {
        Add-Content $FullLogFilePath '<table>'
        Add-Content $FullLogFilePath '<tr><th colspan="6">📬 Mailbox Details</th></tr>'
        Add-Content $FullLogFilePath '<tr><td colspan="6" class="warning">Unable to retrieve mailbox details for this user. Check permissions or mailbox existence.</td></tr>'
        Add-Content $FullLogFilePath '</table>'
    }

        try {
            # Disconnect if needed
            #Disconnect-ExchangeOnline -Confirm:$false | Out-Null
            #Write-Host "`n🔒 Disconnected from Exchange Online." -ForegroundColor DarkGray
        } catch {}
    }
    else {
        Add-Content $FullLogFilePath '<div class="info-after-warning"><strong>Exchange Online connection failed or cancelled by user.</strong></div>'
    }
}
else {
    Add-Content $FullLogFilePath '<div class="info-after-warning"><strong>Exchange Online module not available. Manual Add-in version collection required.</strong></div>'
}
Write-Host "`n✅ Exclaimer Add-in details collection completed." -ForegroundColor Green
}
InpectEXOconfiguration
function GetFirewallLogs {
    Write-Host "`n========== Windows Firewall Logging ==========" -ForegroundColor Cyan

    # --- Elevation check ---
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {

        Write-Host "Administrator privileges are required to collect firewall logs." -ForegroundColor Red
        Write-Host "Please re-run PowerShell using 'Run as administrator'." -ForegroundColor Yellow

        Add-Content $FullLogFilePath '<div class="section">'
        Add-Content $FullLogFilePath '<h2>🔥 Windows Firewall Log Capture</h2>'
        Add-Content $FullLogFilePath '<table>'
        Add-Content $FullLogFilePath '<tr><th>Property</th><th>Value</th></tr>'
        Add-Content $FullLogFilePath '<tr><td><strong>Log capture performed</strong></td><td style="color:#F5A627;font-weight:bold;">No</td></tr>'
        Add-Content $FullLogFilePath '<tr><td><strong>Reason</strong></td><td>PowerShell was not run as Administrator.</td></tr>'
        Add-Content $FullLogFilePath '</table>'
        Add-Content $FullLogFilePath '</div>'

        return
    }

    $firewallLogPath = "C:\Windows\System32\LogFiles\Firewall\pfirewall.log"
    $issueReproduced = $false
    $destination = $null

    Write-Host "This will temporarily enable Windows Firewall logging to capture Outlook traffic." -ForegroundColor Yellow
    $choice = Read-Host "Do you want to continue? (Y/N)"

    if ($choice -notmatch '^[Yy]$') {

        Write-Host "Firewall logging step skipped by user." -ForegroundColor Yellow

        Add-Content $FullLogFilePath '<div class="section">'
        Add-Content $FullLogFilePath '<h2>🔥 Windows Firewall Log Capture</h2>'
        Add-Content $FullLogFilePath '<table>'
        Add-Content $FullLogFilePath '<tr><th>Property</th><th>Value</th></tr>'
        Add-Content $FullLogFilePath '<tr><td><strong>Log capture performed</strong></td><td style="color:#F5A627;font-weight:bold;">No</td></tr>'
        Add-Content $FullLogFilePath '<tr><td><strong>Reason</strong></td><td>User declined to enable firewall logging.</td></tr>'
        Add-Content $FullLogFilePath '</table>'
        Add-Content $FullLogFilePath '</div>'

        return
    }

    try {

        Write-Host "`nEnabling Windows Firewall logging..." -ForegroundColor White

        Set-NetFirewallProfile -Profile Domain,Private,Public `
            -LogAllowed True `
            -LogBlocked True `
            -LogFileName $firewallLogPath `
            -LogMaxSizeKilobytes 32767

        Write-Host "Firewall logging is now enabled." -ForegroundColor Green
        Write-Host "`nPlease reproduce the reported issue now (Classic Outlook, New Outlook, or Outlook on the Web)." -ForegroundColor Yellow
        Write-Host "This should ensure that any relevant logs are captured for review.`n" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  1) Close Outlook."
        Write-Host "  2) Start Outlook again."
        Write-Host "  3) Reproduce the issue reported."
        Write-Host "  4) Once you reprocuded the issue reported, close Outlook."
        Write-Host ""
        Write-Host "`nPress ENTER once the issue has been reproduced to continue..." -ForegroundColor Yellow
        Read-Host

        $confirmed = Read-Host "Was the issue successfully reproduced? (Y/N)"
        if ($confirmed -match '^[Yy]$') {
            $issueReproduced = $true
        }

        if ($issueReproduced -and (Test-Path $firewallLogPath)) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $destination = Join-Path $Global:FilePath "pfirewall_$timestamp.log"
            Copy-Item -Path $firewallLogPath -Destination $destination -Force
        }

    }
    finally {

        Write-Host "Disabling Windows Firewall logging..." -ForegroundColor White

        Set-NetFirewallProfile -Profile Domain,Private,Public `
            -LogAllowed False `
            -LogBlocked False

        Add-Content $FullLogFilePath '<div class="section">'
        Add-Content $FullLogFilePath '<h2>🔥 Windows Firewall Log Capture</h2>'
        Add-Content $FullLogFilePath '<table>'
        Add-Content $FullLogFilePath '<tr><th>Property</th><th>Value</th></tr>'
        Add-Content $FullLogFilePath "<tr><td><strong>Issue reproduced</strong></td><td>$issueReproduced</td></tr>"

        if ($destination) {
            Add-Content $FullLogFilePath "<tr><td><strong>Log capture performed</strong></td><td>Yes</td></tr>"
            Add-Content $FullLogFilePath ("<tr><td><strong>Log file name</strong></td><td>{0}</td></tr>" -f (Split-Path $destination -Leaf))
            Add-Content $FullLogFilePath ("<tr><td><strong>Log file location</strong></td><td>{0}</td></tr>" -f (Split-Path $destination -Parent))
        }
        else {
            Add-Content $FullLogFilePath '<tr><td><strong>Log capture performed</strong></td><td style="color:#F5A627;font-weight:bold;">No</td></tr>'
            Add-Content $FullLogFilePath '<tr><td><strong>Notes</strong></td><td>Issue was not reproduced during the capture window.</td></tr>'
        }

        Add-Content $FullLogFilePath '</table>'
        Add-Content $FullLogFilePath '</div>'

        Write-Host "Firewall logging has been disabled." -ForegroundColor Green
        Write-Host "Firewall log collection complete." -ForegroundColor Cyan
        $Global:destination = $destination
    }
}
if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)','Outlook Web (Windows)','Multiple / All')) {
    GetFirewallLogs
}


function GetSupportSubmissionInstructions {
    Write-Host "`n========== Support Submission Instructions ==========" -ForegroundColor Cyan
    Write-Host "Please provide the following files to the support team:" -ForegroundColor White
    Write-Host ""
    Write-Host "1. Diagnostic report file:" -ForegroundColor Yellow
    Write-Host "   $FullLogFilePath" -ForegroundColor White
    Write-Host ""

    if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)', 'Outlook Web (Windows)', 'Multiple / All')) {
        Write-Host "2. Windows Firewall log file:" -ForegroundColor Yellow
        if ($Global:destination) {
            Write-Host "   $Global:destination" -ForegroundColor White
        } else {
            Write-Host "   Not collected." -ForegroundColor Yellow
        }
        Write-Host ""
    }

    Write-Host "Attach the file(s) listed above to your support ticket for review." -ForegroundColor Green

    # --- HTML Logging ---
    Add-Content $FullLogFilePath '<div class="section">'
    Add-Content $FullLogFilePath '<h2>📩 Support Submission Instructions</h2>'
    Add-Content $FullLogFilePath '<table>'
    Add-Content $FullLogFilePath '<tr><th>File</th><th>Details</th></tr>'

    Add-Content $FullLogFilePath (
        "<tr><td><strong>Diagnostic report</strong></td><td>{0}</td></tr>" -f $FullLogFilePath
    )

    if ($Global:userInput.OutlookAffected -in @('Classic Outlook (Windows)', 'New Outlook (Windows)', 'Outlook Web (Windows)', 'Multiple / All')) {
        $firewallLogValue = if ($Global:destination) { $Global:destination } else { 'Not collected' }
        Add-Content $FullLogFilePath (
            "<tr><td><strong>Windows Firewall log</strong></td><td>{0}</td></tr>" -f $firewallLogValue
        )
    }

    Add-Content $FullLogFilePath '</table>'
    Add-Content $FullLogFilePath '<p>Please attach the file(s) listed above when contacting the support team so they can review the collected data.</p>'
    Add-Content $FullLogFilePath '</div>'
}
GetSupportSubmissionInstructions

Write-Host "`n========================================="
Write-Host "  Script completed successfully." -ForegroundColor Green
Write-Host "  Log file location:'$FullLogFilePath'"
Write-Host "=========================================`n"

Add-Content -Path $FullLogFilePath -Value @"
<div class='section'>
  <h2>📄 Output Log Location</h2>
  <p>This report has been saved to: <code>$FullLogFilePath</code></p>
</div>
"@

@"
<button id="scrollToWarning" class="floating-button-warning">Go to Warning</button>
<button id="scrollToError" class="floating-button-error">Go to Error</button>
<script>
function initScrollButton(buttonId, elementSelector) {
    const elements = document.querySelectorAll(elementSelector);
    const button = document.getElementById(buttonId);
    
    if (elements.length === 0) {
        button.style.display = 'none';
        return;
    }
    
    let currentIndex = 0;
    const scrollOffset = 0;
    
    button.addEventListener('click', function() {
        const element = elements[currentIndex];
        let targetElement = findNearestHeading(element);
        
        const targetY = window.scrollY + targetElement.getBoundingClientRect().top - scrollOffset;
        window.scrollTo({ top: Math.max(0, targetY), behavior: 'smooth' });
        currentIndex = (currentIndex + 1) % elements.length;
    });
}

function findNearestHeading(element) {
    // Find the nearest previous h2 or h3
    let sibling = element.previousElementSibling;
    while (sibling) {
        if (sibling.tagName === 'H2' || sibling.tagName === 'H3') {
            return sibling;
        }
        sibling = sibling.previousElementSibling;
    }
    
    // If no heading found in siblings, look up the tree (max 10 levels)
    let parent = element.parentElement;
    let depth = 0;
    while (parent && depth < 10) {
        if (parent.tagName === 'H2' || parent.tagName === 'H3') {
            return parent;
        }
        parent = parent.parentElement;
        depth++;
    }
    
    return element;
}

// Initialize both scroll buttons
initScrollButton('scrollToWarning', '.info-after-warning');
initScrollButton('scrollToError', '.info-after-error');
</script>
</div>
</body>
</html>
"@ | Add-Content -Path $FullLogFilePath -Encoding UTF8

# Open the file for user to view immediately (optional)
Start-Process $FullLogFilePath