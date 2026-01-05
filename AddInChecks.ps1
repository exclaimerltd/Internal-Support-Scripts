# <#
# .SYNOPSIS
#     Gathers diagnostics and configuration data relevant to Exclaimer Add-In and signature deployment across Outlook clients.
#
# .DESCRIPTION
#     This script collects diagnostics including Outlook client versions, signature agent installations, WebView2 runtime presence, endpoint connectivity, local signatures, and cloud geolocation.
#     It prompts for a user's email to determine their domain, fetches relevant endpoints, tests connectivity, inspects Outlook installations, and produces an HTML report "AddInChecks.html" in the user’s Downloads folder.
#
# .NOTES
#     Email: helpdesk@exclaimer.com
#     Date: 23rd September 2025
#     Version: 1.0.0
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
#     1.1.0
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

# Check if script is running in Windows PowerShell 5.x
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) {
    Write-Host "This script must be run in Windows PowerShell 5.x." -ForegroundColor Red
    Write-Host "Please run this script in Windows PowerShell version 5.x." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to close"
    return  # Stop script execution, but do not exit the PowerShell session
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
        .side-note { color: #555; font-size: 12px; margin-top: 5px; font-style: italic; }
        code { background-color: #f1f1f1; padding: 2px 4px; border-radius: 4px; font-weight: bold; color: #c7254e; }
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

#It should set the below:
$ProdID = "efc30400-2ac5-48b7-8c9b-c0fd5f266be2"
$PreviewID = "a8d42ca1-6f1f-43b5-84e1-9ff40e967ccc"

function ConfirmElevationStatus {

    Write-Host "========== Script Permission Check ==========`n" -ForegroundColor Cyan

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
        Add-Content $FullLogFilePath '<tr><td><strong>Administrator privileges</strong></td><td>Yes</td></tr>'
        Add-Content $FullLogFilePath '<tr><td><strong>Impact</strong></td><td>All diagnostics can be collected.</td></tr>'
    }
    else {
        Add-Content $FullLogFilePath '<tr><td><strong>Administrator privileges</strong></td><td style="color:#F5A627;font-weight:bold;">No</td></tr>'
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
            Network         = $null
        }

        # --- 1) Email (validated) ---
        while ($true) {
            $email = Read-Host "`nEnter the user's email address (e.g. user@company.com)"
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
                Write-Host "  4) Outlook Mobile"
                Write-Host "  5) Multiple / All"
                $oChoice = Read-Host "`nEnter choice (1-5)"
            } while ($oChoice -notmatch '^[1-5]$')

            switch ($oChoice) {
                1 { $Global:userInput.OutlookAffected = 'Classic Outlook' }
                2 { $Global:userInput.OutlookAffected = 'New Outlook' }
                3 { $Global:userInput.OutlookAffected = 'Outlook Web' }
                4 { $Global:userInput.OutlookAffected = 'Outlook Mobile' }
                5 { $Global:userInput.OutlookAffected = 'Multiple / All' }
            }

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
            Write-Host ("Network Scope:    {0}" -f $Global:userInput.Network) -ForegroundColor White
        }

        Write-Host ""
        do {
            $confirm = Read-Host "Is the information correct? (Y/N) [Y]"
            if ([string]::IsNullOrWhiteSpace($confirm)) { $confirm = 'Y' }
            $confirm = $confirm.Substring(0,1).ToUpper()
        } while ($confirm -notin @('Y','N'))

        if ($confirm -eq 'Y') {
            # ✅ Write to HTML log *only once, after confirmation*
            Add-Content $FullLogFilePath "<div class='section'>"
            Add-Content $FullLogFilePath "<h2>🧾 User Input Summary</h2>"
            Add-Content $FullLogFilePath "<table>"
            Add-Content $FullLogFilePath "<tr><td><strong>Purpose:</strong></td><td>$($Global:userInput.Purpose)</td></tr>"
            Add-Content $FullLogFilePath "<tr><td><strong>Email:</strong></td><td>$($Global:userInput.Email)</td></tr>"

            if ($Global:userInput.Purpose -eq 'Troubleshooting') {
                Add-Content $FullLogFilePath "<tr><td><strong>Users Affected:</strong></td><td>$($Global:userInput.UsersAffected)</td></tr>"
                Add-Content $FullLogFilePath "<tr><td><strong>Outlook Affected:</strong></td><td>$($Global:userInput.OutlookAffected)</td></tr>"
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

# --- Script execution ---

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
            $result = Test-NetConnection -ComputerName $endpoint -Port 443 -InformationLevel Quiet
        }

        $status = if ($result) { "Success" } else { "Failed" }

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

    # --- Supported build thresholds ---
    $supportedBuilds = @(
        [PSCustomObject]@{ MinBuild = 26100; Status = '✅ Supported'; Note = 'Windows 11 24H2 or later build.' },
        [PSCustomObject]@{ MinBuild = 22631; Status = '✅ Supported'; Note = 'Windows 11 23H2 or later build.' },
        [PSCustomObject]@{ MinBuild = 22621; Status = '✅ Supported'; Note = 'Windows 11 22H2 build.' },
        [PSCustomObject]@{ MinBuild = 19045; Status = '✅ Supported'; Note = 'Windows 10 22H2 build (supported until October 2025).' }
    )

    # --- Collect OS Info ---
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $caption = $os.Caption
    $version = $os.Version
    $build   = [int]$os.BuildNumber

    # --- Default to unsupported ---
    $supportStatus = '❌ Unsupported or legacy Windows version.'
    $supportNote   = 'Consider upgrading to Windows 10 22H2 (bild 19045 or above) or Windows 11 for compatibility.'

    # --- Determine if build is supported ---
    foreach ($entry in $supportedBuilds) {
        if ($build -ge $entry.MinBuild) {
            $supportStatus = $entry.Status
            $supportNote   = $entry.Note
            break
        }
    }

    # --- Console Output ---
    Write-Host "Windows Version: $caption ($version)" -ForegroundColor White
    Write-Host "Build Number:    $build" -ForegroundColor White
    Write-Host "Support Status:  $supportStatus" -ForegroundColor Yellow
    Write-Host "Note:            $supportNote" -ForegroundColor DarkGray

    # --- HTML Logging (safe, no '+' concat) ---
    Add-Content $FullLogFilePath ("<tr><td><strong>Windows Version</strong></td><td>{0} ({1})</td></tr>" -f $caption, $version)
    Add-Content $FullLogFilePath ("<tr><td><strong>Build Number</strong></td><td>{0}</td></tr>" -f $build)
    Add-Content $FullLogFilePath ("<tr><td><strong>Support Status</strong></td><td>{0}</td></tr>" -f $supportStatus)
    Add-Content $FullLogFilePath ("<tr><td><strong>Notes</strong></td><td>{0}</td></tr>" -f $supportNote)
    Add-Content $FullLogFilePath '</table>'
    Add-Content $FullLogFilePath '</div>'
}
GetWindowsVersion

function GetWindowsNetworkDetails {
    Write-Host "`n========== Network Connection Details ==========" -ForegroundColor Cyan

    # --- HTML Section Header ---
    Add-Content $FullLogFilePath '<div class="section">'
    Add-Content $FullLogFilePath '<h2>🌐 Network Connection Details</h2>'
    Add-Content $FullLogFilePath '<table>'
    Add-Content $FullLogFilePath '<tr><th>Interface</th><th>Network Name</th><th>Category</th><th>IPv4 Connectivity</th><th>IPv6 Connectivity</th></tr>'

    # --- Collect Network Profiles ---
    $profiles = Get-NetConnectionProfile

    if (-not $profiles) {
        Write-Host "No active network connections found." -ForegroundColor Yellow
        Add-Content $FullLogFilePath '<tr><td colspan="5">No active network connections detected.</td></tr>'
    }
    else {
        foreach ($profile in $profiles) {
            $interfaceAlias = $profile.InterfaceAlias
            $networkName    = if ($profile.Name) { $profile.Name } else { 'N/A' }
            $category       = $profile.NetworkCategory
            $ipv4           = $profile.IPv4Connectivity
            $ipv6           = $profile.IPv6Connectivity

            # --- Console Output ---
            Write-Host "Interface:        $interfaceAlias" -ForegroundColor White
            Write-Host "Network Name:     $networkName" -ForegroundColor White
            Write-Host "Category:         $category" -ForegroundColor White
            Write-Host "IPv4 Connectivity $ipv4" -ForegroundColor DarkGray
            Write-Host "IPv6 Connectivity $ipv6`n" -ForegroundColor DarkGray

            # --- HTML Logging ---
            Add-Content $FullLogFilePath (
                "<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>" -f `
                $interfaceAlias, $networkName, $category, $ipv4, $ipv6
            )
        }
    }

    Add-Content $FullLogFilePath '</table>'
    Add-Content $FullLogFilePath '</div>'
}

GetWindowsNetworkDetails

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

    $installedSummary = "<ul>"
    if ($classicInstalled -and $newOutlookInstalled) {
        Write-Host "Both Classic Outlook and New Outlook are installed."
        $installedSummary += "<li><span class='success'>Both Classic and New Outlook are installed</span></li>"
    } elseif ($classicInstalled) {
        Write-Host "Only Classic Outlook is installed."
        $installedSummary += "<li><span class='success'>Only Classic Outlook is installed</span></li>"
    } elseif ($newOutlookInstalled) {
        Write-Host "Only New Outlook is installed."
        $installedSummary += "<li><span class='success'>Only New Outlook is installed</span></li>"
    } else {
        Write-Host "No Outlook installation detected."
        $installedSummary += "<li><span class='fail'>No Outlook installation detected</span></li>"
    }

    $installedSummary += "</ul>"
    Add-Content $FullLogFilePath $installedSummary
    

    if ($newOutlookEnabled) {
        Write-Host "New Outlook is installed, and the toggle is ON (New Outlook is Default)." -ForegroundColor Yellow
        Add-Content $FullLogFilePath "<ul><span class='info-after-note'>New Outlook is installed, and the toggle is ON (New Outlook is Default).</span></ul>"
    } else {
        Write-Host "New Outlook is installed, but the toggle is OFF (Classic Outlook is Default)." -ForegroundColor Yellow
        Add-Content $FullLogFilePath "<ul><span class='info-after-note'>New Outlook is installed, but the toggle is OFF (Classic Outlook is Default)</span></ul>"
    }

    if ($newOutlookInstalled) {
        Add-Content $FullLogFilePath "<h3>New Outlook</h3><ul>"
        $newOutlookVersion = Get-NewOutlookVersion
        Write-Host "`n========== New Outlook Information ==========" -ForegroundColor Cyan
        if ($newOutlookVersion) {
            Write-Host "New Outlook Version: $newOutlookVersion"
            Add-Content $FullLogFilePath "<li>Version: $newOutlookVersion</li>"
        } else {
            Write-Host "New Outlook version could not be determined." -ForegroundColor Yellow
            Add-Content $FullLogFilePath "<li><span class='warning'>Version could not be determined</span></li>"
        }
        Add-Content $FullLogFilePath "</ul>"
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

            $map = @{
            # Current Channel
            "19426.20186"="2511"; "19426.20170"="2511";
            "19328.20266"="2510"; "19328.20244"="2510"; "19328.20232"="2510"; "19328.20190"="2510"; "19328.20178"="2510"; "19328.20158"="2510";
            "19231.20274"="2509"; "19231.20246"="2509"; "19231.20216"="2509"; "19231.20194"="2509"; "19231.20172"="2509"; "19231.20156"="2509";
            "19127.20402"="2508"; "19127.20384"="2508"; "19127.20358"="2508"; "19127.20314"="2508"; "19127.20302"="2508"; "19127.20264"="2508"; "19127.20240"="2508"; "19127.20222"="2508";
            "19029.20300"="2507"; "19029.20274"="2507"; "19029.20208"="2507"; "18925.20268"="2506"; "18925.20242"="2506"; "18827.20244"="2505";

            # Monthly Enterprise Channel overlap
            "18730.20260"="2504"; "18730.20240"="2504"; "18623.20316"="2503"; "18623.20208"="2503"; "18623.20178"="2503"; "18623.20156"="2503";
            "18526.20672"="2502"; "18526.20660"="2502"; "18526.20634"="2502"; "18526.20546"="2502"; "18526.20472"="2502"; "18526.20438"="2502"; "18526.20416"="2502"; "18526.20336"="2502"; "18526.20264"="2502"; "18526.20168"="2502"; "18526.20144"="2502";

            # Early Semi‑Annual Enterprise / Channels
            "18429.20240"="2501"; "18429.20158"="2501"; "18429.20132"="2501";
            "18324.20272"="2412"; "18324.20194"="2412"; "18324.20190"="2412"; "18324.20168"="2412";
            "18227.20240"="2411"; "18227.20162"="2411"; "18227.20152"="2411";
            "18129.20242"="2410"; "18129.20200"="2410"; "18129.20158"="2410";
            "18025.20242"="2409"; "18025.20214"="2409"; "18025.20160"="2409"; "18025.20140"="2409"; "18025.20104"="2409"; "18025.20096"="2409";
            "17928.20742"="2408"; "17928.20730"="2408"; "17928.20708"="2408"; "17928.20588"="2408"; "17928.20572"="2408"; "17928.20538"="2408"; "17928.20512"="2408"; "17928.20392"="2408"; "17928.20336"="2408"; "17928.20286"="2408"; "17928.20156"="2408"; "17928.20114"="2408";

            # Older versions through 2202
            "17726.20222"="2406"; "17726.20160"="2406"; "17726.20126"="2406";
            "17628.20206"="2405"; "17628.20188"="2405"; "17628.20164"="2405"; "17628.20152"="2405"; "17628.20144"="2405"; "17628.20110"="2405";
            "17531.20210"="2404"; "17531.20190"="2404"; "17531.20152"="2404"; "17531.20140"="2404"; "17531.20128"="2404"; "17531.20120"="2404";
            "17425.20258"="2403"; "17425.20176"="2403"; "17425.20146"="2403"; "17425.20138"="2403";
            "17328.20414"="2402"; "17328.20346"="2402"; "17328.20336"="2402"; "17328.20282"="2402"; "17328.20184"="2402"; "17328.20162"="2402"; "17328.20142"="2402";

            # Versions around 2202 (older)
            "16130.20990"="2302"; "16130.20964"="2302"; "16130.20928"="2302"; "16130.20888"="2302"; "16130.20858"="2302"; "16130.20848"="2302"; "16130.20772"="2302"; "16130.20766"="2302"; "16130.20738"="2302"; "16130.20724"="2302"; "16130.20580"="2302"; "16130.20500"="2302"; "16130.20394"="2302"; "16130.20346"="2302"; "16130.20282"="2302"; "16130.20184"="2302";
            "15601.20870"="2208"; "15601.20848"="2208"; "15601.20832"="2208"; "15601.20796"="2208"; "15601.20772"="2208"; "15601.20680"="2208"; "15601.20660"="2208"; "15601.20578"="2208"; "15601.20456"="2208"; "15601.20378"="2208"; "15601.20286"="2208"; "15601.20088"="2208";
            "15225.20422"="2205"; "15225.20394"="2205"; "15225.20288"="2205"; "15225.20204"="2205";
            "15128.20312"="2204"; "15128.20280"="2204"; "15128.20248"="2204"; "15128.20178"="2204";
            "15028.20204"="2203"; "15028.20160"="2203";
            "14931.20724"="2202"; "14931.20660"="2202"; "14931.20494"="2202"; "14931.20392"="2202"; "14931.20274"="2202"; "14931.20132"="2202"; "14931.20120"="2202";
            }

            if ($officeBuild) {
                $outlookVersion = $map[$officeBuild]
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
        <td>$outlookVersion</td>
        <td>$outlookBitness</td>
        <td>$licenseType</td>
        <td>$buildSupport</td>
    </tr>
</table>
"@
        Add-Content $FullLogFilePath $classicOutlookTable
        }
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
                Add-Content $FullLogFilePath "<h3>Local Outlook Signatures</h3>"
                Add-Content $FullLogFilePath "<table><tr><th>Name</th><th>Date Modified</th><th>Exclaimer Signature</th></tr>"

                foreach ($sig in $signatureData) {
                    $sigRow = "<tr><td>$($sig.Name)</td><td>$($sig.DateModified)</td><td>$($sig.Exclaimer)</td></tr>"
                    Add-Content $FullLogFilePath $sigRow
                }

                Add-Content $FullLogFilePath "</table>"

            } else {
                Write-Host "`nNo .htm signature files found in $signaturePath" -ForegroundColor DarkGray
                Add-Content $FullLogFilePath "<p>✅ No local signature files found in $signaturePath</p>"
            }
        }

    }
    
}

InspectOutlookConfiguration

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
    }
    else {
        Write-Host "The Exclaimer Cloud Signature Update Agent is not installed." -ForegroundColor Yellow
        Add-Content $FullLogFilePath "<p>✅ The Exclaimer Cloud Signature Update Agent is not installed.</p>"
    }

    Add-Content $FullLogFilePath "</div>"

    # Define registry paths for 64-bit and 32-bit uninstall keys
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
    )
    Write-Host "`n--- Word File Block Settings Check (Web Pages) ---" -ForegroundColor Yellow
    Add-Content $FullLogFilePath "<h3>Word File Block Settings (Web Pages)</h3>"

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

        foreach ($check in $registryChecks) {
            if (Test-Path $check.Path) {
                $key = Get-ItemProperty -Path $check.Path -ErrorAction SilentlyContinue
                if ($key.HtmlFiles -eq 1 -or $key.HtmlFiles -eq 2) {
                    Write-Host "Web Pages BLOCKED via $($check.Scope)" -ForegroundColor Red
                    Add-Content $FullLogFilePath "<p>❌ Web Pages are <strong>blocked</strong>.</p>"
                    $webPagesBlocked = $true
                }
                elseif ($key.HtmlFiles -eq 0) {
                    Write-Host "Web Pages allowed via $($check.Scope)" -ForegroundColor Green
                    Add-Content $FullLogFilePath "<p>✅ Web Pages allowed via $($check.Scope).</p>"
                }
            }
        }

    if ($webPagesBlocked) {
        Write-Host "`nWARNING: Word is blocking Web Pages. HTML based signatures cannot be inserted into the Outlook email body." -ForegroundColor Red
        Write-Host "This setting must be disabled to allow signature injection." -ForegroundColor Red

Add-Content $FullLogFilePath @"
<p style='color:red;'>
<strong>Impact:</strong> Word is blocking Web Page file types (.htm/.html).  
This prevents Exclaimer signatures from being added to the Outlook message body.
</p>
<p>
<strong>Where to find this setting:</strong><br>
Microsoft Word &gt; File &gt; Options &gt; Trust Center &gt; Trust Center Settings &gt; File Block Settings &gt; Web Pages
</p>
<p>
<strong>Note on the setting:</strong><br>
If the box for Web Pages is <strong>checked</strong>, the file type is blocked.  
The desired state for signatures to work is <strong>unchecked</strong>.
</p>
<p>
<strong>Why this is required:</strong><br>
Outlook uses Word as the email editor, and File Block / Trust Center settings are read at application startup.  
If this setting was recently changed or applied by policy, Outlook must be restarted for the change to take effect.
</p>
<p>
<strong>Note:</strong><br>
If this option is selected but greyed out, the setting is being enforced by Group Policy (GPO) and cannot be changed locally by the user.  
In managed environments, an administrator would need to review the policy controlling Word File Block settings.
</p>
"@
    }
    else {
        Write-Host "No blocking detected for Web Pages." -ForegroundColor Green
        Add-Content $FullLogFilePath "<p>✅ No Word File Block restrictions detected for Web Pages.</p>"
    }

    Write-Host "`n========== Microsoft Edge WebView2 Runtime ==========" -ForegroundColor Cyan

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

# -------------------------------------------------------------------
# 📨 EXCLAIMER ADD-IN DETAILS COLLECTION (User or Admin)
# -------------------------------------------------------------------

Write-Host ""
Write-Host "=== Exclaimer Add-in Information ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "ℹ️  If a Microsoft 365 Global Admin is available, selecting 'Y' on the next prompt allows the script to collect important Exclaimer Add-in details." -ForegroundColor Yellow
Write-Host "This includes deployment information and the current State of the Add-in for the user reporting issues." -ForegroundColor Cyan
Write-Host "`nIf you do not run the next step as a Global Admin, we may need to ask you to run some PowerShell commands manually to collect the required information." -ForegroundColor Red
Write-Host "`nRecommended action: 'Y' continue as a Microsoft 365 Global Administrator to collect full details.`n" -ForegroundColor Cyan

# --- Step: Check if user is Global Admin ---
$adminChoice = Read-Host "Are you a Microsoft 365 Global Admin, or do you have an Admin available to assist with the next step? (Y/N)"

function CaptureManualAddInVersion {
    param (
        [string]$FullLogFilePath
    )

    Write-Host ""
    Write-Host "🧭 No problem — let's capture the Add-in version manually." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please follow these steps:" -ForegroundColor Cyan
    Write-Host "  1) Open Outlook."
    Write-Host "  2) Start a new email message."
    Write-Host "  3) Click the 'Exclaimer' Add-in icon in the toolbar."
    Write-Host "  4) Look at the bottom of the Add-in pane — you’ll see the version number."
    Write-Host ""

    $addInVersion = Read-Host "Enter the version number displayed (e.g. 2.3.45)"

    if ([string]::IsNullOrWhiteSpace($addInVersion)) {
        Write-Host "⚠️ No version entered. Skipping manual version logging." -ForegroundColor Yellow
        Add-Content $FullLogFilePath '<p class="warning">No Add-in version provided by user.</p>'
        return
    }

    Write-Host "`n✅ Thank you — version recorded as: $addInVersion" -ForegroundColor Green

    # --- HTML Logging (safe formatting) ---
    Add-Content $FullLogFilePath '<h2>🧩 Exclaimer Add-in Information (Manual)</h2>'
    Add-Content $FullLogFilePath ('<p>User-provided Add-in version: <strong>{0}</strong></p>' -f [System.Web.HttpUtility]::HtmlEncode($addInVersion))
}


if ($adminChoice.ToUpper() -eq "N") {

    Write-Host "User chose not to run Exchange Online checks as Global Admin." -ForegroundColor Yellow

    # --- HTML Logging ---
    Add-Content $FullLogFilePath '<div class="section">'
    Add-Content $FullLogFilePath '<h2>🔐 Exchange Online Admin Checks</h2>'
    Add-Content $FullLogFilePath '<table>'
    Add-Content $FullLogFilePath '<tr><th>Property</th><th>Value</th></tr>'
    Add-Content $FullLogFilePath '<tr><td><strong>Admin checks performed</strong></td><td style="color:#F5A627;font-weight:bold;">No</td></tr>'
    Add-Content $FullLogFilePath '<tr><td><strong>Reason</strong></td><td>User chose not to run the checks as a Global Administrator.</td></tr>'
    Add-Content $FullLogFilePath '<tr><td><strong>Impact</strong></td><td>Some Exchange Online data could not be collected automatically.</td></tr>'
    Add-Content $FullLogFilePath '</table>'
    Add-Content $FullLogFilePath '</div>'

    CaptureManualAddInVersion -FullLogFilePath $FullLogFilePath
}
else {
    Write-Host "`n🔐 Checking for Exchange Online module..." -ForegroundColor Cyan

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

                    # --- Install the Exchange Online module ---
                    Write-Host "📦 Installing Exchange Online Management module..." -ForegroundColor Cyan
                    Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber

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
        try {
            Write-Host "`n🔗 Connecting to Exchange Online..." -ForegroundColor Cyan
            Write-Host "   You will be prompted to Sign in with Microsoft in order to continue." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
            Connect-ExchangeOnline -ErrorAction Stop
            Write-Host "✅ Connected successfully!" -ForegroundColor Green
            return $true
        } catch {
            Write-Host "❌ Connection failed: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    # --- Proceed only if module available ---
    if (CheckExchangeOnlineModule) {
            # --- HTML Logging (safe formatting) ---
            Add-Content $FullLogFilePath '<h2>🧩 Exclaimer Add-in Information (EXO Admin)</h2>'
        if (ConnectExchangeOnlineSession) {
            Write-Host "`n🎯 Querying Exclaimer Add-in deployment..." -ForegroundColor Cyan

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
                Add-Content $FullLogFilePath '<h3>Exclaimer Add-in Information (Admin)</h3>'
                Add-Content $FullLogFilePath '<table><tr><th>Type</th><th>Display Name</th><th>Version</th><th>Enabled</th><th>Scope</th><th>Deployment</th></tr>'

                if ($ProdResult) {
                    $enabledColor = if ($ProdResult.Enabled -ne $true) { ' style="color:red;font-weight:bold;"' } else { '' }

                    Add-Content $FullLogFilePath ('<tr><td>Production</td><td>{0}</td><td>{1}</td><td{5}>{2}</td><td>{3}</td><td>{4}</td></tr>' -f `
                            [System.Web.HttpUtility]::HtmlEncode($ProdResult.DisplayName),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.AppVersion),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.Enabled),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.Scope),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.Type),
                        $enabledColor)
                }

                if ($PreviewResult) {
                    $enabledColor = if ($PreviewResult.Enabled -ne $true) { ' style="color:red;font-weight:bold;"' } else { '' }

                    Add-Content $FullLogFilePath ('<tr><td>Preview</td><td>{0}</td><td>{1}</td><td{5}>{2}</td><td>{3}</td><td>{4}</td></tr>' -f `
                            [System.Web.HttpUtility]::HtmlEncode($PreviewResult.DisplayName),
                        [System.Web.HttpUtility]::HtmlEncode($PreviewResult.AppVersion),
                        [System.Web.HttpUtility]::HtmlEncode($PreviewResult.Enabled),
                        [System.Web.HttpUtility]::HtmlEncode($PreviewResult.Scope),
                        [System.Web.HttpUtility]::HtmlEncode($PreviewResult.Type),
                        $enabledColor)
                }

                Add-Content $FullLogFilePath '</table>'

                # Add attention note if either is not enabled
                $attentionMessages = @()

                if ($ProdResult -and $ProdResult.Enabled -ne $true) {
                    $identity = "$user\$ProdID"
                    $enableCommand = "Enable-App -Identity `"$identity`""
                    $attentionMessages += "<span><b>ℹ️ Production Add-in is Disabled:</b> Run the following command in PowerShell:</span><br><code>$enableCommand</code>"
                }

                if ($PreviewResult -and $PreviewResult.Enabled -ne $true) {
                    $identity = "$user\$PreviewID"
                    $enableCommand = "Enable-App -Identity `"$identity`""
                    $attentionMessages += "<span><b>ℹ️ Preview Add-in is Disabled:</b> Run the following command in PowerShell:</span><code>$enableCommand</code>"
                }

                if ($attentionMessages.Count -gt 0) {
                    $fullMessage = '<div class="info-after-error">' + ($attentionMessages -join "<br><br>") + '</div>'
                    Add-Content -Path $FullLogFilePath -Value $fullMessage

                    $sideNote = '<p class="side-note">If you have both Production and Preview versions deployed, only one requires being enabled.</p><p class="side-note">If you have re-opened PowerShell, then you may need to run the command below before enabling the Add-in.</p><code>Connect-ExchangeOnline</code><p class="side-note">When an Add-in is disabled for a user, it should not appear or function in Outlook. We have observed cases where it may still load in Outlook on the web, but this is not expected behaviour. If this occurs, it may need to be raised with Microsoft for further review.</p>'
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
                Add-Content $FullLogFilePath ('<p class="warning">No Exclaimer Add-ins found for user {0}.</p>' -f [System.Web.HttpUtility]::HtmlEncode($user))
            }

            # --- Organization-level Settings ---
            Write-Host "`nCollecting organization configuration related to Outlook Add-ins..." -ForegroundColor Cyan

            try {
                $orgConfig = Get-OrganizationConfig | Select-Object `
                    ReleaseTrack,
                    OAuth2ClientProfileEnabled,
                    OutlookMobileGCCRestrictionsEnabled,
                    AppsForOfficeEnabled,
                    EwsApplicationAccessPolicy

                Add-Content $FullLogFilePath '<div class="section">'
                Add-Content $FullLogFilePath '<h2>Organization Configuration - Add-in Compatibility</h2>'
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
                                '✅ Required for modern add-ins.'
                            }
                        }
                        'OutlookMobileGCCRestrictionsEnabled' {
                            $impact = if ($rawValue) {
                                '⚠️ Cloud add-ins not supported on Outlook Mobile.'
                            } else {
                                '✅ Mobile add-ins supported.'
                            }
                        }
                        'AppsForOfficeEnabled' {
                            $impact = if (-not $rawValue) {
                                '❌ Add-ins disabled organization-wide.'
                            } else {
                                '✅ Add-ins allowed.'
                            }
                        }
                        'EwsApplicationAccessPolicy' {
                            if ([string]::IsNullOrEmpty($rawValue) -or $rawValue -eq 'EnforceNone') {
                                $impact = '✅ No EWS restrictions detected.'
                            } elseif ($rawValue -eq 'EnforceAllowList') {
                                $impact = '⚠️ Only specific apps can use EWS. Verify Exclaimer is in the allow list.'
                            } elseif ($rawValue -eq 'EnforceBlockList') {
                                $impact = '⚠️ Some apps are blocked from EWS. Verify Exclaimer is not in the block list.'
                            } else {
                                $impact = "⚠️ Unrecognized policy value ($value). Review manually."
                            }
                        }
                        Default {
                            $impact = 'Review manually.'
                        }
                    }

                    Add-Content $FullLogFilePath ("<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>" -f $name, $value, [System.Web.HttpUtility]::HtmlEncode($impact))
                }

                Add-Content $FullLogFilePath '</table></div>'
            }
            catch {
                Write-Host "⚠️ Could not retrieve OrganizationConfig values." -ForegroundColor Yellow
                Add-Content $FullLogFilePath '<div class="section">'
                Add-Content $FullLogFilePath '<h2>Organization Configuration - Add-in Compatibility</h2>'
                Add-Content $FullLogFilePath '<p class="warning">Unable to retrieve organization configuration. Ensure proper Exchange Online connection and permissions.</p>'
                Add-Content $FullLogFilePath '</div>'
            }

            try {
                # Disconnect if needed
                #Disconnect-ExchangeOnline -Confirm:$false | Out-Null
                #Write-Host "`n🔒 Disconnected from Exchange Online." -ForegroundColor DarkGray
            } catch {}
        }
        else {
            Add-Content $FullLogFilePath '<p class="warning">Exchange Online connection failed or cancelled by user.</p>'
            CaptureManualAddInVersion -FullLogFilePath $FullLogFilePath
        }
    }
    else {
        Add-Content $FullLogFilePath '<p class="warning">Exchange Online module not available. Manual Add-in version collection required.</p>'
        CaptureManualAddInVersion -FullLogFilePath $FullLogFilePath
    }
    Write-Host "`n✅ Exclaimer Add-in details collection completed." -ForegroundColor Green
} # <-- closes main "else" for admin branch
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
GetFirewallLogs

function GetSupportSubmissionInstructions {

    Write-Host "`n========== Support Submission Instructions ==========" -ForegroundColor Cyan
    Write-Host "Please provide the following files to the support team:" -ForegroundColor White
    Write-Host ""
    Write-Host "1. Diagnostic report file:" -ForegroundColor Yellow
    Write-Host "   $FullLogFilePath" -ForegroundColor White
    Write-Host ""
    Write-Host "2. Windows Firewall log file:" -ForegroundColor Yellow
    Write-Host "   $Global:destination" -ForegroundColor White
    Write-Host ""
    Write-Host "Attach both files to your support ticket for review." -ForegroundColor Green

    # --- HTML Logging ---
    Add-Content $FullLogFilePath '<div class="section">'
    Add-Content $FullLogFilePath '<h2>📩 Support Submission Instructions</h2>'
    Add-Content $FullLogFilePath '<table>'
    Add-Content $FullLogFilePath '<tr><th>File</th><th>Details</th></tr>'

    Add-Content $FullLogFilePath (
        "<tr><td><strong>Diagnostic report</strong></td><td>{0}</td></tr>" -f $FullLogFilePath
    )
    Add-Content $FullLogFilePath (
        "<tr><td><strong>Firewall log</strong></td><td>{0}</td></tr>" -f $Global:destination
    )

    Add-Content $FullLogFilePath '</table>'
    Add-Content $FullLogFilePath '<p>Please attach listed files above when contacting the support team so they can review the collected data.</p>'
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
  <p>This report has been saved to:<br><code>$FullLogFilePath</code></p>
</div>
"@

@"
</div>
</body>
</html>
"@ | Add-Content -Path $FullLogFilePath -Encoding UTF8

# Open the file for user to view immediately (optional)
Start-Process $FullLogFilePath