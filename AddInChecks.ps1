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
$Path = [System.IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), 'Downloads')
$LogFile = "AddInChecks.html"

# Check if the path exists, if not, use C:\Temp
if (-not (Test-Path -Path $Path)) {
    $Path = "C:\Temp"
    if (-not (Test-Path -Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

# Final full path for the log file
$FullLogFilePath = Join-Path $Path $LogFile
$DateTimeRun = Get-Date -Format "ddd dd MMMM yyyy, HH:mm 'UTC' K"

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
        .section { margin-bottom: 30px; }
        .success { color: green; font-weight: bold; }
        .fail { color: red; font-weight: bold; }
        .warning { color: orange; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
        th { background-color: #eee; }
        a { color: #0078D4; text-decoration: none; } a:hover { text-decoration: underline; }
        .info-after-error { color: #0c5460; background-color: #d1ecf1; border: 1px solid #bee5eb; padding: 10px; border-radius: 4px; margin-top: 10px; }
    </style>
</head>
<body>
<div class="container">
<h1>Exclaimer Diagnostics Report - Client-Side</h1>
<p><strong>Run Date:</strong> $DateTimeRun</p>
"@ | Set-Content -Path $FullLogFilePath -Encoding UTF8


Write-Host "           -----------------------------------------------" -ForegroundColor Cyan
Write-Host "           |                 EXCLAIMER                   |" -ForegroundColor Yellow
Write-Host "           |       Diagnostics Script Collection         |" -ForegroundColor Yellow
Write-Host "           -----------------------------------------------" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1

#It should set the below:
$Path = [System.IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), 'Downloads')
$LogFile = "AddInChecks.html" 
$DateTimeRun = Get-Date -Format "ddd dd MMMM yyyy, HH:MM 'UTC' K"
$ProdID = "efc30400-2ac5-48b7-8c9b-c0fd5f266be2"
$PreviewID = "a8d42ca1-6f1f-43b5-84e1-9ff40e967ccc"

function Get-ExclaimerUserInput {
    [CmdletBinding()]
    param ()

    while ($true) {
        # Initialize object
        $userInput = [PSCustomObject]@{
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
                $userInput.Email = $email.Trim()
                break
            } else {
                Write-Host "Invalid email format. Try again." -ForegroundColor Red
            }
        }

        # --- 2) Purpose ---
        do {
            Clear-Host
            Write-Host "`nWhat would you like to do?" -ForegroundColor Cyan
            Write-Host "  1) Troubleshoot an issue"
            Write-Host "  2) Review configuration overview"
            $choice = Read-Host "`nEnter choice (1 or 2)"
        } while ($choice -notmatch '^[12]$')

        $userInput.Purpose = if ($choice -eq '1') { 'Troubleshooting' } else { 'Configuration Overview' }

        # --- 3) If troubleshooting, ask follow-ups ---
        if ($userInput.Purpose -eq 'Troubleshooting') {
            # Users affected
            do {
                Clear-Host
                Write-Host "`nHow many users are affected?" -ForegroundColor Cyan
                Write-Host "  1) All users"
                Write-Host "  2) Specify number"
                $uc = Read-Host "`nEnter choice (1 or 2)"
            } while ($uc -notmatch '^[12]$')

            if ($uc -eq '1') {
                $userInput.UsersAffected = 'All Users'
            } else {
                do {
                    $num = Read-Host "Enter the approximate number of affected users (digits only)"
                } while ($num -notmatch '^\d+$')
                $userInput.UsersAffected = [int]$num
            }

            # Outlook versions
            do {
                Clear-Host
                Write-Host "`nWhich Outlook version(s) are affected?" -ForegroundColor Cyan
                Write-Host "  1) Outlook Desktop"
                Write-Host "  2) Outlook Web"
                Write-Host "  3) Outlook Mobile"
                Write-Host "  4) Multiple / All"
                $oChoice = Read-Host "`nEnter choice (1-4)"
            } while ($oChoice -notmatch '^[1-4]$')

            switch ($oChoice) {
                1 { $userInput.OutlookAffected = 'Outlook Desktop' }
                2 { $userInput.OutlookAffected = 'Outlook Web' }
                3 { $userInput.OutlookAffected = 'Outlook Mobile' }
                4 { $userInput.OutlookAffected = 'Multiple / All' }
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
                1 { $userInput.Network = 'Internal Only' }
                2 { $userInput.Network = 'External Only' }
                3 { $userInput.Network = 'Both Networks' }
            }
        }

        # --- 4) Show console summary ---
        Clear-Host
        Write-Host ""
        Write-Host "========================================" -ForegroundColor DarkGray
        Write-Host "            Summary captured" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor DarkGray

        Write-Host ("Purpose:          {0}" -f $userInput.Purpose) -ForegroundColor Cyan
        Write-Host ("Email:            {0}" -f $userInput.Email) -ForegroundColor Yellow

        if ($userInput.Purpose -eq 'Troubleshooting') {
            Write-Host ("Users Affected:   {0}" -f $userInput.UsersAffected) -ForegroundColor White
            Write-Host ("Outlook Affected: {0}" -f $userInput.OutlookAffected) -ForegroundColor White
            Write-Host ("Network Scope:    {0}" -f $userInput.Network) -ForegroundColor White
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
            Add-Content $FullLogFilePath "<tr><td><strong>Purpose:</strong></td><td>$($userInput.Purpose)</td></tr>"
            Add-Content $FullLogFilePath "<tr><td><strong>Email:</strong></td><td>$($userInput.Email)</td></tr>"

            if ($userInput.Purpose -eq 'Troubleshooting') {
                Add-Content $FullLogFilePath "<tr><td><strong>Users Affected:</strong></td><td>$($userInput.UsersAffected)</td></tr>"
                Add-Content $FullLogFilePath "<tr><td><strong>Outlook Affected:</strong></td><td>$($userInput.OutlookAffected)</td></tr>"
                Add-Content $FullLogFilePath "<tr><td><strong>Network Scope:</strong></td><td>$($userInput.Network)</td></tr>"
            }

            Add-Content $FullLogFilePath "</table></div>"
            return $userInput
        } else {
            Write-Host "`nLet's try again..." -ForegroundColor Yellow
            Start-Sleep -Seconds 1
            Clear-Host
        }
    }
}

function Get-Region {
    param (
        [PSCustomObject]$userInput
    )

    # Define log file path (adjust as needed)

    $email = $userInput.Email.ToLower().Trim()

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
    $supportNote   = 'Consider upgrading to Windows 10 22H2 or Windows 11 for compatibility.'

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
            "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
        )

        foreach ($path in $registryPaths) {
            if (Test-Path $path) {
                try {
                    $props = Get-ItemProperty -Path $path
                    if ($props.PSObject.Properties.Name -contains "IsUsingNewOutlook" -and $props.IsUsingNewOutlook -eq 1) {
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

        if ($newOutlookEnabled) {
            Write-Host "New Outlook experience is ENABLED (user preference)." -ForegroundColor Green
            Add-Content $FullLogFilePath "<li><span class='success'>New Outlook is enabled</span></li>"
        } else {
            Write-Host "New Outlook is installed, but NOT set as default." -ForegroundColor Yellow
            Add-Content $FullLogFilePath "<li><span class='warning'>New Outlook is installed but not set as 'Default'</span></li>"
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
    <tr><th>Office Version</th><th>Build</th><th>License Type</th><th>Compatibility</th></tr>
    <tr>
        <td>$officeVersion</td>
        <td>$officeBuild</td>
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
                } else {
                    'MSI'
                }
            }}

        if ($apps) {
            $foundApps += $apps
        }
    } catch {
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
} else {
    Write-Host "The Exclaimer Cloud Signature Update Agent is not installed." -ForegroundColor Yellow
    Add-Content $FullLogFilePath "<p>✅ The Exclaimer Cloud Signature Update Agent is not installed.</p>"
}

Add-Content $FullLogFilePath "</div>"

# Define registry paths for 64-bit and 32-bit uninstall keys
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
)

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
    } catch {
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
} else {
    Write-Host "Microsoft Edge WebView2 Runtime is not installed." -ForegroundColor Yellow
    Add-Content $FullLogFilePath "<p class='warning'>Microsoft Edge WebView2 Runtime is not installed.</p>"
}


# Client-Side
# Try get "Connected Experience" on/off (not possible for user on/off, only if managed policy which is very uncommonly used)

# -------------------------------------------------------------------
# 📨 EXCLAIMER ADD-IN DETAILS COLLECTION (User or Admin)
# -------------------------------------------------------------------

Write-Host ""
Write-Host "=== Exclaimer Add-in Information ===" -ForegroundColor Cyan
Write-Host ""

# --- Step: Check if user is Global Admin ---
$adminChoice = Read-Host "Are you a Microsoft 365 Global Admin, or do you have an Admin available to assist with the next part? (Y/N)"

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
    Add-Content $FullLogFilePath '<h2>Exclaimer Add-in Information (Manual)</h2>'
    Add-Content $FullLogFilePath ('<p>User-provided Add-in version: <strong>{0}</strong></p>' -f [System.Web.HttpUtility]::HtmlEncode($addInVersion))
}


if ($adminChoice.ToUpper() -eq "N") {
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
            Write-Host "   You will be promted to Sign in with Microsoft in order to continue." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
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
            Add-Content $FullLogFilePath '<h2>Exclaimer Add-in Information (EXO)</h2>'
        if (ConnectExchangeOnlineSession) {
            Write-Host "`n🎯 Querying Exclaimer Add-in deployment..." -ForegroundColor Cyan

            $user = $userInput.Email
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
                    Add-Content $FullLogFilePath ('<tr><td>Production</td><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f `
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.DisplayName),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.AppVersion),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.Enabled),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.Scope),
                        [System.Web.HttpUtility]::HtmlEncode($ProdResult.Type))
                }
                if ($PreviewResult) {
                    Add-Content $FullLogFilePath ('<tr><td>Preview</td><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>' -f `
                        [System.Web.HttpUtility]::HtmlEncode($PreviewResult.DisplayName),
                        [System.Web.HttpUtility]::HtmlEncode($PreviewResult.AppVersion),
                        [System.Web.HttpUtility]::HtmlEncode($PreviewResult.Enabled),
                        [System.Web.HttpUtility]::HtmlEncode($PreviewResult.Scope),
                        [System.Web.HttpUtility]::HtmlEncode($PreviewResult.Type))
                }

                Add-Content $FullLogFilePath '</table>'

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
} # <-- closes main "else" for admin branch


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