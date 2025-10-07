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
#     1.0.0
#         - Initial version
#         - HTML-based diagnostics
#         - Checks for Outlook installation and version
#         - Checks for Exclaimer Agent and WebView2
#         - Cloud geolocation and endpoint connectivity
#         - Local signature inspection
#
# .INSTRUCTIONS
#     1. Open PowerShell (Administrator if possible)
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
    <title>Exclaimer Diagnostics Report</title>
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
<h1>Exclaimer Diagnostics Script Report</h1>
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



function Get-WindowsVersion {
    Write-Host "`n========== Microsoft Windows Version ==========" -ForegroundColor Cyan
    Add-Content $FullLogFilePath "<div class='section'>"
    Add-Content $FullLogFilePath "<h2>💻 Microsoft Windows Version</h2>"
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $version = $os.Version
    $caption = $os.Caption
    Write-Host "Windows Version: $caption ($version)"
    Add-Content $FullLogFilePath "<p>Windows Version: <strong>$caption ($version)</strong></p>"
}
Get-WindowsVersion

function Inspect-OutlookConfiguration {
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

    function Is-NewOutlookAppInstalled {
        return [bool](Get-NewOutlookPackage)
    }

    function Get-NewOutlookVersion {
        $package = Get-NewOutlookPackage
        if ($package) { return $package.Version }
        return $null
    }

    function Is-NewOutlookEnabled {
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

    function Is-ClassicOutlookInstalled {
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

    $classicInstalled     = Is-ClassicOutlookInstalled
    $newOutlookInstalled  = Is-NewOutlookAppInstalled
    $newOutlookEnabled    = Is-NewOutlookEnabled

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

Inspect-OutlookConfiguration



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

# Client-Side
# Try get "Connected Experience" on/off (not possible for user on/off, only if managed policy which is very uncommonly used)



# EXO
# Is Admin?
# Has Module? (No = Install module)
# Get Add-in version and state
# Get-OrganizationConfig | fl *OAuth*
# Get-OrganizationConfig | fl *EwsApp*
# Get-OrganizationConfig | fl *Outlook*


@"
</div>
</body>
</html>
"@ | Add-Content -Path $FullLogFilePath -Encoding UTF8

# Open the file for user to view immediately (optional)
Start-Process $FullLogFilePath