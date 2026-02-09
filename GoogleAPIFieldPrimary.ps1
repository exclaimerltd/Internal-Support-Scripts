# =========================================
# Google Workspace User Attribute Repair
# (Exclaimer fix for missing primary fields)
# =========================================
# SYNOPSIS:
# This script will set the 'primary' flag on a user's 'organizations' entry in Google Directory 
# so it is correctly synced with Exclaimer. It uses a service account with domain-wide delegation
# to perform updates via the Admin SDK API.
#
# INSTRUCTIONS:
# Please review the pre requisites in page below:
# https://github.com/exclaimerltd/Internal-Support-Scripts/blob/master/resources/GoogleAPIFieldPrimary.MD
# Push Enter to open it on your default browser, then follow the steps.
# Once completed, enter "C" to continue.
# =========================================

# Check if script is running in PowerShell 7+
if ($PSVersionTable.PSEdition -ne 'Core' -or $PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script must be run in PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Please run this script using pwsh (PowerShell 7+)." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to close"
    return  # Stop script execution without closing the session
}

# Default path: Downloads folder
$Global:FilePath = [System.IO.Path]::Combine([Environment]::GetFolderPath('UserProfile'), 'Downloads')
$LogFile = "GoogleUsersUpdateLog_$(Get-Date -Format 'HHmmss').csv"

# Check if the path exists, if not, use C:\Temp
if (-not (Test-Path -Path $Global:FilePath)) {
    $Global:FilePath = "C:\Temp"
    if (-not (Test-Path -Path $Global:FilePath)) {
        New-Item -Path $Global:FilePath -ItemType Directory -Force | Out-Null
    }
}

# Final full path for the log file
$FullLogFilePath = Join-Path $Global:FilePath $LogFile

# Prompt user to review pre-requisites page
Write-Host "`nPlease review the pre requisites in the page below:" -ForegroundColor Cyan
Write-Host "https://github.com/exclaimerltd/Internal-Support-Scripts/blob/master/resources/GoogleAPIFieldPrimary.MD"
Read-Host "Push Enter to open it on your default browser"
Start-Process "https://github.com/exclaimerltd/Internal-Support-Scripts/blob/master/resources/GoogleAPIFieldPrimary.MD"

# Wait for user confirmation
do {
    $confirm = Read-Host "Once you completed the steps in the article, please enter 'C' followed by Enter to continue"
} until ($confirm -eq 'C')

function EnsureModule {
    param([string]$Name)

    # Ensure the required module is installed
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing required module: $Name"
        Install-Module $Name -Scope CurrentUser -Force
    }
}

# (rest of your script continues unchanged...)


function Get-GoogleAccessToken {
    param(
        [Parameter(Mandatory)]
        [string]$JsonKeyPath,

        [Parameter(Mandatory)]
        [string]$AdminUser
    )

    try {
        # Load service account JSON
        $json = Get-Content $JsonKeyPath -Raw | ConvertFrom-Json
    }
    catch {
        throw "Unable to read the service account JSON file. Check the file path and file permissions."
    }

    try {
        $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

        $jwtHeader = @{ alg = "RS256"; typ = "JWT" } | ConvertTo-Json -Compress
        $jwtClaim = @{
            iss   = $json.client_email
            scope = "https://www.googleapis.com/auth/admin.directory.user"
            aud   = "https://oauth2.googleapis.com/token"
            exp   = $now + 3600
            iat   = $now
            sub   = $AdminUser
        } | ConvertTo-Json -Compress

        $headerEncoded = [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes($jwtHeader)
        ).TrimEnd("=")

        $claimEncoded = [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes($jwtClaim)
        ).TrimEnd("=")

        $rsa = [System.Security.Cryptography.RSA]::Create()
        $rsa.ImportFromPem($json.private_key)

        $signature = $rsa.SignData(
            [Text.Encoding]::UTF8.GetBytes("$headerEncoded.$claimEncoded"),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
        )

        $sigEncoded = [Convert]::ToBase64String($signature).TrimEnd("=")
        $jwt = "$headerEncoded.$claimEncoded.$sigEncoded"
    }
    catch {
        throw "Failed to generate the OAuth JWT. Ensure the script is running in PowerShell 7+ and the JSON key is valid."
    }

    try {
        $body = @{
            grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer"
            assertion  = $jwt
        }

        $token = Invoke-RestMethod -Method Post `
            -Uri "https://oauth2.googleapis.com/token" `
            -Body $body `
            -ErrorAction Stop

        return $token.access_token
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__ 2>$null

        switch ($statusCode) {
            400 { $friendly = "Token request rejected. Invalid Token or Service Account is not authorised for domain-wide delegation or the admin email is invalid." }
            401 { $friendly = "Authentication failed. The service account key may be invalid or revoked." }
            403 { $friendly = "Permission denied. Check domain-wide delegation, OAuth scopes, and ensure the admin user is a Super Admin." }
            default { $friendly = "Failed to obtain an access token from Google. Check service account configuration and permissions." }
        }

        Write-Host "ERROR: $friendly (HTTP $statusCode)" -ForegroundColor Red
        exit 1
    }

}

function Update-GoogleUser {
    param(
        [string]$UserEmail,
        [string]$AccessToken,
        [hashtable]$Payload,
        [int]$Count
    )

    $headers = @{
        Authorization = "Bearer $AccessToken"
    }

    $status = ""
    try {
        Invoke-RestMethod `
            -Method Patch `
            -Uri "https://admin.googleapis.com/admin/directory/v1/users/$UserEmail" `
            -Headers $headers `
            -Body ($Payload | ConvertTo-Json -Depth 5) `
            -ContentType "application/json"

        $status = "Success"
        return $true
    }
    catch {
        $status = "Failed: $($_.Exception.Message)"
        return $false
    }
    finally {
        $logEntry = [PSCustomObject]@{
            Count     = $Count
            Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            UserEmail = $UserEmail
            Status    = $status
        }

        if (-not (Test-Path -Path $FullLogFilePath)) {
            $logEntry | Export-Csv -Path $FullLogFilePath -NoTypeInformation
        }
        else {
            $logEntry | Export-Csv -Path $FullLogFilePath -NoTypeInformation -Append
        }
    }
}

# ================== SCRIPT START ==================

EnsureModule -Name "PowerShellGet"
Clear-Host
Write-Host "           -----------------------------------------------" -ForegroundColor Cyan
Write-Host "           |                   EXCLAIMER                 |" -ForegroundColor Yellow
Write-Host "           |     Google Workspace User Attribute Fix     |" -ForegroundColor Yellow
Write-Host "           -----------------------------------------------" -ForegroundColor Cyan
Write-Host ""
Start-Sleep -Seconds 1

# Prompt for JSON key file path
do {
    $jsonPath = Read-Host "Enter full path to JSON key file (e.g. C:\Temp\my_token.json)"

    # Remove surrounding quotes if present
    $jsonPath = $jsonPath.Trim('"')

    if (-not (Test-Path -Path $jsonPath -PathType Leaf)) {
        Write-Host "File not found. Please enter a valid file path." -ForegroundColor Yellow
        $jsonPath = $null
    }
} until ($jsonPath)


# Prompt for Google Admin email address
$emailRegex = '^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$'

do {
    $adminUser = Read-Host "Enter Google Admin email address for delegation"

    # Force lowercase
    $adminUser = $adminUser.ToLower()

    if ($adminUser -notmatch $emailRegex) {
        Write-Host "Invalid email address format. Please try again." -ForegroundColor Yellow
        $adminUser = $null
    }
} until ($adminUser)

$accessToken = Get-GoogleAccessToken -JsonKeyPath $jsonPath -AdminUser $adminUser

$field = "organizations"
Write-Host "`nField to update: $field"

# Hardcoded JSON to set primary=true on first organization
$orgUpdate = ConvertFrom-Json '[{"primary": true}]'

$scope = Read-Host "Update (1) single user or (2) all users?"
Write-Host ""

if ($scope -eq "1") {
    $users = @((Read-Host "Enter user email address"))
}
else {
    $headers = @{ Authorization = "Bearer $accessToken" }

    $users = @()
    $pageToken = $null

    do {
        $uri = "https://admin.googleapis.com/admin/directory/v1/users?customer=my_customer&maxResults=500&query=isSuspended=false"

        if ($pageToken) {
            $uri += "&pageToken=$pageToken"
        }

        $response = Invoke-RestMethod -Uri $uri -Headers $headers

        if ($response.users) {
            $users += $response.users | ForEach-Object { $_.primaryEmail }
        }

        $pageToken = $response.nextPageToken
    }
    while ($pageToken)
}

# Count goes here
$totalUsers = $users.Count
$counter = 0
Write-Host "`nTotal users to be processed: $totalUsers`n"
foreach ($user in $users) {
    $counter++

    Write-Host "`rProcessing ($counter / $totalUsers): $user".PadRight(120) `
        -ForegroundColor Yellow -NoNewline

    # Fetch current organizations array
    $current = Invoke-RestMethod `
        -Uri "https://admin.googleapis.com/admin/directory/v1/users/$user" `
        -Headers @{ Authorization = "Bearer $accessToken" }

    $orgs = @()
    if ($current.organizations) { $orgs = $current.organizations }
    if ($orgs.Count -eq 0) {
        Write-Host "`rSkipping ($counter / $totalUsers): no org - $user".PadRight(120) `
            -ForegroundColor DarkYellow -NoNewline
        continue
    }

    # CHECK: already primary?
    if ($orgs[0].primary -eq $true) {
        Write-Host "`rSkipping ($counter / $totalUsers): already primary - $user".PadRight(120) `
            -ForegroundColor DarkYellow -NoNewline

        # Optional: log skip
        $logEntry = [PSCustomObject]@{
            Count     = $counter
            Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            UserEmail = $user
            Status    = "Skipped (already primary)"
        }

        if (-not (Test-Path $FullLogFilePath)) {
            $logEntry | Export-Csv $FullLogFilePath -NoTypeInformation
        }
        else {
            $logEntry | Export-Csv $FullLogFilePath -NoTypeInformation -Append
        }

        continue
    }

    # Only update the first organization
    foreach ($prop in $orgUpdate[0].PSObject.Properties.Name) {
        if (-not $orgs[0].PSObject.Properties[$prop]) {
            $orgs[0] | Add-Member -MemberType NoteProperty -Name $prop -Value $orgUpdate[0].$prop -Force
        }
        else {
            $orgs[0].$prop = $orgUpdate[0].$prop
        }
    }

    $payload = @{ organizations = $orgs }

    Update-GoogleUser -UserEmail $user -AccessToken $accessToken -Payload $payload -Count $counter | Out-Null
}


Write-Host "`rProcessing completed".PadRight(120) -ForegroundColor Green -NoNewline
# Final message after all users are processed
Write-Host "`n`nUpdate completed. Log of processed users has been saved to:" -ForegroundColor Green
Write-Host $FullLogFilePath -ForegroundColor Cyan
Write-Host "`nAllow Google sync time before validating changes."

