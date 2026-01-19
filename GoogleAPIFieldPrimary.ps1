# =========================================
# Google Workspace User Attribute Repair
# (Exclaimer fix for missing primary fields)
# =========================================
# INSTRUCTIONS:
# Before running this script, complete the following steps:
# 1. Go to https://console.cloud.google.com/ and create a new project.
# 2. Navigate to IAM & Admin > Service Accounts and create a service account.
#    - Note the service account name.
# 3. Create an OAuth 2.0 Client ID for the service account.
#    - Select "Service Account" as the application type.
#    - Choose the service account created in step 2.
# 4. Download the JSON key for the OAuth client and store it securely
#    (for example: C:\Temp\service-account.json).
# 5. Enable the Admin SDK API for the project:
#    https://console.developers.google.com/apis/api/admin.googleapis.com/overview
# 6. In the Google Admin Console (https://admin.google.com):
#    Security > Access and Data Control > API Controls > Manage Domain-wide Delegation
#    - Authorize the service account client ID.
#    - Add the scope:
#      https://www.googleapis.com/auth/admin.directory.user
# 7. Ensure you have the Google Admin email address used for
#    domain-wide delegation (for example: admin@yourdomain.com).
# =========================================

function EnsureModule {
    param([string]$Name)

    # Display instructions before running the script
    $instructions = @(
        "Google Workspace User Attribute Repair (Exclaimer fix)",
        "",
        "This script will set the 'primary' flag on a user's 'organizations' entry in Google Directory so it is synced with Exclaimer.",
        "",
        "Before running this script, complete the following steps:",
        "1. Go to https://console.cloud.google.com/ and create a new project.",
        "2. Navigate to IAM & Admin > Service Accounts, create a service account.",
        "   - Note the service account name.",
        "3. Create an OAuth 2 Client ID for the service account:",
        "   - Choose 'Service Account' as the application type.",
        "   - Select the service account you created in Step 2 from the dropdown.",
        "4. Download the JSON key for the OAuth client and save it to a secure location (e.g., C:\\Temp\\service-account.json).",
        "5. Enable the Admin SDK API for the project here:",
        "   https://console.developers.google.com/apis/api/admin.googleapis.com/overview",
        "6. In Google Admin Console (https://admin.google.com):",
        "   Security > Access and Data Control > API Controls > Manage Domain-wide Delegation",
        "   - Authorize the service account client ID",
        "   - Scope: https://www.googleapis.com/auth/admin.directory.user",
        "7. Ensure you know the Google admin email for domain-wide delegation (e.g., admin@yourdomain.com)"
    )

    foreach ($line in $instructions) {
        Write-Host $line
    }

    Read-Host "`nPress Enter to continue after completing all the steps above..."

    # Ensure the required module is installed
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Installing required module: $Name"
        Install-Module $Name -Scope CurrentUser -Force
    }
}

function Get-GoogleAccessToken {
    param(
        [string]$JsonKeyPath,
        [string]$AdminUser
    )

    $json = Get-Content $JsonKeyPath -Raw | ConvertFrom-Json
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

    $headerEncoded = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jwtHeader)).TrimEnd("=")
    $claimEncoded  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($jwtClaim)).TrimEnd("=")

    $rsa = [System.Security.Cryptography.RSA]::Create()
    $rsa.ImportFromPem($json.private_key)

    $signature = $rsa.SignData(
        [Text.Encoding]::UTF8.GetBytes("$headerEncoded.$claimEncoded"),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )

    $sigEncoded = [Convert]::ToBase64String($signature).TrimEnd("=")
    $jwt = "$headerEncoded.$claimEncoded.$sigEncoded"

    $body = @{
        grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer"
        assertion  = $jwt
    }

    $token = Invoke-RestMethod -Method Post `
        -Uri "https://oauth2.googleapis.com/token" `
        -Body $body

    return $token.access_token
}

function Update-GoogleUser {
    param(
        [string]$UserEmail,
        [string]$AccessToken,
        [hashtable]$Payload
    )

    $headers = @{
        Authorization = "Bearer $AccessToken"
    }

    try {
        Invoke-RestMethod `
            -Method Patch `
            -Uri "https://admin.googleapis.com/admin/directory/v1/users/$UserEmail" `
            -Headers $headers `
            -Body ($Payload | ConvertTo-Json -Depth 5) `
            -ContentType "application/json"

        Write-Host "Successfully updated: $UserEmail" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to update $UserEmail - $($_.Exception.Message)"
    }
}

# ================== SCRIPT START ==================

EnsureModule -Name "PowerShellGet"

Write-Host "Google Workspace user attribute repair (Exclaimer fix)`n"

$jsonPath = Read-Host "Enter full path to JSON key file (i.e: C:\Temp\my_token.json)"
$adminUser = Read-Host "Enter Google admin email for domain-wide delegation"

$accessToken = Get-GoogleAccessToken -JsonKeyPath $jsonPath -AdminUser $adminUser

$field = "organizations"
Write-Host "`nField to update: $field"

# Hardcoded JSON to set primary=true on first organization
$orgUpdate = ConvertFrom-Json '[{"primary": true}]'

$scope = Read-Host "Update (1) single user or (2) all users?"

if ($scope -eq "1") {
    $users = @((Read-Host "Enter user email address"))
}
else {
    $headers = @{ Authorization = "Bearer $accessToken" }
    $allUsers = Invoke-RestMethod `
        -Uri "https://admin.googleapis.com/admin/directory/v1/users?customer=my_customer&maxResults=500" `
        -Headers $headers

    $users = $allUsers.users | ForEach-Object { $_.primaryEmail }
}

foreach ($user in $users) {
    Write-Host "`nProcessing: $user" -ForegroundColor Yellow

    # Fetch current organizations array
    $current = Invoke-RestMethod `
        -Uri "https://admin.googleapis.com/admin/directory/v1/users/$user" `
        -Headers @{ Authorization = "Bearer $accessToken" }

    $orgs = @()
    if ($current.organizations) { $orgs = $current.organizations }
    if ($orgs.Count -eq 0) {
        Write-Warning "$user has no existing organization. Skipping."
        continue
    }

    # Only update the first organization
    foreach ($prop in $orgUpdate[0].PSObject.Properties.Name) {
        if (-not $orgs[0].PSObject.Properties[$prop]) {
            # Add the property if it doesn't exist
            $orgs[0] | Add-Member -MemberType NoteProperty -Name $prop -Value $orgUpdate[0].$prop -Force
        }
        else {
            # Otherwise, update existing
            $orgs[0].$prop = $orgUpdate[0].$prop
        }
    }
    
    $payload = @{ organizations = $orgs }

    Update-GoogleUser -UserEmail $user -AccessToken $accessToken -Payload $payload  | Out-Null
}

Write-Host "`nUpdate completed. Allow Google sync time before re-running Exclaimer sync."
