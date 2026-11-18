# 
<#
.SYNOPSIS
    Creates a mail-enabled security group in Exchange Online.
.DESCRIPTION
    Signs in using the ExchangeOnlineManagement module and creates a mail-enabled security group,
    for example to scope the Exclaimer 'Identify messages to send to Exclaimer Cloud' transport rule
    to a pilot group of users.

    The script checks the module, the connection and the account's Exchange roles, collects the
    group details, asks for confirmation, creates the group, optionally adds members, and then
    verifies the result.
.NOTES
    Date: 18th November 2026
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    - Account holding the Distribution Groups and View-Only Recipients roles
      (Organization Management role group or equivalent)
    - SMTP domain for the group already added to Microsoft 365 and verified
.HISTORY
    1.0 - Initial version
#>
$Global:scriptVersion = "'1.0.0'"

function confirmExclaimerSupportApproval {
    Write-Host ""
    Write-Host "IMPORTANT NOTICE" -ForegroundColor Yellow
    Write-Host "This script must only be run when advised by the Exclaimer Technical Support Team." -ForegroundColor Yellow
    Write-Host "It creates a new mail-enabled security group in your Microsoft 365 tenant." -ForegroundColor Yellow
    Write-Host "No existing groups, rules or connectors are changed or removed." -ForegroundColor Yellow
    Write-Host ""

    $confirm = Read-Host "Do you understand and accept these conditions? (Y/N)"

    if ($confirm.ToUpper() -ne "Y") {
        Write-Host "Operation cancelled. No changes have been made." -ForegroundColor Red
        exit
    }
}

function checkExchangeOnlineModule {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Write-Host "[OK] Exchange Online Management module is already installed." -ForegroundColor Green
    } else {
        Write-Host "[..]  Exchange Online Management module not found." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "[i]  The installation requires the NuGet provider and PowerShell Gallery access." -ForegroundColor Cyan
        Write-Host "    You may see prompts asking to install NuGet or trust the PowerShell Gallery; please answer 'Y' when prompted." -ForegroundColor Cyan
        Write-Host ""

        $installChoice = Read-Host "Would you like to install it now? (Y/N)"
        if ($installChoice.ToUpper() -eq "Y") {
            try {
                Write-Host "`n[..] Preparing to install prerequisites..." -ForegroundColor Cyan
                # --- Ensure NuGet provider is installed ---
                if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                    Write-Host "[..] Installing NuGet provider..." -ForegroundColor Cyan
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false | Out-Null
                }
                # --- Ensure PowerShell Gallery is trusted ---
                $galleryTrusted = (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue).InstallationPolicy
                if ($galleryTrusted -ne 'Trusted') {
                    Write-Host "[..] Trusting PowerShell Gallery repository..." -ForegroundColor Cyan
                    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
                }
                # --- Install the Exchange Online module ---
                Write-Host "[..] Installing Exchange Online Management module..." -ForegroundColor Cyan
                Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber
                Write-Host "[OK] Installation completed successfully!" -ForegroundColor Green
            } catch {
                Write-Host "[X] Failed to install the module: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "   Script cannot continue without the Exchange Online Management module." -ForegroundColor Yellow
                exit
            }
        } else {
            Write-Host "[X] The Exchange Online Management module is required. Script cannot continue." -ForegroundColor Red
            exit
        }
    }
}

function modernAuthConnect {
    Import-Module ExchangeOnlineManagement
    $connections = Get-ConnectionInformation
    if ($connections.Count -gt 1) {
        Write-Host "[!]  Multiple Exchange sessions detected. Disconnecting..." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false
        $connections = $null
    }
    if (-not $connections) {
        Write-Host "   You will be prompted to Sign in with Microsoft in order to continue." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        Connect-ExchangeOnline -ShowBanner:$false
    } else {
        $connectedAs = [string]($connections | Select-Object -First 1 -ExpandProperty UserPrincipalName)
        Write-Host ""
        Write-Host "[OK] Already connected as: $connectedAs" -ForegroundColor Green
        Write-Host ""
        Write-Host "   Y: Continue with this session." -ForegroundColor Cyan
        Write-Host "   N: Disconnect and sign in as a different account." -ForegroundColor Cyan
        Write-Host ""
        $confirm = Read-Host "Is this the correct account? (Y/N)"
        if ($confirm.ToUpper() -ne "Y") {
            Write-Host "   Disconnecting current session..." -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Host "   You will be prompted to Sign in with Microsoft in order to continue." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
            Connect-ExchangeOnline -ShowBanner:$false
        } else {
            Write-Host "[OK] Continuing with existing session." -ForegroundColor Green
        }
    }

    if (-not (Get-ConnectionInformation)) {
        Write-Host "[X] Unable to establish an Exchange Online session. No changes have been made." -ForegroundColor Red
        exit
    }
}

function checkExchangePermissions {
    Write-Host "`nChecking Exchange permissions..." -ForegroundColor Cyan
    $requiredRoles = @("Distribution Groups", "View-Only Recipients")
    $me = [string](Get-ConnectionInformation | Select-Object -First 1 -ExpandProperty UserPrincipalName)
    $myGuid = [string](Get-User -Identity $me -ErrorAction SilentlyContinue).ExternalDirectoryObjectId
    Write-Host "   Connected as: $me" -ForegroundColor Cyan
    if (-not (Get-Command Get-ManagementRoleAssignment -ErrorAction SilentlyContinue)) {
        Write-Host "[!] Exchange cmdlets unavailable. Attempting to reimport module..." -ForegroundColor Yellow
        Import-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
        if (-not (Get-Command Get-ManagementRoleAssignment -ErrorAction SilentlyContinue)) {
            Write-Host "[X] Exchange permissions check failed." -ForegroundColor Red
            $requiredRoles | ForEach-Object { Write-Host "   Missing role: $_" -ForegroundColor Red }
            Write-Host "   Please ensure the account holds the Organization Management role group or equivalent." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
            exit
        }
    }
    $allRoles = (Get-ManagementRoleAssignment -GetEffectiveUsers -Delegating $false -ErrorAction SilentlyContinue 2>$null | Where-Object { $_.EffectiveUserName -eq $myGuid -or $_.EffectiveUserName -eq $me }).Role | Select-Object -Unique
    $missing = $requiredRoles | Where-Object { $_ -notin $allRoles }
    if ($missing) {
        Write-Host "[X] Exchange permissions check failed." -ForegroundColor Red
        $missing | ForEach-Object { Write-Host "   Missing role: $_" -ForegroundColor Red }
        Write-Host "   Please ensure the account holds the Organization Management role group or equivalent." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        exit
    }
    Write-Host "[OK] Exchange permissions check passed." -ForegroundColor Green
    Start-Sleep -Seconds 2
}

# User inputs
function getGroupDetails {
    Clear-Host
    $ps1Version = $Global:scriptVersion
    $me = [string](Get-ConnectionInformation | Select-Object -First 1 -ExpandProperty UserPrincipalName)

    # ----------------------------
    # Group name
    # ----------------------------
    Write-Host "`nMail-Enabled Security Group Details" -ForegroundColor Cyan
    $groupName = (Read-Host "Enter a display name for the group (for example: Exclaimer Pilot Users)").Trim()
    if ([string]::IsNullOrWhiteSpace($groupName)) {
        Write-Host "No group name provided." -ForegroundColor Red
        return $null
    }

    $nameEscaped = $groupName.Replace("'", "''")
    if (Get-Recipient -Filter "DisplayName -eq '$nameEscaped'" -ErrorAction SilentlyContinue) {
        Write-Host "A recipient with the display name '$groupName' already exists. No changes have been made." -ForegroundColor Red
        return $null
    }

    # ----------------------------
    # Accepted domain (select by number)
    # ----------------------------
    Write-Host "`nRetrieving accepted domains..." -ForegroundColor Yellow
    $acceptedDomains = @(Get-AcceptedDomain |
        Where-Object { $_.DomainName -notlike "*.mail.onmicrosoft.com" } |
        Sort-Object -Property @{Expression='Default';Descending=$true}, @{Expression='DomainName';Descending=$false})
    if ($acceptedDomains.Count -eq 0) {
        Write-Host "No accepted domains found in this tenant. Script cannot continue." -ForegroundColor Red
        return $null
    }

    if ($acceptedDomains.Count -eq 1) {
        $smtpDomain = [string]$acceptedDomains[0].DomainName
        Write-Host "Only one accepted domain found. Using: $smtpDomain" -ForegroundColor Green
    } else {
        Write-Host "`nSelect the domain for the group's email address:`n" -ForegroundColor Yellow
        for ($i = 0; $i -lt $acceptedDomains.Count; $i++) {
            $defaultTag = if ($acceptedDomains[$i].Default) { "   (default)" } else { "" }
            Write-Host "[$($i+1)] $($acceptedDomains[$i].DomainName)$defaultTag"
        }
        $selection = Read-Host "`nEnter the number for the domain you want to use"
        if ($selection -notmatch '^[0-9]+$' -or [int]$selection -lt 1 -or [int]$selection -gt $acceptedDomains.Count) {
            Write-Host "Invalid selection. No changes have been made." -ForegroundColor Red
            return $null
        }
        $smtpDomain = [string]$acceptedDomains[[int]$selection - 1].DomainName
    }
    Write-Host "Selected domain: $smtpDomain" -ForegroundColor Green

    # ----------------------------
    # Username (group name without spaces, unless already in use)
    # ----------------------------
    Write-Host "`nChecking the email address is not already in use..." -ForegroundColor Yellow
    $alias = ($groupName -replace '\s','').ToLower() -replace '[^a-z0-9._-]',''
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($alias)) {
            Write-Host "The group name doesn't produce a usable username." -ForegroundColor Yellow
        } else {
            $smtp = "$alias@$smtpDomain"
            $existing = Get-Recipient -Filter "EmailAddresses -eq 'smtp:$smtp' -or Alias -eq '$alias'" -ErrorAction SilentlyContinue
            if (-not $existing) { break }
            Write-Host "A recipient already exists that matches '$smtp':" -ForegroundColor Yellow
            $existing | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails | Format-Table -AutoSize | Out-String | Write-Host
        }
        $newAlias = (Read-Host "Enter a different username (the part before the @), or press Enter to cancel").Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($newAlias)) {
            Write-Host "Cancelled. No changes have been made." -ForegroundColor Yellow
            return $null
        }
        $alias = $newAlias -replace '[^a-z0-9._-]',''
    }
    Write-Host "Email address is available: $smtp" -ForegroundColor Green

    # ----------------------------
    # Owner
    # ----------------------------
    Write-Host "`nGroup Owner" -ForegroundColor Cyan
    $ownerInput = (Read-Host "Enter the owner's email address (press Enter to use $me)").Trim()
    if ([string]::IsNullOrWhiteSpace($ownerInput)) { $ownerInput = $me }
    $ownerObj = Get-Recipient -Identity $ownerInput -ErrorAction SilentlyContinue
    if (-not $ownerObj) {
        Write-Host "Owner '$ownerInput' was not found in this tenant." -ForegroundColor Red
        return $null
    }
    $owner = [string]$ownerObj.PrimarySmtpAddress

    # ----------------------------
    # Members (optional)
    # ----------------------------
    Write-Host "`nGroup Members" -ForegroundColor Cyan
    Write-Host "   Enter email addresses separated by commas, or press Enter to create the group empty." -ForegroundColor Cyan
    $memberInput = (Read-Host "Members").Trim()
    $members = @()
    $invalidMembers = @()
    if (-not [string]::IsNullOrWhiteSpace($memberInput)) {
        foreach ($m in ($memberInput -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
            $r = Get-Recipient -Identity $m -ErrorAction SilentlyContinue
            if ($r) { $members += [string]$r.PrimarySmtpAddress } else { $invalidMembers += $m }
        }
        if ($invalidMembers.Count -gt 0) {
            Write-Host "These members were not found and will be skipped: $($invalidMembers -join ', ')" -ForegroundColor Yellow
        }
    }

    # ----------------------------
    # External senders
    # ----------------------------
    Write-Host "`nSender Authentication" -ForegroundColor Cyan
    Write-Host "   Y: Only authenticated (internal) senders can email the group. Recommended." -ForegroundColor Cyan
    Write-Host "   N: Allow external senders to email the group." -ForegroundColor Cyan
    $requireAuth = Read-Host "Restrict the group to internal senders only? (Y/N)"
    $requireAuthBool = ($requireAuth -notmatch '^[Nn]$')

    # --------------------------------
    # Setting date and Notes
    # --------------------------------
    $date = Get-Date -Format "dd/MM/yyyy"
    $notes = "Created for Exclaimer on $date via manually run PowerShell script version $ps1Version."

    # --------------------------------
    # Object details for user verification
    # --------------------------------
    $details = [PSCustomObject]@{
        GroupName       = $groupName
        PrimarySmtp     = $smtp
        Alias           = $alias
        Owner           = $owner
        Members         = if ($members.Count -gt 0) { $members -join ", " } else { "(none)" }
        InternalOnly    = $requireAuthBool
        Date            = $date
        MemberList      = $members
        Notes           = $notes
        ps1Version      = $ps1Version
    }

    Write-Host "`nPlease review the collected details:`n" -ForegroundColor Cyan
    $displayDetails = $details | Select-Object GroupName, PrimarySmtp, Alias, Owner, Members, InternalOnly, Date
    $detailsString = $displayDetails | Format-List | Out-String
    Write-Host $detailsString.Trim()
    $confirm = Read-Host "`nAre these details correct? (Y/N)"
    if ($confirm -match "^[Yy]$") {
        return $details
    }
    else {
        Write-Host "Cancelled by user." -ForegroundColor Yellow
        return $null
    }
}

function createSecurityGroup {
    param($Details)
    Write-Host "`nCreating the mail-enabled security group..." -ForegroundColor Cyan
    Write-Host "Creating '$($Details.GroupName)'................... please wait" -ForegroundColor Yellow

    $paramsGroup = @{
        Name                               = $Details.GroupName
        DisplayName                        = $Details.GroupName
        Alias                              = $Details.Alias
        PrimarySmtpAddress                 = $Details.PrimarySmtp
        Type                               = "Security"
        ManagedBy                          = $Details.Owner
        Notes                              = $Details.Notes
        MemberJoinRestriction              = "Closed"
        MemberDepartRestriction            = "Closed"
        ErrorAction                        = "Stop"
    }

    try {
        New-DistributionGroup @paramsGroup | Out-Null
        Write-Host "[OK] Group created." -ForegroundColor Green
    } catch {
        Write-Host "[X] Failed to create the group: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    # Sender restriction is set separately so a failure here does not hide a successful create
    try {
        Set-DistributionGroup -Identity $Details.PrimarySmtp -RequireSenderAuthenticationEnabled $Details.InternalOnly -ErrorAction Stop
        Write-Host "[OK] Sender authentication set to internal only: $($Details.InternalOnly)" -ForegroundColor Green
    } catch {
        Write-Host "[!]  Group was created, but the sender restriction could not be set: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    return $true
}

function addGroupMembers {
    param($Details)
    if ($Details.MemberList.Count -eq 0) {
        Write-Host "`nNo members supplied. The group has been created empty." -ForegroundColor Yellow
        return
    }

    Write-Host "`nAdding members..." -ForegroundColor Cyan
    $added = 0
    $failed = @()
    foreach ($m in $Details.MemberList) {
        try {
            Add-DistributionGroupMember -Identity $Details.PrimarySmtp -Member $m -ErrorAction Stop
            Write-Host "   Added: $m" -ForegroundColor Green
            $added++
        } catch {
            Write-Host "   Failed: $m ($($_.Exception.Message))" -ForegroundColor Red
            $failed += $m
        }
    }
    Write-Host "Members added: $added of $($Details.MemberList.Count)" -ForegroundColor Cyan
    if ($failed.Count -gt 0) {
        Write-Host "These members were not added: $($failed -join ', ')" -ForegroundColor Yellow
    }
}

function verifyGroup {
    param($Details)
    Write-Host "`nVerifying the group..." -ForegroundColor Cyan
    $group = $null
    # New groups can take a short while to be visible, so retry briefly
    for ($i = 1; $i -le 5; $i++) {
        $group = Get-DistributionGroup -Identity $Details.PrimarySmtp -ErrorAction SilentlyContinue
        if ($group) { break }
        Write-Host "   Group not visible yet, retrying ($i/5)..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
    }

    if (-not $group) {
        Write-Host "[!]  The group could not be found yet. Allow a few minutes, then check the Exchange Admin Centre." -ForegroundColor Yellow
        return
    }

    $memberCount = @(Get-DistributionGroupMember -Identity $Details.PrimarySmtp -ResultSize Unlimited -ErrorAction SilentlyContinue).Count
    Write-Host "[OK] Group verified." -ForegroundColor Green
    $group | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, GroupType, ManagedBy, RequireSenderAuthenticationEnabled |
        Format-List | Out-String | Write-Host
    Write-Host "   Members: $memberCount" -ForegroundColor Cyan
}

function showCompletionNotice {
    param($Details)
    Write-Host ""
    Write-Host "CONFIGURATION COMPLETED" -ForegroundColor Green
    Write-Host ""
    Write-Host " The group may take a few minutes to propagate before it can be used in mail flow rules." -ForegroundColor Yellow
    Write-Host " Use this address when asked for the mail-enabled security group:" -ForegroundColor Yellow
    Write-Host "   $($Details.PrimarySmtp)" -ForegroundColor White
    Write-Host ""
    Write-Host " Manage the group in the Exchange Admin Centre:" -ForegroundColor Yellow
    Write-Host "   admin.exchange.microsoft.com -> Recipients -> Groups -> Mail-enabled security" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "You can now safely close PowerShell." -ForegroundColor Green
}

confirmExclaimerSupportApproval
checkExchangeOnlineModule
modernAuthConnect
checkExchangePermissions
$details = getGroupDetails
if (-not $details) {
    Write-Host "Script stopped. No changes made." -ForegroundColor Yellow
    return
}
if (-not (createSecurityGroup -Details $details)) {
    Write-Host "Script stopped. See the error above." -ForegroundColor Red
    return
}
addGroupMembers -Details $details
verifyGroup -Details $details
showCompletionNotice -Details $details