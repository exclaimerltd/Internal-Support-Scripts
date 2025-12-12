# 
<#
.SYNOPSIS
    Creates the rules and connectors needed for Exclaimer Cloud
.DESCRIPTION
    This is designed to be run after you have completed the steps to get the certificate domain.
    The Transport Rule is created in a disabled state so this can be run prior to the steps carried out by Sys Eng.

    Please refer to the REQUIREMENTS for the information needed to run this script correctly.
.NOTES
    Email: helpdesk@exclaimer.com
    Date: 24th July 2018
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    - SMTP domain needs to already be added to Office 365 and verified
    - Group for Transport Rule
    - Global Administrator Accounts
.HISTORY
    1.0 - Creates transport rule, connectors and sets up Exclaimer Cloud in a enabled state
    1.1 - Corrected issue relating to date/time, requested email address for group, added a 1 to the region request.
    1.2 - Added allowed ip list update
    2.0 - Removal of previous configuration
#>
$Global:scriptVersion = "'1.25.1209'"
Add-Type -AssemblyName PresentationFramework
#Getting Exchange Online Module
function confirmExclaimerSupportApproval {
    Write-Host ""
    Write-Host "IMPORTANT NOTICE" -ForegroundColor Yellow
    Write-Host "This script must only be run when advised by the Exclaimer Technical Support Team." -ForegroundColor Yellow
    Write-Host "Running it without guidance can alter mail flow, remove existing connectors, or apply" -ForegroundColor Yellow
    Write-Host "configuration changes that may affect email delivery." -ForegroundColor Yellow
    Write-Host ""

    $confirm = Read-Host "Do you understand and accept these conditions? (Y/N)"

    if ($confirm.ToUpper() -ne "Y") {
        Write-Host "Operation cancelled. No changes have been made." -ForegroundColor Red
        exit
    }
}

    function checkExchangeOnlineModule {
        if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
            Write-Host "✅ Exchange Online Management module is already installed." -ForegroundColor Green
            #return $true
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
                } catch {
                    Write-Host "❌ Failed to install the module: $($_.Exception.Message)" -ForegroundColor Red
                    Add-Content $FullLogFilePath "<p class='warning'>Exchange Online Management module installation failed: $([System.Web.HttpUtility]::HtmlEncode($_.Exception.Message))</p>"
                }
            } else {
                Write-Host "⚠️ Skipping module installation. Admin access required for automated mailbox queries." -ForegroundColor Yellow
                Add-Content $FullLogFilePath "<p class='warning'>User skipped Exchange Online module installation. Manual Add-in version collection required.</p>"
            }
        }
    }

function modern-auth-mfa-connect {    
    Write-Host "   You will be prompted to Sign in with Microsoft in order to continue." -ForegroundColor Yellow
    Start-Sleep -Seconds 3
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline
}

function remove_previous {
    Write-Host "`nChecking for previous configuration..." -ForegroundColor Cyan
    # Removes previous Transport Rules and Connectors
    $tr = Get-TransportRule -Identity *exclaimer* -ErrorAction SilentlyContinue
    $rc = Get-InboundConnector -Identity *exclaimer* -ErrorAction SilentlyContinue
    $sc = Get-OutboundConnector -Identity *exclaimer* -ErrorAction SilentlyContinue

    if (-not $tr -and -not $rc -and -not $sc) {
        Write-Host "No existing Rules or Connectors found" -ForegroundColor Green
        return
    }
    if ($tr) {
        Write-Host "Removing existing Exclaimer Transport Rules..... please wait" -ForegroundColor Yellow
        foreach ($t in $tr) {
            Remove-TransportRule -Identity $t.Name -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
    if ($rc) {
        Write-Host "Removing existing Exclaimer Receive Connector... please wait" -ForegroundColor Yellow
        Remove-InboundConnector -Identity $rc.Identity -Confirm:$false -ErrorAction SilentlyContinue
    }
    if ($sc) {
        Write-Host "Removing existing Exclaimer Send Connector...... please wait" -ForegroundColor Yellow
        Remove-OutboundConnector -Identity $sc.Identity -Confirm:$false -ErrorAction SilentlyContinue
    }
    Write-Host "Previous configuration cleanup completed.`n" -ForegroundColor Green
}

function configure_exclaimer_connectors {
    Write-Host "Configuring the new Exclaimer Connectors..." -ForegroundColor Cyan
    Start-Sleep -Seconds 10
    Write-Host "Configuring 'Send to Exclaimer Cloud'........... please wait" -ForegroundColor Yellow
    # ----------------------------------------------------------
    # CONNECTOR 1: "Send to Exclaimer Cloud"
    # ----------------------------------------------------------
    New-OutboundConnector -Name "Send to Exclaimer Cloud" `
        -Enabled $true `
        -UseMXRecord $false `
        -Comment $Details.Comment `
        -ConnectorType OnPremises `
        -ConnectorSource Default `
        -SmartHosts $Details.Smarthost `
        -TlsDomain $Details.Smarthost `
        -TlsSettings DomainValidation `
        -IsTransportRuleScoped $True `
        -RouteAllMessagesViaOnPremises $false `
        -CloudServicesMailEnabled $True `
        -AllAcceptedDomains $false `
        -TestMode $false |
    Out-Null

    Write-Host "Configuring 'Receive from Exclaimer Cloud V2'... please wait" -ForegroundColor Yellow
    # ----------------------------------------------------------
    # CONNECTOR 2: "Receive from Exclaimer Cloud V2"
    # ----------------------------------------------------------
    New-InboundConnector -Name "Receive from Exclaimer Cloud V2" `
        -Enabled $true `
        -ConnectorType OnPremises `
        -ConnectorSource Default `
        -Comment $Details.Comment `
        -SenderDomains smtp:* `
        -RequireTls $true `
        -RestrictDomainsToIPAddresses $false `
        -RestrictDomainsToCertificate $false `
        -CloudServicesMailEnabled $true `
        -TreatMessagesAsInternal $false `
        -TlsSenderCertificateName $Details.ExclDomain |
    Out-Null
    Write-Host "Exclaimer Connectors configuration completed." -ForegroundColor Green
}

function transport_rules_create {
    Write-Host "`nConfiguring the Exclaimer Transport Rules..." -ForegroundColor Cyan
    # Use values already collected in getConfigDetails
    $groupFilter = $Details.GroupFilter
    # ----------------------------------------------------------
    # RULE 1: Identify messages to send to Exclaimer Cloud
    # ----------------------------------------------------------
    Write-Host "Configuring 'Identify messages to send to Exclaimer Cloud'................... please wait" -ForegroundColor Yellow

    $paramsIdentify = @{
        Name                                  = "Identify messages to send to Exclaimer Cloud"
        Priority                              = 0
        Mode                                  = "Enforce"
        RuleErrorAction                       = "Ignore"
        SenderAddressLocation                 = "Envelope"
        RuleSubType                           = "None"
        UseLegacyRegex                        = $false
        FromScope                             = "InOrganization"
        HasNoClassification                   = $false
        AttachmentIsUnsupported               = $false
        AttachmentProcessingLimitExceeded     = $false
        AttachmentHasExecutableContent        = $false
        AttachmentIsPasswordProtected         = $false
        ExceptIfHasNoClassification           = $false
        ExceptIfHeaderMatchesMessageHeader    = "X-ExclaimerHostedSignatures-MessageProcessed"
        ExceptIfHeaderContainsMessageHeader   = "X-MS-Exchange-UnifiedGroup-SubmittedViaGroupAddress"
        ExceptIfHeaderContainsWords           = "{/o=ExchangeLabs/ou=Exchange Administrative Group}"
        ExceptIfHeaderMatchesPatterns         = "true"
        ExceptIfFromAddressMatchesPatterns    = "&lt;&gt;"
        ExceptIfMessageSizeOver               = 23592960
        ExceptIfMessageTypeMatches            = "Calendaring"
        StopRuleProcessing                    = $true
        RouteMessageOutboundRequireTls        = $false
        RouteMessageOutboundConnector         = "Send to Exclaimer Cloud"
        Comments                              = $Details.Comment
    }
    # Apply group restriction only if provided
    if ($groupFilter) {
        $paramsIdentify["FromMemberOf"] = $groupFilter
    }
    New-TransportRule @paramsIdentify | Out-Null
    # Disable rule by default
    Disable-TransportRule -Identity "Identify messages to send to Exclaimer Cloud" -Confirm:$false

    # ----------------------------------------------------------
    # RULE 2: Prevent Out of Office messages being sent to Exclaimer
    # ----------------------------------------------------------
    Write-Host "Configuring 'Prevent Out of Office messages being sent to Exclaimer Cloud'... please wait" -ForegroundColor Yellow

    $paramsOOO = @{
        Name                                  = "Prevent Out of Office messages being sent to Exclaimer Cloud"
        Priority                              = 0
        Mode                                  = "Enforce"
        RuleErrorAction                       = "Ignore"
        SenderAddressLocation                 = "Envelope"
        RuleSubType                           = "None"
        UseLegacyRegex                        = $false
        MessageTypeMatches                    = "OOF"
        HasNoClassification                   = $false
        AttachmentIsUnsupported               = $false
        AttachmentProcessingLimitExceeded     = $false
        AttachmentHasExecutableContent        = $false
        AttachmentIsPasswordProtected         = $false
        ExceptIfHasNoClassification           = $false
        SetHeaderName                         = "X-ExclaimerHostedSignatures-MessageProcessed"
        SetHeaderValue                        = "true"
        Comments                              = $Details.Comment
    }
    New-TransportRule @paramsOOO | Out-Null
    Write-Host "Exclaimer Transport Rules configuration completed." -ForegroundColor Green
}

function allowed_ips {
    Write-Host "`nAdding Exclaimer's regional IPs to IPAllowList... please wait" -ForegroundColor Cyan
    $region = $Details.region.ToUpper().Trim()
    $regionMap = @{
        "AU"  = @("104.210.80.79","13.70.157.244")
        "CA"  = @("52.233.37.155","52.242.32.10")
        "EU"  = @("104.40.229.156","52.169.0.179")
        "DE"  = @("20.52.124.58","20.113.192.118")
        "UAE" = @("52.172.222.27","52.172.38.8","20.233.10.24","20.74.156.16")
        "UK"  = @("51.140.37.132","51.141.5.228")
        "US"  = @("191.237.4.149","104.209.35.28")
    }    
    $filteredList = $regionMap[$region]
    Set-InboundConnector -Identity "Receive from Exclaimer Cloud v2" -EFSkipLastIP $true
    Set-HostedConnectionFilterPolicy "Default" -IPAllowList $filteredList
    Write-Host "Exclaimer added to 'IPAlowList', completed." -ForegroundColor Green
}

# User inputs
function getConfigDetails {    
    Clear-Host
    $ps1Version = $Global:scriptVersion
    Write-Host "`nRequesting an email address to find your Exclaimer Region..." -ForegroundColor Yellow
    $email = (Read-Host "Enter your Microsoft 365 tenant Global Admin email address").Trim().ToLower()
    if ($email -notmatch '^[^@]+@[^@]+\.[^@]+$') { Write-Host "Invalid email format." -ForegroundColor Red; return }
    $emailDomain = $email.Split("@")[1]

    $hostToTest = "outlookclient.exclaimer.net"
    Write-Host "`nChecking connectivity to $hostToTest..."
    if (-not (Test-NetConnection -ComputerName $hostToTest -Port 443 -InformationLevel Quiet)) {
        Write-Host "Unable to connect to $hostToTest on port 443." -ForegroundColor Red; $region = $null; return
    }

    Write-Host "Connectivity OK." -ForegroundColor Green
    Write-Host "Fetching region data for domain '$emailDomain'..."
    $url = "https://$hostToTest/cloudgeolocation/$emailDomain"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get -ErrorAction Stop
        $endpoint = $response.OutlookSignaturesEndpoint
        if ($endpoint) {
            $endpoint = ($endpoint.Trim() -replace '^https://','').TrimEnd('/')
            $region = $endpoint.Split(".")[0]
            Write-Host "Region identified: $region" -ForegroundColor Green
        }
        else { Write-Host "Region endpoint missing for this domain." -ForegroundColor Yellow; $region = $null }
    }
    catch { Write-Host "No region data found for domain '$emailDomain'." -ForegroundColor Red; $region = $null }

    if ($region -eq "dog") { $smarthost = "smtp.beta.exclaimer.net" }
    elseif ($region) { $smarthost = "smtp.$region`1.exclaimer.net" }
    else { $smarthost = $null }

    # ----------------------------
    # Getting the Exclaimer Ceertificate Domain
    # ----------------------------
    Write-Host "`nChecking if the Exclaimer Domain is present in your Microsoft 365 Tenant..." -ForegroundColor Yellow
    $exclDomains = Get-AcceptedDomain |
        Where-Object { $_.DomainName -like "*.excl.cloud" } |
        Select-Object DomainName, WhenCreated

    if (-not $exclDomains -or $exclDomains.Count -eq 0) {
        Write-Host "No *.excl.cloud domain found in this tenant. Script cannot continue." -ForegroundColor Red
        Write-Host ""
        Write-Host "Please ensure the Exclaimer tenant is correctly connected to your Microsoft 365 environment." -ForegroundColor Yellow
        Write-Host "1. Re-authorize 'Enable Azure AD Access' via the Exclaimer Portal:" -ForegroundColor Cyan
        Write-Host "   https://support.exclaimer.com/hc/en-gb/articles/360018691838-Enable-Azure-AD-Access" -ForegroundColor White
        Write-Host "2. Attempt to configure Mail Flow ('Connect to Microsoft 365') via the Exclaimer Portal:" -ForegroundColor Cyan
        Write-Host "   https://support.exclaimer.com/hc/en-gb/articles/17192747888669-Connect-to-Microsoft-365" -ForegroundColor White
        Write-Host ""
        Write-Host "Once these steps are completed, allow 5 minutes and re-run the script, only if the Mail Flow configuration fails." -ForegroundColor Yellow
        return
    }

    if ($exclDomains.Count -gt 1) {
        Write-Host "`nMultiple Exclaimer domains found. Please select one:`n" -ForegroundColor Yellow
        for ($i = 0; $i -lt $exclDomains.Count; $i++) {
            Write-Host "[$($i+1)] $($exclDomains[$i].DomainName)   Created: $($exclDomains[$i].WhenCreated)"
        }

        $selection = Read-Host "`nEnter the number for the domain you want to use"
        if ($selection -notmatch '^[0-9]+$' -or [int]$selection -lt 1 -or [int]$selection -gt $exclDomains.Count) {
            Write-Host "Invalid selection. Script cannot continue." -ForegroundColor Red
            return
        }

        $exclDomain = $exclDomains[[int]$selection - 1].DomainName
    } else {
        $exclDomain = $exclDomains[0].DomainName
    }

    Write-Host "Found Exclaimer domain: $exclDomain" -ForegroundColor Green

    # ----------------------------
    # Group restriction for Transport Rules
    # ----------------------------
    Write-Host "`nTransport Rule Options" -ForegroundColor Cyan
    $restrictToGroup = Read-Host "Would you like to restrict this to a group? N/y"

    $groupFilter = $null

    if ($restrictToGroup -match '^[Yy]$') {
        $groupFilter = Read-Host "Enter the email address of the mail enabled security group"

        if ([string]::IsNullOrWhiteSpace($groupFilter)) {
            Write-Host "No group provided. Cancelling config details collection." -ForegroundColor Red
            return $null
        }
    }

    # --------------------------------
    # Setting date and Comments
    # --------------------------------
    $date = Get-Date -Format "dd/MM/yyyy"
    $comment = "Created for Exclaimer on $date via manually run PowerShell script version $ps1Version."

    # --------------------------------
    # Object details for user verification
    # --------------------------------
    $details = [PSCustomObject]@{
        Email       = $email
        Domain      = $emailDomain
        Region      = $region
        Smarthost   = $smarthost
        ExclDomain  = $exclDomain
        GroupUsed   = $restrictToGroup
        GroupFilter = $groupFilter
        Date        = $date
        Comment     = $comment
        ps1Version  = $ps1Version
    }

    Write-Host "`nPlease review the collected details:`n" -ForegroundColor Cyan
    $displayDetails = $details | Select-Object Email,Domain,Region,Smarthost,ExclDomain,GroupUsed,GroupFilter,Date
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
function showExclaimerCompletionNotice {
    Write-Host ""
    Write-Host "CONFIGURATION COMPLETED" -ForegroundColor Green
    Write-Host ""
    Write-Host " The changes you applied may take a few minutes to propagate before taking effect." -ForegroundColor Yellow
    Write-Host " If you encounter any Mail Flow issues, you can disable the Transport Rule" -ForegroundColor Yellow
    Write-Host "  'Identify messages to send to Exclaimer Cloud' in the Exchange Admin Centre:" -ForegroundColor Yellow
    Write-Host "  admin.exchange.microsoft.com -> Mail Flow -> Rules" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Please review the above instructions." -ForegroundColor Cyan

    # First prompt
    $ack = Read-Host "Have you read and understood these instructions? Press Y to acknowledge"

    if ($ack.ToUpper() -ne "Y") {
        # Second and final prompt
        Write-Host "`nThe Transport Rule 'Identify messages to send to Exclaimer Cloud'is currently Disabled." -ForegroundColor Yellow
        Write-Host "`n Y: Enable the Transport Rule before finishing." -ForegroundColor Cyan
        Write-Host " N: Keep the Transport Rule Disabled.`n" -ForegroundColor Cyan
        $ack = Read-Host "Please confirm. Type Y or N to continue without enabling the Transport Rule"
    }

    if ($ack.ToUpper() -eq "Y") {
        Write-Host "`nAcknowledged." -ForegroundColor Green
        Write-Host "Enabling Transport Rule 'Identify messages to send to Exclaimer Cloud... please wait" -ForegroundColor Yellow
        Enable-TransportRule -Identity "Identify messages to send to Exclaimer Cloud" -Confirm:$false
        Write-Host "Transport Rule 'Identify messages to send to Exclaimer Cloud' now Enabled." -ForegroundColor Cyan
        Write-Host "`nYou can now safely close PowerShell." -ForegroundColor Green
    }
    else {
        Write-Host "`nAcknowledgment not received. Please read the below carefully before closing PowerShell." -ForegroundColor Red
        Write-Host "`nThe Transport Rule 'Identify messages to send to Exclaimer Cloud' will stay Disabled." -ForegroundColor Yellow
        Write-Host "You can manually enable the Transport Rule" -ForegroundColor Yellow
        Write-Host "  'Identify messages to send to Exclaimer Cloud' from the Exchange Admin Centre:" -ForegroundColor Yellow
        Write-Host "  admin.exchange.microsoft.com -> Mail Flow -> Rules" -ForegroundColor Yellow
    }
}

confirmExclaimerSupportApproval
checkExchangeOnlineModule
modern-auth-mfa-connect
$details = getConfigDetails
if (-not $details) {
    Write-Host "Script stopped. No changes made." -ForegroundColor Yellow
    return
}
remove_previous
configure_exclaimer_connectors -Details $details
transport_rules_create -Details $details
allowed_ips -Details $details
showExclaimerCompletionNotice