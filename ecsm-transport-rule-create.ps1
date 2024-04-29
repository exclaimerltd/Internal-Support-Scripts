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

function modern-auth-mfa-connect {
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline
}

 
function remove_previous {
    # Removes previous Transport Rules and Connectors

    # Does previous config exist?
    $tr = Get-TransportRule -Identity *exclaimer* -ErrorAction SilentlyContinue
    $rc = Get-InboundConnector -Identity *exclaimer* -ErrorAction SilentlyContinue
    $sc = Get-OutboundConnector -Identity *exclaimer* -ErrorAction SilentlyContinue

    If ($tr -eq $null -and $rc -eq $null -and $sc -eq $null) {
        write-host "out"
    }
    Else {
        Write-Host "Removing Exclaimer Transport Rule"
		foreach ($t in $tr){
        Remove-TransportRule -Identity $t.Name -Confirm:$false -ErrorAction SilentlyContinue
		}
        Write-Host "Removing Exclaimer Receive Connector"
        Remove-InboundConnector -Identity *Exclaimer* -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Removing Exclaimer Send Connector"
        Remove-OutboundConnector -Identity *Exclaimer* -Confirm:$false -ErrorAction SilentlyContinue
    }
}

function send_connector {
    # Creates the Send to Exclaimer Cloud send connector
    New-OutboundConnector -Name "Send to Exclaimer Cloud" `
    -Enabled $true `
    -UseMXRecord $false `
    -Comment $comment `
    -ConnectorType OnPremises `
    -ConnectorSource Default `
    -SmartHosts $smarthost `
    -TlsDomain $smarthost `
    -TlsSettings DomainValidation `
    -IsTransportRuleScoped $True `
    -RouteAllMessagesViaOnPremises $false `
    -CloudServicesMailEnabled $True `
    -AllAcceptedDomains $false `
    -TestMode $false
}

function receive_connector {
    # Creates the Receive from Exclaimer Cloud receive connector
    New-InboundConnector -Name "Receive from Exclaimer Cloud V2" `
    -Enabled $true `
    -ConnectorType OnPremises `
    -ConnectorSource Default `
    -Comment $comment `
    -SenderDomains smtp:* `
    -RequireTls $true `
    -RestrictDomainsToIPAddresses $false `
    -RestrictDomainsToCertificate $false `
    -CloudServicesMailEnabled $true `
    -TreatMessagesAsInternal $false `
    -TlsSenderCertificateName $accepteddomain
}

function transport_rule_create {
    # Group or all
    Write-Host ("")
    Write-Host ("============")
    $group = read-Host("Would you like to restrict this to a group? N/y")

    If ($group -eq "n" -or $group -eq "N") {
        # Creates transport rule for all
        New-TransportRule -Name "Identify messages to send to Exclaimer Cloud" `
        -Priority 0 `
        -Mode Enforce `
        -RuleErrorAction Ignore `
        -SenderAddressLocation Envelope `
        -RuleSubType None `
        -UseLegacyRegex $false `
        -FromScope InOrganization `
        -HasNoClassification $false `
        -AttachmentIsUnsupported $false `
        -AttachmentProcessingLimitExceeded $false `
        -AttachmentHasExecutableContent $false `
        -AttachmentIsPasswordProtected $false `
        -ExceptIfHasNoClassification $false `
        -ExceptIfHeaderMatchesMessageHeader X-ExclaimerHostedSignatures-MessageProcessed `
        -ExceptIfHeaderContainsMessageHeader "X-MS-Exchange-UnifiedGroup-SubmittedViaGroupAddress" `
        -ExceptIfHeaderContainsWords "{/o=ExchangeLabs/ou=Exchange Administrative Group}" `
        -ExceptIfHeaderMatchesPatterns "true" `
        -ExceptIfFromAddressMatchesPatterns '&lt;&gt;' `
        -ExceptIfMessageSizeOver 23592960 `
        -ExceptIfMessageTypeMatches Calendaring `
        -StopRuleProcessing $true `
        -RouteMessageOutboundRequireTls $false `
        -RouteMessageOutboundConnector "Send to Exclaimer Cloud"
    } 
    Else {
        $usegroup = Read-Host("Which group do you want to use? Add the email address for the mail enabled Security Group")

        # Creates transport rule for group    
        New-TransportRule -Name "Identify messages to send to Exclaimer Cloud" `
        -Priority 0 `
        -Mode Enforce `
        -RuleErrorAction Ignore `
        -SenderAddressLocation Envelope `
        -RuleSubType None `
        -UseLegacyRegex $false `
        -FromScope InOrganization `
        -FromMemberOf $usegroup `
        -HasNoClassification $false `
        -AttachmentIsUnsupported $false `
        -AttachmentProcessingLimitExceeded $false `
        -AttachmentHasExecutableContent $false `
        -AttachmentIsPasswordProtected $false `
        -ExceptIfHasNoClassification $false `
        -ExceptIfHeaderMatchesMessageHeader X-ExclaimerHostedSignatures-MessageProcessed `
        -ExceptIfHeaderContainsMessageHeader "X-MS-Exchange-UnifiedGroup-SubmittedViaGroupAddress" `
        -ExceptIfHeaderContainsWords "{/o=ExchangeLabs/ou=Exchange Administrative Group}" `
        -ExceptIfHeaderMatchesPatterns "true" `
        -ExceptIfFromAddressMatchesPatterns '&lt;&gt;' `
        -ExceptIfMessageSizeOver 23592960 `
        -ExceptIfMessageTypeMatches Calendaring `
        -StopRuleProcessing $true `
        -RouteMessageOutboundRequireTls $false `
        -RouteMessageOutboundConnector "Send to Exclaimer Cloud"
    }
}

function transport_rule_create_ooo {
        # Creates transport rule for all
        New-TransportRule -Name "Prevent Out of Office messages being sent to Exclaimer Cloud" `
        -Priority 0 `
        -Mode Enforce `
        -RuleErrorAction Ignore `
        -SenderAddressLocation Envelope `
        -RuleSubType None `
        -UseLegacyRegex $false `
        -MessageTypeMatches OOF `
        -HasNoClassification $false `
        -AttachmentIsUnsupported $false `
        -AttachmentProcessingLimitExceeded $false `
        -AttachmentHasExecutableContent $false `
        -AttachmentIsPasswordProtected $false `
        -ExceptIfHasNoClassification $false `
        -SetHeaderName "X-ExclaimerHostedSignatures-MessageProcessed"`
        -SetHeaderValue "true"
}



function allowed_ips {
    $iplist = @("104.210.80.79","13.70.157.244"`
    ,"52.233.37.155","52.242.32.10"`
    ,"51.4.231.63","51.5.241.184"`
    ,"104.40.229.156","52.169.0.179"`
    ,"52.172.222.27","52.172.38.8"`
    ,"51.140.37.132","51.141.5.228"`
    ,"191.237.4.149","104.209.35.28"`
    ,"20.52.124.58","20.113.192.118"`
    ,"20.233.10.24","20.74.156.16")
    Set-HostedConnectionFilterPolicy "Default" -IPAllowList $iplist
}

# User inputs
Write-Host "`nThe 'To get the Exclaimer domain name go to 'https://admin.exchange.microsoft.com/#/accepteddomains' or," -ForegroundColor Green
Write-Host "after you log in to your Exchange Admin Center go to Mail Flow -> Accepted domains" -ForegroundColor Green
$accepteddomain = Read-Host ("Please enter the xxxxxxxxxxxxxxxxxxxxx.smtp.excl.cloud domain here") 
write-host ("")
$region = Read-Host("Which region are you in?")
$smarthost = "smtp." + $region + "1.exclaimer.net"

# Comments
$date = (Get-Date -Format "dd/MM/yyyy")
$comment = "Connector created by Exclaimer Support on $date"

modern-auth-mfa-connect
remove_previous
send_connector
receive_connector
transport_rule_create
transport_rule_create_ooo
allowed_ips