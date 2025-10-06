## 
#<#
#.SYNOPSIS
#    Reconfigure Third-Paty outbound connector (Mimecast, Proofpoint, Barracuda, or other) to be managed by a Transport Rule
#.DESCRIPTION
#    This is designed to reconfigure Third-Paty outbound connector (Mimecast, Proofpoint, Barracuda, or other) to be managed by a Transport Rule.
#    This is achieved by achieving the configuration described in the article below, but through PowerShell.
#    https://support.exclaimer.com/hc/en-gb/articles/4405851491101
#
#    Please refer to the REQUIREMENTS for the information needed to run this script correctly.
#.NOTES
#    Email: helpdesk@exclaimer.com
#    Date: 8th August 2024
#.PRODUCTS
#    Exclaimer Cloud - Signatures for Office 365
#.REQUIREMENTS
#    - The PowerShell "ExchangeOnlineManagement" module, will propmt to install if not present
#    - Global Administrator Account
#.HISTORY
#    1.0 - Changes Outbound Connector usage
#        - Configures a Transport Rule to manage that connector
#        - Ensures the Transport Rule "Identify messages to send to Exclaimer Cloud" is configured to "Stop processing more rules"
#        - This will ensure that your emails are processed through Exclaimer before being routed by other Transport Rules.
##>


#Getting Exchange Online Module

function infomative {
Write-Host "`nThis script is to be used only if you use a Third-Party Security Solution such as`
Mimecast, Proofpoint, Barracuda, or other, which you route your ""Outbound"" emails through.`
Only use it if you have connector configured to route emails from ""O365"" to ""Third-Party Security Solution (Mimecast, Proofpoint, Barracuda, etc)""`
`n`nFor more information, see article: 'https://support.exclaimer.com/hc/en-gb/articles/4405851491101'`n" -ForegroundColor Yellow

Write-Host "`nThis is NOT required if you use any of the Third-Party solutions for ""Inbound"" emails only.`n" -ForegroundColor RED

doContinue
}


function doContinue {
        $doContinue = Read-Host "Would you like to continue? (y/n)"
        if ($doContinue -eq "y" -OR $doContinue -eq "Y"){
            Write-Host "Checking if require PowerShell Module is installed...." -ForegroundColor Yellow
            checkExchangeOnlineModule
        }
        else
        {
            Write-Host "Will now disconnect and exit..." -ForegroundColor Red
            endSession
        }
}


function checkExchangeOnlineModule {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {        
        Write-Host "`nThe 'ExchangeOnlineManagement'PowerShell Module is already installed, lets sign in as the Global Admin...`n" -ForegroundColor Green
        modernAuthConnect
    } 
    else {
        $askModInstall = read-Host("Would you like to install the Powershell 'ExchangeOnlineManagement' and continue? N/y")
        if ($askModInstall -eq "n" -or $askModInstall -eq "N") {
            Write-Host "`nCannot continue without the 'ExchangeOnlineManagement'PowerShell Module, will now Exit." -ForegroundColor Red
            Exit
        } else {
        Write-Host "`nContinue and install the 'ExchangeOnlineManagement'PowerShell Module before continuing..." -ForegroundColor Green
        pause
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
        modernAuthConnect
        }
    }
}

function modernAuthConnect {
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline
}

function gatherInfo {
    $global:provider = Read-Host ("Who is your Email Security provider? (i.e Barracuda, Mimecast, Proofpoint, etc...)")
    $global:outboundConnectorName = Read-Host ("What is the full and exact name of the Outbound connector for $global:provider)?")
    Write-Host "`nThird-party provider......: $global:provider" -ForegroundColor Green
    Write-Host "Connector name............: $global:outboundConnectorName" -ForegroundColor Green
    thirdPartyProvider
}

function thirdPartyProvider {
    $outboundConnector = Get-OutboundConnector -Identity $global:outboundConnectorName | Select-Object Name,Identity,Enabled,SmartHosts,IsTransportRuleScoped -ErrorAction SilentlyContinue
    $thirdPartyTransportRule = Get-TransportRule | Where {$_.RouteMessageOutboundConnector -eq $global:outboundConnectorName}
    $exclaimerTransportRule = Get-TransportRule | Where {$_.Name -like "Identify messages to send to Exclaimer Cloud"}
    $global:thirdPartyTransportRuleName = $thirdPartyTransportRule.Name
    $global:thirdPartyTransportRulePriority = $thirdPartyTransportRule.Priority
    $global:exclaimerTransportRuleName = $exclaimerTransportRule.Name
    $global:exclaimerTransportRulePriority = $exclaimerTransportRule.Priority
    $global:exclaimerTransportRuleStopMoreRules = $exclaimerTransportRule.StopRuleProcessing
    if ($outboundConnector){ # If a connector is found
        Write-Host "`nConnector found...`n" -ForegroundColor Yellow
                if ($outboundConnector.IsTransportRuleScoped -ne $False){ #If already managed by a Transport Rule 
                    Write-Host "`nConnector is already managed by a Transport Rule" -ForegroundColor Green
                    IsTransportRuleScoped
                }
                else #If not managed by a Transport Rule 
                {
                    notIsTransportRuleScoped
                }
                }        
            else # If a connector is not found
            {
                Write-Host "`nNo connector was found by the name provided...`nYou can choose to continue (y) to try again, or not (n) to end this session. " -ForegroundColor Yellow
                doTryAgain
        }

}

function IsTransportRuleScoped {
    Write-Host "This connector is already managed by a transport rule ""$thirdPartyTransportRule""." -ForegroundColor Green
        if ($global:thirdPartyTransportRulePriority -gt $global:exclaimerTransportRulePriority){
            Write-Host "No further action required." -ForegroundColor Green
            endSession
        }
        else {
            Write-Host "`nBut the order (priority) of the Transport Rules is not correct..." -ForegroundColor Red
            Write-Host "Please update the Transport Rule ""$thirdPartyTransportRule"" so that it is of lower priority (higher number) `nthan the Transport Rule ""Identify messages to send to Exclaimer Cloud""." -ForegroundColor Red
            Write-Host "`nThe Transport Rule ""$global:thirdPartyTransportRuleName"" is currently priority ""$global:thirdPartyTransportRulePriority""" -ForegroundColor Red
            Write-Host "The Transport Rule ""$global:exclaimerTransportRuleName"" is currently priority ""$global:exclaimerTransportRulePriority""" -ForegroundColor Yellow
        }

}

function notIsTransportRuleScoped {
    Write-Host "`nConnector is not managed by a Transport Rule" -ForegroundColor Red
    Write-Host "`nThe configuration of the connector $global:outboundConnectorName needs to be updated, so it can be managed by a Transport Rule..." -ForegroundColor Yellow
    $doUpdateConnector = Read-Host "Would you like to continue? (y/n)"
            if ($doUpdateConnector -eq "y" -OR $doUpdateConnector -eq "Y"){
                updateOutboundConnector
            }
            else
            {
                endSession
            }
}


function updateOutboundConnector {
    Write-Host "`n============ Updating the connector '$global:outboundConnectorName'...." -ForegroundColor Yellow
    Set-OutboundConnector -Identity $global:outboundConnectorName `
    -IsTransportRuleScoped $True `
    -RecipientDomains @() `
    -Enabled $True `
    -Comment $cn_comment
    Write-Host "`n============ Connector '$global:outboundConnectorName' now updated." -ForegroundColor Green
    createThirdPartyTR
}

function createThirdPartyTR {
    Write-Host "`n============ Creating Transport Rule for connector '$global:outboundConnectorName'...." -ForegroundColor Yellow

    New-TransportRule -Name "Route emails through $global:provider" `
    -Mode Enforce `
    -RuleErrorAction Ignore `
    -FromScope InOrganization `
    -SentToScope NotInOrganization `
    -RouteMessageOutboundConnector $global:outboundConnectorName `
    -SenderAddressLocation Envelope `
    -RuleSubType None `
    -UseLegacyRegex $false `
    -HasNoClassification $false `
    -AttachmentIsUnsupported $false `
    -AttachmentProcessingLimitExceeded $false `
    -AttachmentHasExecutableContent $false `
    -AttachmentIsPasswordProtected $false `
    -ExceptIfHasNoClassification $false `
    -Comments $tr_comment
    Write-Host "`n============ Transport Rule for connector '$global:outboundConnectorName' successfully  created." -ForegroundColor Green
    if ($global:provider -like "Barracuda") {
        Write-Host "`n============ $global:provider does not support Out of Office emails..."  -ForegroundColor Yellow
        Write-Host "`n============ Updating the Transport Rule ""Route emails through $global:provider""..."  -ForegroundColor Yellow
        excludeAutomaticEmails
    }
}

function excludeAutomaticEmails {
    Set-TransportRule -Identity "Route emails through $global:provider" `
    -ExceptIfMessageTypeMatches "OOF"
    Write-Host "`n============ Successfully updated Transport Rule ""Route emails through $global:provider"" to exclude Out of Office emails...." -ForegroundColor Green

}

function checkExclTR {
    if ($global:exclaimerTransportRuleStopMoreRules -eq $True) {
        Write-Host "`nThe Transport Rule ""Identify messages to send to Exclaimer Cloud"" is correctly configured." -ForegroundColor Green
    }
    else {
        Write-Host "============ The Transport Rule ""Identify messages to send to Exclaimer Cloud"" is not correctly configured to stop processing more rules..." -ForegroundColor Yellow
        Write-Host "============ Updating Transport Rule ""Identify messages to send to Exclaimer Cloud"" to stop processing more rules...`n============ This will ensure that your emails are processed through Exclaimer before being routed by other Transport Rules...." -ForegroundColor Yellow
        Set-TransportRule -Identity "Identify messages to send to Exclaimer Cloud" `
        -StopRuleProcessing $True       
        Write-Host "============ Successfully updated Transport Rule ""Identify messages to send to Exclaimer Cloud"", now correctly configured to stop processing more rules." -ForegroundColor Green

    }

}

function doTryAgain {
        $doTryAgain = Read-Host "Would you like to try again? (y/n)"
        if ($doTryAgain -eq "y" -OR $doTryAgain -eq "Y"){
            gatherInfo
        }
        else
        {
            Write-Host "Will now disconnect and exit." -ForegroundColor Red
            endSession
        }
}

function endSession {            
            # Disconnecting from Exchange Online
            Write-Host "`nWill now disconnect and exit." -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false
            Write-Host "Session Ended`n" -ForegroundColor Red
            Start-Sleep -Seconds 5
            Exit
}

# Comments
$date = (Get-Date -Format "dd/MM/yyyy")
$tr_comment = "Created by Exclaimer Support PowerShell script on $date `nRoutes messages through the connector '$global:outboundConnectorName'"
$cn_comment = "Updated by Exclaimer Support PowerShell script on $date `nThis connector is now managed by Transport Rule ""Route emails through $provider"""

infomative
gatherInfo
checkExclTR
endSession