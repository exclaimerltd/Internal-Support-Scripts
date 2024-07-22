## 
#<#
#.SYNOPSIS
#    Creates e Transport Rules to stop "Message Recalling" related messages from being routed through Exclaimer
#.DESCRIPTION
#    This is designed to fix issues with message recall.
#    This is achieved by creating 2 Transport Rules, one to stop recall messages from routing through Exclaimer,
#    and the other to stop the Microsoft 365 Message Recall reports from routing through Exclaimer.
#    After creating the Transport Rules, this will check if "TNEFEnabled" is set to "true" or not for the "Default" RemoteDomain,
#    if it is not, then it recommends setting "TNEFEnabled" for the internal domains.
#
#    Please refer to the REQUIREMENTS for the information needed to run this script correctly.
#.NOTES
#    Email: helpdesk@exclaimer.com
#    Date: 24th July 2024
#.PRODUCTS
#    Exclaimer Cloud - Signatures for Office 365
#.REQUIREMENTS
#    - The PowerShell "ExchangeOnlineManagement" module, will propmt to install if not present
#    - Global Administrator Account
#.HISTORY
#    1.0 - Creates 2 Transport Rules in Exchange Admin Center -> Mail Flow Rules
#          It will find all Accepted Domains, and set TNEFEnabled for Accepted domains, except for "Microsoft" and "Exclaimer" domains.
##>


#Getting Exchange Online Module
function checkExchangeOnline-Module {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        
        Write-Host "`nThe 'ExchangeOnlineManagement'PowerShell Module is already installed, continue...`n" -ForegroundColor Green
    } 
    else {
        $askModInstall = read-Host("Would you like to install the Powershell 'ExchangeOnlineManagement' and continue? N/y")
        if ($askModInstall -eq "n" -or $askModInstall -eq "N") {
            Write-Host "`nCannot continue without the 'ExchangeOnlineManagement'PowerShell Module, will now Exit..." -ForegroundColor Red
            Exit
        } else {
        Write-Host "`nContinue and install the 'ExchangeOnlineManagement'PowerShell Module before continuing..." -ForegroundColor Green
        pause
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
        }
    }
}

function modern-auth-mfa-connect {
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline
}


function remove_previous {

    # Does previous config exist?
    $tr = Get-TransportRule -Identity "Message Recall Fix - TR*" -ErrorAction SilentlyContinue

    If ($tr -eq $null) {
        write-host "Please wait..." -ForegroundColor Yellow
    }
    Else {
        Write-Host "Removing the pre-existing 'Message Recall Fix - TR*' rules before continuing..."
		foreach ($t in $tr){
        Remove-TransportRule -Identity $t.Name -Confirm:$false -ErrorAction SilentlyContinue
		}
    }
}

function transport_rule_tr {
    write-host "Creating Transport Rules..."  -ForegroundColor Green
    Write-Host "`n============ Creating the first rule, please wait ============" -ForegroundColor Yellow

        New-TransportRule -Name "Message Recall Fix - TR1" `
        -Priority 0 `
        -Mode Enforce `
        -RuleErrorAction Ignore `
        -SenderAddressLocation Envelope `
        -RuleSubType None `
        -UseLegacyRegex $false `
        -HeaderContainsMessageHeader 'x-ms-exchange-recallreportgenerated' `
        -HeaderContainsWords "true", "false" `
        -HasNoClassification $false `
        -AttachmentIsUnsupported $false `
        -AttachmentProcessingLimitExceeded $false `
        -AttachmentHasExecutableContent $false `
        -AttachmentIsPasswordProtected $false `
        -ExceptIfHasNoClassification $false `
        -SetHeaderName 'X-ExclaimerHostedSignatures-MessageProcessed' `
        -SetHeaderValue 'true' `
        -Comments $tr_comment
        
    Write-Host "============ Almost there, creating the second rule ============`n" -ForegroundColor Yellow

        New-TransportRule -Name "Message Recall Fix - TR2" `
        -Priority 1 `
        -Mode Enforce `
        -RuleErrorAction Ignore `
        -SenderAddressLocation Envelope `
        -RuleSubType None `
        -UseLegacyRegex $false `
        -HeaderContainsMessageHeader 'X-MS-Exchange-Generated-Message-Source' `
        -HeaderContainsWords 'Transport Message Recall Routing Agent' `
        -HasNoClassification $false `
        -AttachmentIsUnsupported $false `
        -AttachmentProcessingLimitExceeded $false `
        -AttachmentHasExecutableContent $false `
        -AttachmentIsPasswordProtected $false `
        -ExceptIfHasNoClassification $false `
        -SetHeaderName 'X-ExclaimerHostedSignatures-MessageProcessed' `
        -SetHeaderValue 'true' `
        -Comments $tr_comment
}


function remote_domain_tnef {
    $domainCount = 0
    $countNoRemoteDomain = 0
    [array]$remoteDomains = @()
    $tnef_check = Get-RemoteDomain "Default" | Select-Object TNEFEnabled
    #Set-RemoteDomain "Default" -TNEFEnabled $true
    Write-Host "`nNow checking if TNEF is enabled for internal domains..." -ForegroundColor Green

    if ($tnef_check.TNEFEnabled -eq "true") {
        Write-Host "`n'TNEFEnabled' is already set to '$true' for the 'Default' RemoteDomain, no further action required..." -ForegroundColor Green    
    } else {
    # Get the accepted domains excluding certain patterns
    $acceptedDomains = Get-AcceptedDomain | Where-Object {
        $_.DomainName -notlike "*.excl.cloud" -and $_.DomainName -notlike "*.onmicrosoft.com"
    }
    
    foreach ($acceptedDomain in $acceptedDomains) {
        $domainCount++
        
        # Fetch remote domain details
        $remoteDomain = Get-RemoteDomain | Where-Object {
            $_.DomainName -eq $acceptedDomain.DomainName
        }

        # Create a new object with the details
        $domainDetails = New-Object psobject -Property @{
            ID = $domainCount
            AcceptedDomain = $acceptedDomain.DomainName
            RemoteDomain = if ($remoteDomain) { $remoteDomain.Name } else { "None" }
            TNEFEnabled = if ($remoteDomain) { $remoteDomain.TNEFEnabled } else { $false }
        }

        # Add the new object to the array
        $remoteDomains += $domainDetails
    }
    
    Write-Host "`nWe recommend that you configure a 'RemoteDomain' with 'TNEFEnabled' enabled for any internal domains you wish to recall messages for..." -ForegroundColor Red
    Write-Host "`nSee the'Resolution' options in the article: https://support.exclaimer.com/hc/en-gb/articles/6739198496413"

    Write-Host "`nListing 'AcceptedDomains' that do not have 'TNEFEnabled' set to'$true'..." -ForegroundColor Green
    # Output the result
    return $remoteDomains | Select-Object ID, AcceptedDomain, RemoteDomain,TNEFEnabled
    }
}

# Comments
$date = (Get-Date -Format "dd/MM/yyyy")
$tr_comment = "Created by Exclaimer Support on $date `nStops message recall related emails from being routed through Exclaimer"

checkExchangeOnline-Module
modern-auth-mfa-connect
remove_previous
transport_rule_tr
remote_domain_tnef