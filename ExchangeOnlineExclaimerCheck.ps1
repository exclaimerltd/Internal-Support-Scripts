#﻿<#
#.SYNOPSIS
#    Checks and outputs the configuration of Transport Rule and Connectors that may affected routing of emails to Exclaimer.
#
#.DESCRIPTION
#    It will first prompt to login with Microsoft, then it will check and outputs the configuration of Transport Rule and Connectors that may affected routing of emails to Exclaimer.
#    Output filename is "ExchangeOnlineExclaimerCheck.txt"
#
#.NOTES
#    Email: helpdesk@exclaimer.com
#    Date: 27th June 2021
#
#.PRODUCTS
#    Exclaimer Signature Management - Microsoft 365
#
#.REQUIREMENTS
#    - Global Administrator access to the Microsoft Tenant
#    - Requires path "C:\Temp\"
#    - ExchangeOnlineManagement - https://www.powershellgallery.com/packages/ExchangeOnlineManagement/0.4578.0
#
#.VERSION
#
#
#	1.0.1
#		- Added call to get other transport rules
#		- Added check for Out of Office Transport Rule
#		- Added "Priority" to info collected from Transport Rules
#		- Conditioned getting of Transport Rules output by pre-checking for its existence (avoiding errors)
#		
#	1.0.0
#		- Check if the required Module is installed, installs if not present
#		- Calls for Login with Microsoft using Modern-Auth
#		- Checks if "C:\Temp" exists, creates it if not found
#		- Stamps Date/Time when ran
#		- Gets Mail Flow Configuration relevant to Exclaimer
#		- Gets an Output of all Distribution Groups with "ReportToOriginatorEnabled" not set to "True"
#		- Gets "JournalingReportNdrTo" mailbox
#		- Gets all AcceptedDomains
#		- Gets Default IPAllowList settings
#		- Gets Remote Domain settings relevant to Exlcaimer (based on previous tickets)
#		- Displays a Message pop-up asking that the file/output is sent back to Support
#		- Opens Directory where the Output file was saved
#
#.INSTRUCTIONS
#	- Open PowerShell as Administrator
#	- Run: set-executionpolicy unrestricted
#	- Go to directory where the Script is saved (i.e 'cd "C:\Users\ReplaceWithUserName\Downloads"')
#	- Run the Script (i.e '.\ExchangeOnlineExclaimerCheck.ps1')
##>

#Setting variables to use later
$Path = "C:\Temp"
$LogFile = "C:\Temp\ExchangeOnlineExclaimerCheck.txt"
$TransportRuleIdentity = "Identify messages to send to Exclaimer Cloud"
$TransportRuleOOOExclaimer = "Prevent Out of Office messages being sent to Exclaimer Cloud"
$OutboundConnector = "Send to Exclaimer Cloud"
$InboundConnector = "Receive from Exclaimer Cloud*"
$DateTimeRun = Get-Date -Format "ddd dd MMMM yyyy, HH:MM 'UTC' K"

Add-Type -AssemblyName PresentationFramework


#Getting Exchange Online Module
function checkExchangeOnline-Module {
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        #[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        #[System.Windows.MessageBox]::Show('ExchangeOnlineManagement module already installed, will continue..."', 'ExchangeOnlineExclaimerCheck', 'OK', 'Information')
    } 
    else {
        [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
        [System.Windows.MessageBox]::Show('ExchangeOnlineManagement module not installed, will attempt to install it now...', 'ExchangeOnlineExclaimerCheck', 'OK', 'Information')
        Install-Module ExchangeOnlineManagement
    }
}

#Connecting to Exchange Online
function modern-auth-connect {
    Connect-ExchangeOnline
}

#Check "C:\Temp"
function path-checkTemp {
    if (Test-Path -Path $Path){
    Write-Output ("Temp folder exists") | Out-File $LogFile
    }
    Else {
    New-Item $Path -ItemType Directory
    Write-Output "Folder Created successfully" | Out-File $LogFile
    }
}

function stampIt {    
    Write-Output ("Starting all required checks: $DateTimeRun") | Out-File $LogFile -Append
}

#Collecting configuration
function get-ConfigOutput {    
    Write-Output ("`n###########---Getting Mail Flow configuration.....---###########") | Out-File $LogFile -Append
    # Getting variables count
    $CountTransportRuleOOOExclaimer = Get-TransportRule | Where {$_.Name -like $TransportRuleOOOExclaimer} | Measure-Object
    $CountTransportRuleIdentity = Get-TransportRule | Where {$_.Name -like $TransportRuleIdentity} | Measure-Object
    $CountTransportRules = Get-TransportRule | Measure-Object
    $CountExclaimerOutboundConnector = Get-OutboundConnector | Where {$_.Name -like $OutboundConnector} | Measure-Object
    $CountExclaimerInboundConnector = Get-InboundConnector | Where {$_.Name -like $InboundConnector} | Measure-Object 
    $CountOtherOutConnectors = Get-OutboundConnector | Where {$_.Name -ne $OutboundConnector} | Measure-Object

    Write-Output ("`n----------- Exclaimer Transport Rules..... -----------") | Out-File $LogFile -Append
    #Checking Transport Rule "Prevent Out of Office messages being sent to Exclaimer Cloud" 
        if ($CountTransportRuleOOOExclaimer.Count -ne "0") {
            Get-TransportRule | Where {$_.Name -like $TransportRuleOOOExclaimer} | Select-Object Name, State, Priority | Out-File $LogFile -Append
            }else{
                Write-Output ("`n##### NOTE #####`nNo 'Prevent Out of Office messages being sent to Exclaimer Cloud' Transport Rule Found`
Issues expected with Automated emails, see article below section 'The email was an out of office email':`
'https://support.exclaimer.com/hc/en-gb/articles/4406732893457'") | Out-File $LogFile -Append
        }

    #Checking Transport Rule "Identify messages to send to Exclaimer" 
        if ($CountTransportRuleIdentity.Count -ne "0") {
            Get-TransportRule | Where {$_.Name -like $TransportRuleIdentity} | Select-Object Name, State, Priority | Out-File $LogFile -Append
            Get-TransportRule | Where {$_.Name -like $TransportRuleIdentity} | Select-Object -ExpandProperty Description | Out-File $LogFile -Append
            }else{
            Write-Output ("`nNo Transport Rule 'Identify messages to send to Exclaimer' Found") | Out-File $LogFile -Append
        }

    Write-Output ("`n----------- Exclaimer Connectors..... -----------") | Out-File $LogFile -Append
    #Checking for the Exclaimer Outbound Connector
        if ($CountExclaimerOutboundConnector.Count -ne "0") {
            Get-OutboundConnector | Where {$_.Name -like $OutboundConnector} | Select-Object Identity, Enabled, SmartHosts | Out-File $LogFile -Append
            }else{
            Write-Output ("`nNo Exclaimer Outbound Connector Found Found") | Out-File $LogFile -Append
        }

    #Checking for the Exclaimer Inbound Connector
        if ($CountExclaimerInboundConnector.Count -ne "0") {
            Get-InboundConnector | Where {$_.Name -like $InboundConnector} | Select-Object Identity, Enabled, TlsSenderCertificateName | Out-File $LogFile -Append
            }else{
            Write-Output ("`nNo Exclaimer Inbound Connector Found Found") | Out-File $LogFile -Append
        }

    Write-Output ("`n----------- Other Outbound Connectors..... -----------") | Out-File $LogFile -Append
    #Checking for Other Outbound Connectors
        if ($CountOtherOutConnectors.Count -ne "0") {
            Get-OutboundConnector | Where {$_.Name -ne $OutboundConnector} | Select-Object Identity, Enabled, IsTransportRuleScoped, SmartHosts | Out-File $LogFile -Append
            }else{
            Write-Output ("`nNo Other Outbound Connectors Found") | Out-File $LogFile -Append
        }
            
    Write-Output ("`n----------- All Transport Rules..... -----------") | Out-File $LogFile -Append
    #Checking for Other Outbound Connectors
        if ($CountTransportRules.Count -gt "0") {
            Get-TransportRule | Select-Object Name, State, Priority | Out-File $LogFile -Append
            }else{
            Write-Output ("`nNo Other Transport Rules Found") | Out-File $LogFile -Append
        }
}

function get-DistributionGroups {
    $groups = Get-DistributionGroup -Filter ('ReportToOriginatorEnabled -ne $True -and IsDirSynced -eq $False')
    $dirsync = Get-DistributionGroup -Filter ('ReportToOriginatorEnabled -ne $True -and IsDirSynced -eq $true')
    $dynamicgroups = Get-DynamicDistributionGroup -Filter ('ReportToOriginatorEnabled -ne $True')
    Write-Output ("###########---Getting Distribution Groups with where 'ReportToOriginatorEnabled' is not 'TRUE'.....---###########") | Out-File $LogFile -Append

    If ($groups -ne $null) {
        Write-Output ("`nBelow are the Office 365 Distribution Groups currently set to False") | Out-File $LogFile -Append
        Write-Output $groups | Select DisplayName,ReportToOriginatorEnabled | Format-Table | Out-File $LogFile -Append
    }
    If ($dirsync -ne $null) {
        Write-Output ("`nBelow are the Office 365 Distribution groups sync'd from AD with the value of False") | Out-File $LogFile -Append
        Write-Output $dirsync | Select DisplayName,ReportToOriginatorEnabled | Format-Table | Out-File $LogFile -Append
    }
    If ($dynamicgroups -ne $null) {
        Write-Output ("`nBelow are the Office 365 Dynamic groups with the value of False") | Out-File $LogFile -Append
        Write-Output $dynamicgroups | Select DisplayName,ReportToOriginatorEnabled | Format-Table | Out-File $LogFile -Append
    }
    If ($groups -ne $null -OR $dirsync -ne $null -OR $dynamicgroups -ne $null) {
        Write-Output ("##### NOTE #####`nAny Groups that emails are sent to should have 'ReportToOriginatorEnabled' set to 'True' or`
some emails may not be delivered due to 'No Mail From' see article below:`
'https://support.exclaimer.com/hc/en-gb/articles/4406732893457'") | Out-File $LogFile -Append
    }
    Else {
        Write-Output ("`nThere are no Distribution Groups for which 'ReportToOriginatorEnabled' is not set to 'True'") | Out-File $LogFile -Append
        }
}


function get-JournalingReportNdrTo {
    Write-Output ("`n###########---Getting the mailbox configured as JournalingReportNdrTo.....---###########") | Out-File $LogFile -Append
    Get-TransportConfig | Select-Object JournalingReportNdrTo | Out-File $LogFile -Append
}

function get-AcceptedDomains {
    Write-Output ("###########---Getting all Accepted Domains.....---###########") | Out-File $LogFile -Append
    Get-AcceptedDomain | Select-Object DomainName, DomainType, Default | Out-File $LogFile -Append
}

function get-IPAllowList {    
    Write-Output ("###########---Getting IPAllowList.....---###########") | Out-File $LogFile -Append
    Get-HostedConnectionFilterPolicy -Identity Default | Select-Object Identity, IPAllowList | Out-File $LogFile -Append
}

function get-RemoteDomainOutput {    
    Write-Output ("###########---Getting Remote Domain Configuration.....---###########") | Out-File $LogFile -Append
    Get-RemoteDomain * | Select-Object Name, DomainName, CharacterSet, ContentType, TNEFEnabled | Out-File $LogFile -Append    
}

#Open Ouput directory
function open-OutputDir {
    Start "C:\Temp"
}

#Message
function user-Message {
    [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [System.Windows.MessageBox]::Show('Please provide Support the output file named "ExchangeOnlineExclaimerCheck.txt"', 'ExchangeOnlineExclaimerCheck', 'OK', 'Information')
}

checkExchangeOnline-Module
modern-auth-connect
path-checkTemp
stampIt
get-ConfigOutput
get-DistributionGroups
get-JournalingReportNdrTo
get-AcceptedDomains
get-IPAllowList
get-RemoteDomainOutput
user-Message
open-OutputDir
