<#
.SYNOPSIS
    Checks and outputs the configuration of Transport Rule and Connectors that may affected routing of emails to Exclaimer.

.DESCRIPTION
    It will first prompt to login with Microsoft, then it will check and outputs the configuration of Transport Rule and Connectors that may affected routing of emails to Exclaimer.
    Output filename is "ExchangeOnlineExclaimerCheck.txt"

.NOTES
    Email: helpdesk@exclaimer.com
    Date: 27th June 2021

.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365

.REQUIREMENTS
    - Global Administrator Account
    - Requires path "C:\Temp"
    - ExchangeOnlineManagement - https://www.powershellgallery.com/packages/ExchangeOnlineManagement/0.4578.0

.VERSION
		1.0.0
			- Check if the required Module is installed, installs it if not present
			- Calls for Login with Microsoft using Modern-Auth
			- Checks if "C:\Temp" exists, creates it if not found
			- Stamps Date/Time when ran
			- Gets Mail Flow Configuration relevant to Exclaimer
			- Gets an Output of all Distribution Groups with "ReportToOriginatorEnabled" not set to "True"
			- Gets "JournalingReportNdrTo" mailbox
			- Gets all AcceptedDomains
			- Gets Default IPAllowList settings
			- Gets Remote Domain settings relevant to Exlcaimer (based on previous tickets)
			- Displays a Message pop-up asking that the file/outpt is sent back to Support
			- Opens Directory where the Output file was saved
			
.INSTRUCTIONS
		- Open PowerShell as Administrator
		- Run: set-executionpolicy unrestricted
		- Go to directory where the Scrit is saved (i.e 'cd "C:\Users\ReplaceWithUserNAme\Downloads"')
		- Run the Script (i.e '.\ExchangeOnlineExclaimerCheck.ps1')
#>

# Setting variables to use later
$Path = "C:\Temp"
$LogFile = "C:\Temp\ExchangeOnlineExclaimerCheck.txt"
$TransportRuleIdentity = "Identify messages to send to Exclaimer Cloud"
$OutboundConnector = "Send to Exclaimer Cloud"
$InboundConnector = "Receive from Exclaimer Cloud*"
$DateTimeRun = Get-Date -DisplayHint Date

Add-Type -AssemblyName PresentationFramework


#Getting Exchange Online Module
function checkExchnageOnline-Module {
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
Write-Output ("Starting all required checks $DateTimeRun") | Out-File $LogFile -Append
}

# Collecting configuration
function get-ConfigOutput {    
    Write-Output ("`n########### Getting Mail Flow configuration..... ###########") | Out-File $LogFile -Append
    Get-TransportRule -Identity $TransportRuleIdentity | Select-Object Identity, State| Out-File $LogFile -Append
    Get-TransportRule -Identity $TransportRuleIdentity | Select-Object -ExpandProperty Description | Out-File $LogFile -Append
    Get-OutboundConnector -Identity $OutboundConnector | Select-Object Identity, Enabled, SmartHosts | Out-File $LogFile -Append
    Get-InboundConnector -Identity $InboundConnector | Select-Object Identity, Enabled, TlsSenderCertificateName  | Out-File $LogFile -Append
    Get-OutboundConnector | Where Identity -ne $OutboundConnector | Select-Object Identity, Enabled, IsTransportRuleScoped, SmartHosts | Out-File $LogFile -Append
}

function get-DistributionGroups {
    $groups = Get-DistributionGroup -Filter ('ReportToOriginatorEnabled -ne $True -and IsDirSynced -eq $False')
    $dirsync = Get-DistributionGroup -Filter ('ReportToOriginatorEnabled -ne $True -and IsDirSynced -eq $true')
    $dynamicgroups = Get-DynamicDistributionGroup -Filter ('ReportToOriginatorEnabled -ne $True')
    Write-Output ("########### Getting any Distribution Groups that could be affected by 'No Mail From'..... ########### `nAny Groups that emails are sent to should have 'ReportToOriginatorEnabled' set to 'True' or some may not be delivered due to 'No Mail From'") | Out-File $LogFile -Append

    If ($groups -ne $null) {
        Write-Output ("Below are the Office 365 Distribution Groups currently set to False") | Out-File $LogFile -Append
        Write-Output $groups | Select DisplayName,ReportToOriginatorEnabled | Format-Table | Out-File $LogFile -Append
    }
    If ($dirsync -ne $null) {
        Write-Output ("Below are the Office 365 Distribution groups sync'd from AD with the value of False") | Out-File $LogFile -Append
        Write-Output $dirsync | Select DisplayName,ReportToOriginatorEnabled | Format-Table | Out-File $LogFile -Append
    }
    If ($dynamicgroups -ne $null) {
        Write-Output ("Below are the Office 365 Dynamic groups with the value of False") | Out-File $LogFile -Append
        Write-Output $dynamicgroups | Select DisplayName,ReportToOriginatorEnabled | Format-Table | Out-File $LogFile -Append
    }
    If ($groups -ne $null -OR $dirsync -ne $null -OR $dynamicgroups -ne $null) {
        Write-Output ("See article'https://support.exclaimer.com/hc/en-gb/articles/4406732893457'") | Out-File $LogFile -Append
    }
    Else {
        Write-Output ("There are no Distribution Groups with 'ReportToOriginatorEnabled' set to 'False'") | Out-File $LogFile -Append
        }
}


function get-JournalingReportNdrTo {
    Write-Output ("`n########### Getting the mailbox configured as JournalingReportNdrTo..... ###########") | Out-File $LogFile -Append
    Get-TransportConfig | Select-Object JournalingReportNdrTo | Out-File $LogFile -Append
}

function get-AcceptedDomains {
    Write-Output ("########### Getting all Accepted Domains..... ###########") | Out-File $LogFile -Append
    Get-AcceptedDomain | Select-Object DomainName, DomainType, Default | Out-File $LogFile -Append
}

function get-IPAllowList {    
    Write-Output ("########### Getting IPAllowList..... ###########") | Out-File $LogFile -Append
    Get-HostedConnectionFilterPolicy -Identity Default | Select-Object Identity, IPAllowList | Out-File $LogFile -Append
}

function get-RemoteDomainOutput {    
    Write-Output ("########### Getting Remote Domain Configuration..... ###########") | Out-File $LogFile -Append
    Get-RemoteDomain * | Select-Object Name, CharacterSet, ContentType, TNEFEnabled | Out-File $LogFile -Append    
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

checkExchnageOnline-Module
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
