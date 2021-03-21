# 
<#
.SYNOPSIS
    Sets the RequireSenderAuthenticationEnabled field to $false for all Office 365 Groups
.DESCRIPTION
    If the field is set to true it can cause the issue outlined here 
    (https://cloudsupport.exclaimer.com/hc/en-us/articles/207803529-Messages-sent-to-SharePoint-or-Office-365-groups-generate-an-NDR)
 
    This script aims to achieve the steps in that guide en masse
.NOTES
    Email: support@exclaimer.com
    Date: 30th August 2017
.PRODUCTS
    Exclaimer Cloud - Signature for Office 365
.REQUIREMENTS
    - Global Administrator Account     
#>
 
function o365_connect {
    # below connects to Office 365
    $credential = Get-Credential
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -Credential $credential -ConnectionUri https://ps.outlook.com/powershell -Authentication Basic -AllowRedirection
    Import-PSSession $session
}
 
function o365_gather {
    # Check for O365 group
    $groups = Get-UnifiedGroup -filter * | ? {$_.RequireSenderAuthenticationEnabled -eq $true}
 
    If ($Ogroups -ne $null) {
    Write-Host ("Below are the Office 365 groups with the RequireSenderAuthenticationEnabled field set to true") -ForegroundColor Green
    Write-Host ("####") -ForegroundColor Green
    Write-Host ("")
    Write-Output $OGroups | Select Name,RequireSenderAuthenticationEnable | Format-Table
    Write-Host ("")
    Write-Host ("####") -ForegroundColor Green
    }
    Else {
        Write-Host ("There are no groups with RequireSenderAuthenticationEnabled set to true") -ForegroundColor Green
        Write-Host ("The script will now end") -ForegroundColor Green
        Return
        }
}
 
function o365_change {
    $change = Read-Host ("Do you want to change these groups to false? Y/n")
 
    If ($change -eq "y") {
        $Ogroups | Set-UnifiedGroup -RequireSenderAuthenticationEnabled $false
        Write-Host ("Group Changes Complete!")
    }
}
 
o365_connect
o365_gather
o365_change
 
Remove-PSSession $Session
Write-Host ("Script Complete!") -ForegroundColor Green