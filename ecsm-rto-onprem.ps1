# 
<#
.SYNOPSIS
    Sets the ReportToOriginator field to $true for on premise groups
.DESCRIPTION
    If the field is set to false, it can cause a number of issues when emailing groups due to the emails sent
    not containing any sender envelope data or a return-path.  This causes knock on affects with spam and signature
    application.
 
    The script is a quicker corrective measure to the issue that setting it manually and can be used regularly to
    ensure all groups are set correctly.
 
    This script requires the ADSchema to have been updated for Exchange.  Without this, the field does not exist
    on premise and when the group syncs it will always be $false.
.NOTES
    Email: helpdesk@exclaimer.com
    Date: 13th May 2017
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    - ADSchema update from Exchange has been completed - https://www.petri.com/how-to-install-exchange-server-2013
    - Script should be executed on the Domain Controller
.VERSION
    1.0 - Sets the ReportToOriginator field for groups on premise where the field isnt set
    1.1 - Added a line to set for ReportToOwner also based on recent tickets
#>
 
function ad_connect {
    Import-Module ActiveDirectory
}
 
function ad_gatherchange {
    $group = Get-ADGroup -Filter ('ReportToOriginator -eq $False -or ReportToOriginator -notlike "*"')
 
    If ($group -ne $null) {
        Write-Host ("Below are the on premise groups with ReportToOriginator set to $false or nothing") -ForegroundColor Green
        Write-Host ("###############")
        Write-Output $group | Select -Property Name
        }
    Else {
        Write-Host ("All groups are set to $true already") -ForegroundColor Green
        Exit
    }

    $change = Read-Host ("Do you want to change these groups to True? y/N")
     
    If (!($change)) {
        Write-Host ("No selection made, this script will now exit")
        start-sleep -Seconds 5
        exit
    }
    Else {
        If ($change -eq "y") {
            $group | Set-ADGroup -Replace @{ReportToOriginator=$true}
            $group | Set-ADGroup -Replace @{ReportToOwner=$false}
            Write-Host ("Group Changes Complete!") -ForegroundColor Green
            Write-Host ("Please synchronise your On Premise AD with Office 365") -ForegroundColor Green
            exit
        }
    }
}
 
ad_connect
ad_gatherchange