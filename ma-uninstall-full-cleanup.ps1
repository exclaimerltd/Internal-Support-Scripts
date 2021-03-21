# 
<#
.SYNOPSIS
    Cleans up associated setup from Mail Archiver
.USAGE
    Designed to clean up a lab environment of the installation post usage without requiring the lab to be rebuilt
.NOTES
    Email: helpdesk@exclaimer.com
    Date: 10th August 2018
.PRODUCTS
    Mail Archiver
.REQUIREMENTS
    - Domain Administrator
    - Administrator rights to the Archiver machine
    - Access to the archiver machine
.HISTORY
    1.0 - Removes the journal rule, the default EMA account and the default Journal Account
#>

function remove-rule {
    # Removes to Exclaimer Journal Rule
    $JRules = Get-JournalRule | Where-object {$_.Identity -like "Created by Exclaimer Mail Archiver*"}
    {ForEach ($i in $JRules)
        {
            $DelJRi = Read-Host -prompt "Would you like to delete the rule: $i ? (y/n)"
        }
        If ($DelJRi -eq "y") {Remove-JournalRule -identity "$i"}
    }
}

function remove-ema {
    # Removes the Exclaimer EMA account
    $emaaccount = Read-Host ("Please enter the name of your Exchange Mailbox Access account")

    Write-Host "#####"
    Write-Host ""
    Write-Host "The $emaaccount account is about to be removed"
    Write-Host ""
    $ans = Read-Host "Are you sure you want to proceed? Y/n"

    If ($ans = "y") {
        Remove-Mailbox -Identity $emaaccount
    }
}

function remove-journal {
    # Removes the Exclaimer Mail Archiver Journal account
    $journalaccount = Read-Host ("Please enter the name of your Exchange Journal Access account")

    Write-Host "#####"
    Write-Host ""
    Write-Host "The $journalaccount account is about to be removed"
    Write-Host ""
    $ans = Read-Host "Are you sure you want to proceed? Y/n"

    If ($ans = "y") {
        Remove-Mailbox -Identity $journalaccount
    }
}

function uninstall {
    # Uninstalls the web search and console using msiexec
    msiexec.exe /x "C:\ProgramData\Exclaimer Ltd\Cache\Mail Archiver Search Website Install.msi" /l*v $env:TEMP\web-search.log /qn
    msiexec.exe /x "C:\ProgramData\Exclaimer Ltd\Cache\Mail Archiver Install.msi" /l*v $env:TEMP\console.log /qn
}

function remove-files {
    # Cleans up the contents of the cache folder
    Remove-Item -Path "C:\ProgramData\Exclaimer Ltd\Cache\*Archiver*" -Force
}

# Connects to Exchange 
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010;
Import-Module ActiveDirectory

remove-rule
remove-ema
remove-journal
uninstall

Start-Sleep -Seconds 60

remove-files

