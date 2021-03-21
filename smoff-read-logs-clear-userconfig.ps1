# 
<#
.SYNOPSIS
    Script searches logs and removes user.config based on errors
.DESCRIPTION
    This script searches the Windows Application logs for a specific search
    term relating to an error for the Signature Manager Office 365 Edition agent
    and then attempts to clean up the agent by automating the initial troubleshooting
    steps for this software.

    If the error is found, the software will remove the user credentials file that the software
    uses and restart the agent.

    A log file will be created in "C:\Windows\Temp\", location can be changes by editiing variable "$SaveToLog"
.NOTES
    Email: nuno.chaves@exclaimer.com
    Date: 10th September 2018
.PRODUCT
    Signature Manager Office 365 Edition
.VERSION
    1.0 - Deletes the user.config file from the users local appdata folders after querying Windows Logs  
.README
    - Prior to use, confirm below variable values meet your requirements
    - Recommend deploying via GPO, however, comment out line 53 "Start-Sleep" if you wish to run manually

#>


#Will be searched for in the Events details
$SearchTerm = 'due to failure to connect to Office 365'

#How far back to search Windows Logs
$HoursToSearch = 3

#Name the Event Source to
$Source = 'Application'

#Event ID to search for
$EventID = 0

#Name of the processed to be restarted
$ProcessName = 'Outlook Signature Update Agent'

#File to log results to
$SaveLogTo = 'C:\Windows\Temp\ExclaimerScriptLog.txt'

#FilePath
$ClickOnceConfig = "$Env:LOCALAPPDATA\local\Apps\2.0\Data\*\*\*\Data\*.*.*.*"
$MSIConfig = "$Env:LOCALAPPDATA\local\Exclaimer_Ltd\*\*.*.*.*"

#Resetting the counter
$Counter = 0
Start-Sleep -s 180
$(Get-Date -Format G) + " Searching Logs" | Out-File -filepath $SaveLogTo -Append
$Events = @()
$Events += Get-EventLog -LogName $Source -InstanceId $EventID -After (Get-Date).AddHours(-"$HoursToSearch") | select message -ExpandProperty Message
$(Get-Date -Format G) + " Found" + $Events.Count + "logs from the source" + "'$Source'" + "and Event ID" + "'$EventID'" | Out-File -filepath $SaveLogTo -Append

#If no Events found it will Exit, no changes will be made
If (!($Events)) {
    $(Get-Date -Format G) + " No Logs Found... Exiting" | Out-File -filepath $SaveLogTo -Append
    Exit
    }
    Else {
        ForEach ($Event in $Events | Where {$_.Message -like "*$SearchTerm*"})
        {
        $Counter++
        }
    }

#If Counter is 0 it will Exit without making changes, if larger than 0 stops the process, deletes files and starts the process
If ($Counter -eq 0) {
    $(Get-Date -Format G) + " Logs found do not contain the search term" + "'$SearchTerm'" | Out-File -filepath $SaveLogTo -Append
    $(Get-Date -Format G) + " Exiting..." | Out-File -filepath $SaveLogTo -Append
    Exit
    }
    Else {
        $(Get-Date -Format G) + " A total of" + $Counter + "logs contain the sequence" + "'$SearchTerm'" | Out-File -filepath $SaveLogTo -Append
        $(Get-Date -Format G) + " Stopping $ProcessName" | Out-File -filepath $SaveLogTo -Append
        $ProcessPath = Get-Process -Name $ProcessName | Select-Object -ExpandProperty Path
        Get-Process -Name $ProcessName | Stop-Process
        Start-Sleep -s 10
        
        $(Get-Date -Format G) + ' Deleting "user.config" file' | Out-File -filepath $SaveLogTo -Append
        Remove-Item $ClickOnceConfig\user.config -Recurse -Force
        Remove-Item $MSIConfig\user.config -Recurse -Force
        $(Get-Date -Format G) + " Starting $ProcessName" | Out-File -filepath $SaveLogTo -Append
        Start-Process -FilePath $ProcessPath
    }


