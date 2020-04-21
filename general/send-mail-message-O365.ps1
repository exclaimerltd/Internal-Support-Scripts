# Send Mail Script
#{
<#
.SYNOPSIS
    Batch sends emails for Office 365 accounts
.DESCRIPTION
    This script takes the users input for a few variables and then loops through
the Send-MailMessage command
    until hitting the user defined number of emails that need to be sent.
.NOTES
    Email: support@exclaimer.com
    Date: 21st April, 2020

#>
Function Mail {
    $body = "This was sent at $date"
    $subject = "Test $i email sent at $date"
    Send-MailMessage -Credential $smtpcred -from $from -to $to -Subject $subject -body $body -SmtpServer "smtp.office365.com" -port "587" -BodyAsHtml -UseSsl
}
$smtpcred = (Get-Credential)

$from = Read-Host ("Who will the messages be sent from?")
$to = Read-Host ("Who will the messages be sent to?")

# Default to 50 Messages Sent
$messagecount = Read-Host ("How many messages do you want to send? [50]")
    If (!$messagecount) {
        $messagecount = "50"
    }
ForEach ($i in 1..$messagecount) {
    # Gets Date for loop
    $date = Get-Date
    $day = $date.Day
    $month = $date.Month
    $year = $date.Year
    Mail
    # Below is a 5 second pause between emails
    Start-Sleep -s 5
}
