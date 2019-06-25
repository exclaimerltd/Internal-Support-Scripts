# Send Mail Script
#{
<#
.SYNOPSIS
    Batch sends emails through the smtp server of your choice
.DESCRIPTION
    This script takes the users input for a few variables and then loops through the Send-MailMessage command
    until hitting the user defined number of emails that need to be sent.
.NOTES
    Email: support@exclaimer.com
    Date: 24th January, 2017
.USAGE
    [item] This signifies the default options for each question
#>
 
Function Mail {
    $body = "This was sent at $date"
    $subject = "Test $i email sent at $date"
    Send-MailMessage -From $from -To $to -Subject $subject -Body $body -BodyAsHtml -SmtpServer $server
}
 
# Finds logged in users email address
$email = $searcher = [adsisearcher]"(samaccountname=$env:USERNAME)"
$email = $searcher.FindOne().Properties.mail
 
# Default for $to and $from will be $email if none specified
$from = Read-Host ("Who will the messages be sent from? [$email]")
$to = Read-Host ("Who will the messages be sent to? [$email]")
    If (!$to) {
        $to = $email
    }
    If (!$from) {
        $from = $email
    }
 
# Default $server to localhost if none specified
$server = Read-Host ("Which server will the messages be sent from? [localhost]")
    If (!$server) {
        $server = "localhost"
    }
 
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