## Script to enable mailboxes for domain
## Only needed where you have a lot of AD accounts without Mailboxes

## Load modules for AD commands and Exchange Commands
Import-Module ActiveDirectory
add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010

## Query for AD accounts
# Below commands queries the top level OU for the company you created and then outputs the details into a CSV
$toplevelOU = Read-Host ("What is the company OU?")
Get-ADUser -Filter * -SearchBase "OU=$toplevelOU,DC=$env:USERDOMAIN,DC=local" | `
Select Name,SamAccountName,UserPrincipalName | `
Export-Csv -Path $env:tmp\output.csv -NoTypeInformation


## Enable Mailbox for all users in the CSV
Import-Csv $env:tmp\output.csv | ForEach-Object {Enable-Mailbox -Identity $_.UserPrincipalName -Alias $_.SamAccountName}

$delete = Read-Host ("Do you want to remove the CSV file of Users? Y/N")

If ($delete -eq "y") {
    Remove-Item $env:tmp\output.csv }
#Else { $env:tmp }

Remove-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
Remove-Module ActiveDirectory