#
<#
.SYNOPSIS
    Copy the values of one attribute to another
.DESCRIPTION
    This script copies the data from a user defined to another user defined field
.NOTES
    Email: helpdesk@exclaimer.com
    Date: 31st August 2016
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    - Domain Administrator Account
    - Domain controller access
#>

Import-Module ActiveDirectory
 
    Clear-Host
    $originalfield = Read-Host("What field would you copy data from from?")
    $changefield = Read-Host("What field would you like to copy data too?")
    $users = Get-ADUser -LDAPFilter "($originalfield=*)" -Properties $originalfield, $changefield
 
    write-host ("The below list of users are about to be altered")
    write-host ($users)
    $confirm = read-host ("Are you sure? y/N")
 
# Changes the value of each users $changefield in the $users variable to $originalfield
if ($confirm -eq "y") {
    foreach ($user in $users) {
        Select-Object * -First 5
        ForEach-Object {Set-ADObject -Identity $user.DistinguishedName -Replace @{$changefield=$($user.$originalfield)}}
        }
}