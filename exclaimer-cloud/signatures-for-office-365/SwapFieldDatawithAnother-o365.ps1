# 
<#
.SYNOPSIS
    Copy an O365 attribute to another field
.DESCRIPTION
    This script copies the data from a user defined to another user defined field
.NOTES
    Email: support@exclaimer.com
    Date: 31st August 2016
.USAGE
    As the attributes cannot be specified using variables you will need to do the following to tailor the script to your use
    - Replace all instances of <field1> with the field you have data you want copying
    - Replace all instances of <field2> with the field you wish to copy to
    - You can further filter the Get-User command that obtains the user information with the -filter parameter.  Examples of this below
        - Get-User -Filter "HomePhone -like '*'"
.PRODUCTS
    Exclaimer Cloud - Sigonatures for Office 365
.REQUIREMENTS
    - Global Administrator Account
#>
 
$users = Get-User | Select UserPrincipalName,"<field1>","<field2>" | Export-CSV -Path $env:APPDATA\Users.csv -NoTypeInformation
$out = Get-User | Select UserPrincipalName,"<field1>","<field2>"
 
Write-Host ("These are the users that are going to be altered. There is a CSV of this outputted to %appdata%\Users.csv")
Write-Output $Out
Write-host ("")
$confirm = Read-Host ("Are you sure you want to make these changes? y/N")
 
if ($confirm -eq "y" -or "yes") {
    Import-CSV $env:appdata\users.csv | foreach { Set-User $_.UserPrincipalName -"<field2>" $_."<field1>" } }