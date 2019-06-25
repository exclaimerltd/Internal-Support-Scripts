# Empty local signatures
#{
<#
.SYNOPSIS
    Removes any deployed signatures from Appdata folders
.USAGE
    Can be executed manually from a user machine or via a logon script
    This script will delete any existing signatures from a customers microsoft\signatures folder
.NOTES
    Email: support@exclaimer.com
    Date: 30th August 2017
.REQUIREMENTS
    None
.VERSION
    1.0 - Removes the signatures from the English and Dutch signature folders
#>
 
$user = $env:APPDATA
 
Function english {
    $engsig = "$user\Microsoft\Signatures"
    $engexist = Test-Path -path $engsig
 
    If ($engexist -eq "True") {
        Remove-Item $engsig -Recurse -Force
        New-Item -Path $engsig -ItemType Directory
    }
}
 
Function dutch {
    $dutsig = "$user\Microsoft\Handtekeningen"
    $dutchexist = Test-Path -Path $dutsig
 
    If ($dutchexist -eq "True") {
        Remove-Item $dutsig -Recurse -Force
        New-Item -Path $dutsig -ItemType Directory
    }
}
 
english
dutch