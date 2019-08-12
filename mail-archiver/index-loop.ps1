#
<#
.SYNOPSIS
    Script to loop the reindex tool on all stores in Archiver
.DESCRIPTION
    This script will obtain a list of stores from the registry and 
    the reindex tool through this array
.NOTES
    Authored By: David Milward
    Email: support@exclaimer.com
    Date: 12th August 2019
.PRODUCTS
    Mail Archiver 4.0
.REQUIREMENTS
    - Latest Mail Archiver install
    - Reindex tool
.HISTORY
    - Initial resync done.  Message output to console for each completed store.
#>

$path = 'HKLM:\SOFTWARE\Exclaimer Ltd\Mail Archiver 1.0\Configuration'
$stores = Get-ChildItem $path | Get-ItemProperty | Select-Object -ExpandProperty PSChildName

ForEach ($store in $stores) {
    $logs = "C:\ProgramData\Exclaimer Ltd\reindex\logs.txt"
    C:\ProgramData\"Exclaimer Ltd"\reindex\reindex.exe $store > $logs
    write-output "This store has been completed $store"
    rm $logs
}