#
<#
.SYNOPSIS
    Rebuild.exe loop script
.DESCRIPTION
    This script loops the Rebuild tool for repairing BIN files based on a store variable
.NOTES
    Email: helpdesk@exclaimer.com
    Date: 22nd January 2019
.PRODUCTS
    Mail Archiver
.REQUIREMENTS
    A copy of the rebuild tool available here - https://exclaimerproducts.blob.core.windows.net/exclaimer/Tools/Mail%20Archiver/rebuild-tool.zip
    Visual C++ Redistributable for Visual Studio 2012 Update 4 available here - https://www.microsoft.com/en-gb/download/details.aspx?id=30679
.HISTORY
    1.0 - Script will .old original Bins and create repaired bins in the original bins folder
#>

$store = "REPLACE WITH PATH TO BINS"

Write-Verbose -Message "Processing Store: $store"
$binFiles = Get-ChildItem -Path $Store -Recurse -Filter "*.BIN"
Write-Verbose -Message "BIN files found: $($binFiles.Count)"

ForEach($binFile in $binFiles)
{
  $bin = $binFile.FullName
  $out = "$($binFile.DirectoryName)\$($binFile.BaseName)" + "1.bin"
  If(Test-Path $out)
  {
      Write-Verbose "Out Exists, skipping File: $bin"}
      else {
        Write-Verbose "Processing File: $bin"
        Write-Verbose " -> Output File: $out"

        & 'C:\BinRepair\Rebuild.exe' "$bin" "$out"
        
        Rename-Item $bin "$($binFile.Fullname).old"
    }
}