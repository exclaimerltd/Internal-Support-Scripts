
# 
<#
.SYNOPSIS
    Execute the Exclaimer Email Archiver BIN2PST utility.
.DESCRIPTION
    This script takes a directory containing a number of .BIN files used by
    Exclaimers Email Archiver application and iterates over them running the
    BIN2PST utility to export all emails to PST files.
.PARAMETER Store
    The directory or directories containing the archive store(s). Folders are recursed. 
.EXAMPLE
    Run-Bin2Pst -Store C:\Store
 
    This will find all .BIN files in C:\Store and run BIN2PST with the
    resulting PST files being placed in the same directory.
.EXAMPLE
    Get-ChildItem -Directory -Path C:\Store | Run-Bin2Pst 
 
    This will find all .BIN files in C:\Store and run BIN2PST with the
    resulting PST files being placed in the same directory.
.EXAMPLE
    ("C:\Store1", "C:\Store2") | Run-Bin2Pst 
 
    This will find all .BIN files in C:\Store1 and C:\Store2 and run BIN2PST with the
    resulting PST files being placed in the same directory.
.NOTES
    Authored By: Simon Buckner
    Email: simon@onebyte.net
    Date: 13th July, 2016
.PRODUCTS
    Mail Archiver
#>
[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
    [ValidateLength(3,256)]
    [Alias("BIN")]
    [String[]]$Store
)

BEGIN
{
    # Run one-time set-up tasks here, like defining variables, etc.
    Set-StrictMode -Version Latest
    Write-Verbose -Message "Run-Bin2Pst started."
}
PROCESS
{
    # The process block can be executed multiple times as objects are passed through the pipeline into it.
    ForEach($currentStore In $Store)
    {
        If($PSCmdlet.ShouldProcess($currentStore))
        {
            Write-Verbose -Message "Processing Store: $currentStore"
            $binFiles = Get-ChildItem -Path $currentStore -Recurse -Filter "*.BIN"
            Write-Verbose -Message "BIN files found: $($binFiles.Count)"
            ForEach($binFile in $binFiles)
            {   
                $bin = $binFile.FullName
                $pst = "$($binFile.DirectoryName)\$($binFile.BaseName).pst"
                If(Test-Path $pst) {
                    Write-Verbose "PST Exists, skipping File: $bin"
                    Continue    
                }
                Write-Verbose "Processing File: $bin"
                Write-Verbose " -> Output File: $pst"

                $allArgs = @( "$bin", "$pst", "/passive")

                & 'C:\Program Files (x86)\Exclaimer Ltd\Mail Archiver\BIN2PST.exe' $allArgs  

            }
        }       
    }
}
END
{       
    # Finally, run one-time tear-down tasks here.
    Write-Verbose -Message "Running End block."
}

