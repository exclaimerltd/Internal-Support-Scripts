# Extract templates from backup
#{
<#
.SYNOPSIS
    Extracts all template files from a customers .econfig file
.USAGE
    Used in support tickets where the main configuration has been corrupted and you need to obtain the signatures from a backup
.NOTES
    Email: support@exclaimer.com
    Date: 14th June 2018
.REQUIREMENTS
    PowerShell V5
.VERSION
    1.0 - Unzips config file and creates a zip
    1.1 - Fixed a bug where multiple zips in the folder will cause it to fail
#>

# Get current path
$currentpath = (Get-Item -Path ".\").FullName
$output = "\backup-output"
$done = "\completed-export"
$unzippath = Join-Path $currentpath $output
$donepath = Join-Path $currentpath $done
New-Item -Path $donepath -ItemType Directory

# Gather details
$backup_files = Get-ChildItem | Where-Object {$_.Extension -eq ".econfig"}

# Copy backup file
function backup-copy {
    ForEach ($file in $backup_files) {
        $filenew = $file.Name + ".bak"
        Copy-Item $file $filenew
    }
}

# Rename the original backup file
function rename-backup {
    ForEach ($file in $backup_files) {
        $filenew = $file.Name + ".zip"
        Rename-Item $file $filenew
    }
}

# Unzip the backup file
function unzbackup {
    $filein = (Get-ChildItem -Path *.zip).fullname
    ForEach ($file in $filein) {    
        expand-archive $file -DestinationPath $unzippath
    }
}

# Rename all files inside extracted folders
function rename-output {
    $completedbackups = Get-ChildItem -Path $unzippath -Recurse | Where-Object {!($_.Extension)} | Where-Object {!($_.PSIsContainer) -and ($Exclude -notcontains $_.Name)}
    ForEach ($file in $completedbackups) {
        $filenew = $file.FullName + ".zip"
        Rename-Item $file.FullName $filenew
    }  
}

# Move outputed zip to new folder
function move-zip {
    $unzipfiles = Get-ChildItem -Path $unzippath -Recurse | Where-Object {$_.Extension -eq ".zip"}
    foreach ($file in $unzipfiles) {
        Move-Item -Path $file.FullName -Destination $donepath
    }   
}

backup-copy
rename-backup
unzbackup
rename-output
move-zip

$destination = $currentpath + "\complete-backup.zip"
Compress-Archive -LiteralPath $donepath -DestinationPath $destination

# Clean up
Remove-Item $unzippath -Recurse -Force
Remove-Item $donepath -Recurse -Force
Remove-Item (Get-ChildItem | Where-Object {$_.Extension -eq ".bak"}) -Recurse -Force