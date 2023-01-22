#Declaring Signatures folder location
$SignaturesFolder = "$env:APPDATA\Microsoft\Signatures"
#Declaring Signature backup folder location
$SignaturesBackupFolder = "$env:APPDATA\Microsoft\SignaturesBackup"
#Backup Existing signature files
If (! (Test-Path -path "$SignaturesBackupFolder")){
Copy-Item -Path "$SignaturesFolder" -Destination "$SignaturesBackupFolder" -Recurse
}
#Delete original signature files
Get-ChildItem -Path "$SignaturesFolder" -Include *.* -Recurse | foreach { $_.Delete()}
