<#
.SYNOPSIS
    Script to delete known registry and file system entries for the MSI installer of the Cloud Signature Update Agent
.DESCRIPTION
    This script is designed to be run as an administrator on machines where all other attempts to uninstall the Cloud Signature Update Agent
     have failed due to Microsoft Windows Registry corruption.
.NOTES
    Date: 23rd September 2021
    Update: 30/09 - Correction for Current user SID variable in Registry values list. Correction for $false result when stopping CSUA.
.PRODUCTS
    Exclaimer Cloud - Signatures for Office 365
.REQUIREMENTS
    Warning: Windows Registry modifications should always be approached with extreme care - serious problems can occur if you modify the Windows registry incorrectly!
    We strongly advise you to back up the Windows registry before any modifications are made - in doing so you will have the option to restore the backup if a problem occurs.
    For more information, see https://support.microsoft.com/en-gb/topic/how-to-back-up-and-restore-the-registry-in-windows-855140ad-e318-2a13-2829-d428a2ab0692.

    Run script as administrator
#>
#Stop the CSUA process
Write-Host "Attempting to stop the Cloud Signature Update Agent Process..." -ForegroundColor Cyan
$csua = Get-Process "Exclaimer.CloudSignatureAgent" -ErrorAction SilentlyContinue
if ($csua) {$csua | Stop-Process -Force
}
else
{
Write-Host "The Cloud Signature Update Agent is not running." -ForegroundColor Green
}
Remove-Variable csua
Write-Host "Complete." -ForegroundColor Green
#Map HKCR
Write-Host "Mapping HKCR for use with Powershell..." -ForegroundColor Cyan
New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue | out-null
Write-Host "Complete." -ForegroundColor Green
#Map Current User SID
$CUSID = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
#Declaring registry keys to be removed
$CSUARegKey = @(
'HKCR:\Installer\Features\EA514395CD9738C49BF53DFB6032CCA0'
'HKCR:\Installer\Products\EA514395CD9738C49BF53DFB6032CCA0'
'HKCR:\INSTALLER\PRODUCTS\C40DE104D46D23147B9294B69E63C1CD'
'HKCU:\SOFTWARE\Exclaimer Ltd\CloudSignatureUpdateAgent'
'HKCU:\Software\Microsoft\Installer\Features\C40DE104D46D23147B9294B69E63C1CD'
'HKCU:\Software\Microsoft\Installer\Products\C40DE104D46D23147B9294B69E63C1CD'
'HKCU:\Software\Microsoft\Installer\Products\EA514395CD9738C49BF53DFB6032CCA0'
'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Exclaimer Cloud Signature Update Agent'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\08C2F4D0EC2F8705882282685B8C2AF7'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\0FB2DCFBD81742C54A07374C9183AE11'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\25029BFCF4B3B7E56894506915234549'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\3A7E404BA977E695BA8F14D32E87071C'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\41FAE916668ADF756B9AAE81216978C0'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\4840992B9A56CE85998E24CD4395D0A8'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\746C8D349DB1F775DB5123F57C01ADE8'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\7F9C1A2B86565A95D8C51AAC57A3C839'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\81D7F39E4CB41A059B3AF8D0DE7F92CB'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\82E455B753812505AB86B8357A7CD622'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\991285CE294111D549229DC3C2146C6B'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\9ABE4934A889CB5559ED3EDE33180B30'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\A3DC272CC180D1850A5D4A996FD527F8'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\B8AB1AC9AF4D8D856A9B2D7818307D20'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\BC80D7512767B7957BD158BFE3F45A3E'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\D2FF0449664C57E409661C68FE6275AE'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\D71C0E6BFFA09715BA73B4DB1D29A81B'
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\EA514395CD9738C49BF53DFB6032CCA0\'
'HKLM:\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\INSTALLER\USERDATA\S-1-5-18\Products\C40DE104D46D23147B9294B69E63C1CD'
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\08C2F4D0EC2F8705882282685B8C2AF7"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\0FB2DCFBD81742C54A07374C9183AE11"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\25029BFCF4B3B7E56894506915234549"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\3A7E404BA977E695BA8F14D32E87071C"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\41FAE916668ADF756B9AAE81216978C0"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\4840992B9A56CE85998E24CD4395D0A8"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\746C8D349DB1F775DB5123F57C01ADE8"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\7F9C1A2B86565A95D8C51AAC57A3C839"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\81D7F39E4CB41A059B3AF8D0DE7F92CB"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\82E455B753812505AB86B8357A7CD622"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\991285CE294111D549229DC3C2146C6B"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\9ABE4934A889CB5559ED3EDE33180B30"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\A3DC272CC180D1850A5D4A996FD527F8"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\B8AB1AC9AF4D8D856A9B2D7818307D20"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\BC80D7512767B7957BD158BFE3F45A3E"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\D2FF0449664C57E409661C68FE6275AE"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\D71C0E6BFFA09715BA73B4DB1D29A81B"
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Products\EA514395CD9738C49BF53DFB6032CCA0"
"HKLM:\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\INSTALLER\USERDATA\$CUSID\Products\C40DE104D46D23147B9294B69E63C1CD"
'HKLM:\SOFTWARE\WOW6432Node\Exclaimer Ltd\CloudSignatureUpdateAgent'
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{401ED04C-D64D-4132-B729-496BE9361CDC}'
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{593415AE-79DC-4C83-B95F-D3BF0623CC0A}'
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Exclaimer Cloud Signature Update Agent'
)
#Declaring registry entries to be removed
$CSUARegEntries = @(
[pscustomobject]@{CSUARegPath='HKCR:\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5\';CSUARegName='EA514395CD9738C49BF53DFB6032CCA0'}
[pscustomobject]@{CSUARegPath='HKCU:\Software\Microsoft\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5\';CSUARegName='C40DE104D46D23147B9294B69E63C1CD'}
[pscustomobject]@{CSUARegPath='HKCU:\Software\Microsoft\Windows\CurrentVersion\Run\';CSUARegName='Cloud Signature Update Agent'}
[pscustomobject]@{CSUARegPath='HKCU:\Software\Microsoft\Windows\CurrentVersion\Run\';CSUARegName='Exclaimer Cloud Signature Update Agent'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\de\'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\es\'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\fr\'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\it\'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\nl\'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\pt\'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\"}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\"}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\de\"}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\fr\"}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\it\"}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\nl\"}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\pt\"}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Exclaimer\"}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:APPDATA\Microsoft\Installer\{401ED04C-D64D-4132-B729-496BE9361CDC}\"}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Windows\Installer\{593415AE-79DC-4C83-B95F-D3BF0623CC0A}\'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5\';CSUARegName='C40DE104D46D23147B9294B69E63C1CD'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5\';CSUARegName='EA514395CD9738C49BF53DFB6032CCA0'}
[pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run\';CSUARegName='Cloud Signature Update Agent'}
)
#Declaring file system entries to be removed
$CSUAFSEntries = @(
'C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\'
"$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent"
"$env:LOCALAPPDATA\Exclaimer"
"$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Exclaimer Ltd"
"$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Exclaimer"
'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Exclaimer'
)
Write-Host "Removing Registry entries for the Cloud Signature Update Agent..." -ForegroundColor Cyan
#Remove Registry Keys
$CSUARegKey | ForEach-Object {Remove-Item $PSItem -Recurse -ErrorAction SilentlyContinue}
$CSUARegKey = $null
#Remove Registry Entries
$CSUARegEntries | ForEach-Object {Remove-ItemProperty -Path $PSItem.CSUARegPath -name $PSItem.CSUARegName -ErrorAction SilentlyContinue}
$CSUARegEntries = $null
Remove-Variable CUSID
Write-Host "Complete." -ForegroundColor Green
#Remove file system entries
Write-Host "Removing file system entries for the Cloud Signature Update Agent..." -ForegroundColor Cyan
$CSUAFSEntries | ForEach-Object {Remove-Item -Path $PSItem -Recurse -ErrorAction SilentlyContinue}
$CSUAFSEntries = $null
Write-Host "Complete." -ForegroundColor Green
#Remove HKCR mapping
Write-Host "Removing mapping for HKCR..." -ForegroundColor Cyan
Remove-PSDrive -Name HKCR  -ErrorAction SilentlyContinue
Write-Host "Complete." -ForegroundColor Green
#Output errors to log file
Write-Host "Writing log file..." -ForegroundColor Cyan
$Error | Out-File -FilePath $env:LOCALAPPDATA\Temp\CSUACleanupScript.log
Write-Host "Script Complete.. Log saved to $env:LOCALAPPDATA\Temp\CSUACleanupScript.log" -ForegroundColor Green