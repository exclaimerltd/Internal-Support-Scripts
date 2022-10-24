<#
.SYNOPSIS
    Script to delete known registry and file system entries for the MSI installer of the Cloud Signature Update Agent
.DESCRIPTION
    This script is designed to be run as an administrator on machines where all other attempts to uninstall the Cloud Signature Update Agent
     have failed due to Microsoft Windows Registry corruption.
.NOTES
    Date: 23rd September 2021
    Update: 30/09/2021 - Correction for Current user SID variable in Registry values list. Correction for $false result when stopping CSUA.
    Update 22/11/2021  - Correction for missing product codes for some versions.
                       - Added search for product codes for future versions.
                       - Added registry locations for Intune Managed deployments.
    Update 05/01/2022  - Additional correction for missing product codes for one version.
                       - Added correction for future version checker.
    Update 11/04/2022  - Added additional registry locations.
    Update 17/10/2022  - Added additional registry locations.
                       - Added search for group policy deployments

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
#Declaring known CSUA Product Codes
    $CSUAProdCode = @(
    '{BC9DA548-0FBF-4720-9C7D-CC92C91B4623}'
    '{C03C0E99-F22B-4359-9F9D-AA97123E937C}'
    '{2521B871-AAD4-4B29-8CC6-63CF2F69AD8E}'
    '{AF8EA4D8-13B4-4F44-AE0F-B0272BC398AA}'
    '{D11487D1-56A6-4315-91C3-D6800CA8FBCE}'
    '{8E4424AD-F11F-4ABF-B852-FF45CA0F7A27}'
    '{381D92C1-4760-4E5C-B6C0-E8EA055DE731}'
    '{530A9AD4-B33D-4B69-B207-03979EAB246F}'
    '{0BD9CE27-5655-43D8-9C43-BFA9026B4C67}'
    '{34F11100-BE83-445B-BFAE-2A277FA3959E}'
    '{CFFA0CC0-3CD7-4631-839D-3847C67A7412}'
    '{67F0AA08-C61D-45CB-8E5C-E9BFB4EFC974}'
    '{9FBE9BB5-8B18-4862-8954-B62F2F1457E3}'
    '{908F8BDC-E36F-4A85-B360-AD8FF373A805}'
    '{B8208FDF-173E-438E-BE19-3BF605C35047}'
    '{39B3BA76-9F07-45DF-9BA2-4272ABBBBA51}'
    '{DF7F5D27-5E41-49A2-8813-4EF926225086}'
    '{D3CD1B44-71F4-46C8-8B26-022791DC6067}'
    '{67EC6A29-95ED-40FC-A8C5-B89D47167061}'
    '{A171FDEC-A09F-4EE9-9D4E-D326F8FC40C3}'
    '{593415AE-79DC-4C83-B95F-D3BF0623CC0A}'
    '{401ED04C-D64D-4132-B729-496BE9361CDC}'
    '{2BD32D76-AAFE-4D32-BCF9-D5055B1F244A}'
    '{C509AD6E-9754-4511-8E24-6DA662DC2F7A}'
    )
#Declaring Known CSUA Product IDs
    $CSUAProdID = @(
    '845AD9CBFBF00274C9D7CC299CB16432'
    '99E0C30CB22F9534F9D9AA7921E339C7'
    '178B12524DAA92B4C86C36FCF296DAE8'
    '8D4AE8FA4B3144F4EAF00B72B23C89AA'
    '1D78411D6A655134193C6D08C08ABFEC'
    'DA4244E8F11FFBA48B25FF54ACF0A772'
    '1C29D1830674C5E46B0C8EAE50D57E13'
    '4DA9A035D33B96B42B703079E9BA42F6'
    '72EC9DB055658D34C934FB9A20B6C476'
    '00111F4338EBB544FBEAA272F73A59E9'
    '0CC0AFFC7DC3136438D983746CA74721'
    '80AA0F76D16CBC54E8C59EFB4BFE9C47'
    '5BB9EBF981B8268498456BF2F241753E'
    'CDB8F809F63E58A43B06DAF83F378A50'
    'FDF8028BE371E834EB91B36F503C0574'
    '67AB3B9370F9FD54B92A2427BABBAB15'
    '72D5F7FD14E52A948831E49F62220568'
    '44B1DC3D4F178C64B862207219CD0676'
    '92A6CE76DE59CF048A5C8BD974610716'
    'CEDF171AF90A9EE4D9E43D628FCF043C'
    'EA514395CD9738C49BF53DFB6032CCA0'
    'C40DE104D46D23147B9294B69E63C1CD'
    '67D23DB2EFAA23D49FCBA442F1B5505D'
    'E6DA905C45791154E842D66A26CDF2A7'
    )
#Declaring Component IDs
    $ComponentID = @(
    '08C2F4D0EC2F8705882282685B8C2AF7'
    '0FB2DCFBD81742C54A07374C9183AE11'
    '25029BFCF4B3B7E56894506915234549'
    '3A7E404BA977E695BA8F14D32E87071C'
    '41FAE916668ADF756B9AAE81216978C0'
    '4840992B9A56CE85998E24CD4395D0A8'
    '746C8D349DB1F775DB5123F57C01ADE8'
    '7F9C1A2B86565A95D8C51AAC57A3C839'
    '81D7F39E4CB41A059B3AF8D0DE7F92CB'
    '82E455B753812505AB86B8357A7CD622'
    '991285CE294111D549229DC3C2146C6B'
    '9ABE4934A889CB5559ED3EDE33180B30'
    'A3DC272CC180D1850A5D4A996FD527F8'
    'B8AB1AC9AF4D8D856A9B2D7818307D20'
    'BC80D7512767B7957BD158BFE3F45A3E'
    'D2FF0449664C57E409661C68FE6275AE'
    'D71C0E6BFFA09715BA73B4DB1D29A81B'
    )
#Searching for unknown Product IDs
$SearchID1 = @(Get-ChildItem -Path "HKCU:\Software\Microsoft\Installer\Products\"-recurse
Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Products" -recurse
Get-ChildItem -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
)
$SearchID2 = $SearchID1 | ForEach-Object {
                                Get-ItemProperty -path Registry::$psitem | Where-Object {$_.Displayname -clike "*Cloud Signature Update Agent*"} | Select-Object PSParentPath
                           }
$SearchID3 = $SearchID2 -replace ".*\\" -replace "}"
#Adding new Product IDs to array
$AddID = @()
ForEach ($SearchID in $SearchID3)
{
    If ($CSUAProdID -NotContains $SearchID)
     {
        $AddID += $SearchID
     }
}
$CSUAProdID += $AddID
$SearchID1 = $null
$SearchID2 = $null
$SearchID3 = $null
$AddID = $null
#Searching for unknown Product Codes
$SearchCode1 = Get-ChildItem -path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
$SearchCode2 = $SearchCode1 | ForEach-Object {
Get-ItemProperty -path Registry::$psitem | Where-Object {$_.Displayname -clike "*Cloud Signature Update Agent*"} | Select-Object PSPath
}
$SearchCode3 = $SearchCode2 -replace ".*\\" -replace "}}", "}"
#Adding new Product Codes to array
$AddCode = @()
ForEach ($SearchCode in $SearchCode3)
{
    If ($CSUAProdCode -NotContains $SearchCode)
     {
        $AddCode += $SearchCode
     }
}
$CSUAProdCode += $AddCode
$SearchCode1 = $null
$SearchCode2 = $null
$SearchCode3 = $null
$AddCode = $null
#Declaring Group Policy deployment keys to be removed
$GPRegKey = @(
)
#Searching for User-specific Group Policy deployments
$SearchUGP1 = Get-ChildItem -path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Appmgmt"
$SearchUGP2 = $SearchUGP1 | ForEach-Object {
Get-ItemProperty -path Registry::$psitem | Where-Object {$_."Deployment Name" -clike "*Cloud Signature Update Agent*"} | Select-Object PSPath
}
$SearchUGP3 = $SearchUGP2 -replace".*{", "{" -replace "}}", "}"
#Adding User-specific Group Policy IDs to array
ForEach ($SearchUGP in $SearchUGP3) { $GPRegKey += "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Appmgmt\$SearchUGP"}
$SearchUGP1 = $null
$SearchUGP2 = $null
$SearchUGP3 = $null
#Searching for machine-specific Group Policy deployments
$SearchMGP1 = Get-ChildItem -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Appmgmt"
$SearchMGP2 = $SearchMGP1 | ForEach-Object {
Get-ItemProperty -path Registry::$psitem | Where-Object {$_."Deployment Name" -clike "*Cloud Signature Update Agent*"} | Select-Object PSPath
}
$SearchMGP3 = $SearchMGP2 -replace".*{", "{" -replace "}}", "}"
#Adding machine-specific Group Policy IDs to array
ForEach ($SearchMGP in $SearchMGP3) { $GPRegKey += "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Appmgmt\$SearchMGP"}
$SearchMGP1 = $null
$SearchMGP2 = $null
$SearchMGP3 = $null
#Declaring Product ID-specific registry keys to be removed
    $CSUAPIDSRegKey = $CSUAProdID | foreach-object {
    "HKCR:\Installer\Features\$PSItem"
    "HKCR:\Installer\Products\$PSItem"
    "HKCU:\Software\Microsoft\Installer\Features\$PSItem"
    "HKCU:\Software\Microsoft\Installer\Products\$PSItem"
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$PSItem"
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Products\$PSItem"
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Managed\$CUSID\Installer\Products\$PSItem"
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Managed\$CUSID\Installer\Features\$PSItem"
    "HKLM:\SOFTWARE\Classes\Installer\Products\$PSItem"
    }
#Declaring Product Code-specific registry keys to be removed
    $CSUAPCSRegKey = $CSUAProdCode | foreach-object { 
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$PSItem"
    "HKLM:\SOFTWARE\Microsoft\EnterpriseDesktopAppManagement\$CUSID\MSI\$PSItem"
    }
#Declaring Component ID-specific registry keys to be removed
    $CSUACIDSRegKey = $ComponentID | foreach-object { 
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\$PSItem"
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\$CUSID\Components\$PSItem"
    }
#Declaring generic registry keys to be removed
    $CSUAGenRegKey = @(
    'HKCU:\SOFTWARE\Exclaimer Ltd\CloudSignatureUpdateAgent'
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Exclaimer Cloud Signature Update Agent'
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\'
    'HKLM:\SOFTWARE\WOW6432Node\Exclaimer Ltd\CloudSignatureUpdateAgent'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Exclaimer Cloud Signature Update Agent'
    )
#Combining Reg key Arrays
    $CSUARegKey= $CSUAPIDSRegKey + $CSUAPCSRegKey + $CSUACIDSRegKey + $CSUAGenRegKey + $GPRegKey
    $CSUAPIDSRegKey = $null
    $CSUAPCSRegKey = $null
    $CSUACIDSRegKey = $null
    $CSUAGenRegKey = $null
    $GPRegKey = $null
#Declaring Product ID-specific registry entries to be removed
    $CSUAPIDSRegEnt = $CSUAProdID | foreach-object {
        [pscustomobject]@{CSUARegPath='HKCR:\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5\';CSUARegName=$PSItem}
        [pscustomobject]@{CSUARegPath='HKCU:\Software\Microsoft\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5\';CSUARegName=$PSItem}
        [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5\';CSUARegName=$PSItem}
        [pscustomobject]@{CSUARegPath="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Managed\$CUSID\Installer\UpgradeCodes\5D8BC74BCE235884C8D7107332DA40E5\";CSUARegName=$PSItem}

    }
#Declaring Product Code-specific registry entries to be removed
$CSUAPCodeSRegEnt = $CSUAProdCode | foreach-object {
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:APPDATA\Microsoft\Installer\$PSItem\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="C:\Windows\Installer\$PSItem\"}
}
#Declaring registry entries to be removed
    $CSUAGenRegEnt = @(
    [pscustomobject]@{CSUARegPath='HKCU:\Software\Microsoft\Windows\CurrentVersion\Run\';CSUARegName='Cloud Signature Update Agent'}
    [pscustomobject]@{CSUARegPath='HKCU:\Software\Microsoft\Windows\CurrentVersion\Run\';CSUARegName='Exclaimer Cloud Signature Update Agent'}
    [pscustomobject]@{CSUARegPath='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run';CSUARegName='Exclaimer Cloud Signature Update Agent'}
    [pscustomobject]@{CSUARegPath='HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run';CSUARegName='Exclaimer Cloud Signature Update Agent'}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\'}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\'}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\de\'}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\es\'}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\fr\'}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\it\'}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\nl\'}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName='C:\Program Files (x86)\Exclaimer Ltd\Cloud Signature Update Agent\pt\'}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Exclaimer\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\de\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\fr\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\it\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\nl\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:LOCALAPPDATA\Programs\Exclaimer Ltd\Cloud Signature Update Agent\pt\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders\';CSUARegName="$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Exclaimer\"}
    [pscustomobject]@{CSUARegPath='HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run\';CSUARegName='Cloud Signature Update Agent'}
    )
#Combining Reg entry Arrays
$CSUARegEntries = $CSUAPIDSRegEnt + $CSUAPCodeSRegEnt + $CSUAGenRegEnt
$CSUAPIDSRegEnt = $null
$CSUAPCodeSRegEnt = $null
$CSUAGenRegEnt = $null
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
    $CSUAProdCode = $null
    $CSUAProdID = $null
    $ComponentID = $null
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