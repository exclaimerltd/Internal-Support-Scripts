' CHECK FOR OSUA AGENT INSTALL
 
' Registry constants
Const HKEY_CLASSES_ROOT   = &H80000000
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKEY_USERS          = &H80000003
 
Const REG_SZ        = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY    = 3
Const REG_DWORD     = 4
Const REG_MULTI_SZ  = 7
 
strComputer = "."
 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Run"
strValueName = "Outlook Signature Update Agent"
 
objRegistry.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue
 
If IsNull(strValue) Then
    wscript.echo "Outlook Signature Update Agent not installed"
Else   
    wscript.echo "Outlook Signature Update Agent installed"
End If