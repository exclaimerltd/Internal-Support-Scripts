' Script checks for the presence of an Outlook profile
' prior to running ExSync
 
Const HKEY_CURRENT_USER = &H80000001
strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
strKeyPath2010 = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
strKeyPath2013 = "Software\Microsoft\Office\15.0\Outlook\Profiles"
strKeyPath2016 = "Software\Microsoft\Office\16.0\Outlook\Profiles"
strValueName = "DefaultProfile"
objRegistry.GetStringValue HKEY_CURRENT_USER,strKeyPath2010,strValueName,dwValue
objRegistry.EnumKey HKEY_CURRENT_USER,strKeyPath2013, arrSubKeys2013
objRegistry.EnumKey HKEY_CURRENT_USER,strKeyPath2016, arrSubKeys2016
 
If IsNull(dwValue) AND IsNull(arrSubKeys2013) AND IsNull(arrSubKeys2016) Then
'    Wscript.Echo "No Outlook Profile Exists"
Else
   call signature()
End If
 
Function signature()
Set wshShell = WScript.CreateObject ("WSCript.shell")
wshshell.run """\\EXCHANGE\Signatures\exsync.exe""", 6, True
End Function