' VBS function that will run msiexec
' Can be added as part of a larger script

Function InstallAgent ()
    Set wshShell = WScript.CreateObject ("WScript.Shell")
    sCmd = "msiexec /i ""MSI LOCATION HERE"" /qn"
    wshShell.Run sCmd ,1,True
End Function