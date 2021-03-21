' Installation script to install ExSync on a machine and pull through variables

if wscript.arguments.count < 2 then

	wscript.echo "Expected 2 arguments: " & Vbcrlf & _
	"1 - Signature Manager deployment location, e.g. http://servername/signatures" & vbcrlf & _
	"2 - Install location on this machine, e.g. c:\exsync"
	wscript.quit

end if

smlocation = wscript.arguments(0)
installlocation = wscript.arguments(1)


SET objNetwork = CREATEOBJECT("wscript.network")
Set WshShell = WScript.CreateObject("WScript.Shell")
strADsPath = getUser(objNetwork.Username)

SET objUser = GETOBJECT(strADsPath)

smtpaddress = objuser.mail

if smtpaddress = "" then

	wscript.echo "Unable to determine primary SMTP address"
	wscript.quit

end if

command = "msiexec /q /i ""Exclaimer Outlook Settings Update Client.msi"" UI_SMTPADDRESS=" & smtpaddress & " UI_REMOTELOCATIONURL=" & smlocation & " INSTALLLOCATION=""" & installlocation & """"

wshshell.run command


FUNCTION getUser(BYVAL UserName)

	DIM objRoot
	DIM getUserCn,getUserCmd,getUserRS

	SET objRoot = GETOBJECT("LDAP://RootDSE")

	SET getUserCn = CREATEOBJECT("ADODB.Connection")
	SET getUserCmd = CREATEOBJECT("ADODB.Command")
	SET getUserRS = CREATEOBJECT("ADODB.Recordset")

	getUserCn.open "Provider=ADsDSOObject;"
	
	getUserCmd.activeconnection=getUserCn
	getUserCmd.commandtext="<LDAP://" & objRoot.GET("defaultNamingContext") & ">;" & _
			"(&(objectCategory=user)(sAMAccountName=" & username & "));" & _
			"adsPath;subtree"


	
	SET getUserRs = getUserCmd.EXECUTE

	
	IF NOT getuserrs.BOF AND NOT getuserrs.EOF THEN
     		getUserRs.MoveFirst
     		getUser = getUserRs(0)
	ELSE
		getUser = ""
	END IF

	getUserCn.close
END FUNCTION