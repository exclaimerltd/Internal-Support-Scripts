'on error resume next

'AD Constants
'Group constants
Const ADS_PROPERTY_APPEND = 3
Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &H8

'sAMAccountType constants
Const SAT_NORMAL_USER_ACCOUNT = 805306368

'Initialise list of countries
dim arrCountry(50)
arrCountry(1) = "United States"
arrCountry(2) = "China"
arrCountry(3) = "Japan"
arrCountry(4) = "Germany"
arrCountry(5) = "France"
arrCountry(6) = "Brazil"
arrCountry(7) = "United Kingdom"
arrCountry(8) = "Italy"
arrCountry(9) = "India"
arrCountry(10) = "Russia"
arrCountry(11) = "Canada"
arrCountry(12) = "Australia"
arrCountry(13) = "Spain"
arrCountry(14) = "Mexico"
arrCountry(15) = "South Korea"
arrCountry(16) = "Indonesia"
arrCountry(17) = "Netherlands"
arrCountry(18) = "Turkey"
arrCountry(19) = "Switzerland"
arrCountry(20) = "Saudi Arabia"
arrCountry(21) = "Sweden"
arrCountry(22) = "Iran"
arrCountry(23) = "Belgium"
arrCountry(24) = "Poland"
arrCountry(25) = "Norway"
arrCountry(26) = "Argentina"
arrCountry(27) = "Austria"
arrCountry(28) = "South Africa"
arrCountry(29) = "Thailand"
arrCountry(30) = "United Arab Emirates"
arrCountry(31) = "Columbia"
arrCountry(32) = "Denmark"
arrCountry(33) = "Venezuela"
arrCountry(34) = "Greece"
arrCountry(35) = "Malaysia"
arrCountry(36) = "Finland"
arrCountry(37) = "Singapore"
arrCountry(38) = "Chile"
arrCountry(39) = "Nigeria"
arrCountry(40) = "Israel"
arrCountry(41) = "Portugal"
arrCountry(42) = "Egypt"
arrCountry(43) = "Philippines"
arrCountry(44) = "Ireland"
arrCountry(45) = "Czech Republic"
arrCountry(46) = "Pakistan"
arrCountry(47) = "Algeria"
arrCountry(48) = "Romania"
arrCountry(49) = "Kazakhstan"
arrCountry(50) = "Peru"

'Initialise list of offices
dim arrOffice(5)
arrOffice(1) = "Central"
arrOffice(2) = "Northern"
arrOffice(3) = "Southern"
arrOffice(4) = "Eastern"
arrOffice(5) = "Western"

'Initialise list of departments
dim arrDept(5)
arrDept(1) = "Development"
arrDept(2) = "Engineering"
arrDept(3) = "Marketing"
arrDept(4) = "Sales"
arrDept(5) = "Support"

'Initialise list of titles
dim arrDeptTitle(5)
arrDeptTitle(1) = "Developer"
arrDeptTitle(2) = "Engineer"
arrDeptTitle(3) = "Marketing Consultant"
arrDeptTitle(4) = "Sales Person"
arrDeptTitle(5) = "Support Engineer"

'Ask running user for company parameters
strCompany = CStr(inputbox("Please enter the Company Name."))
intNoUsers = CLng(inputbox("How many users do you require?" & vbcrlf & "Max 500000"))
intNoCountry = CInt(inputbox("How many Countries do you require?" & vbcrlf & "Max 50"))
intNoOffice = CInt(inputbox("How many Offices do you require?" & vbcrlf & "Max 5"))
intNoDept = CInt(inputbox("How many Departments per Office do you require?" & vbcrlf & "Max 5"))
intAddressList = MsgBox("Do you require an address list to be written" & vbcrlf & "for the users added to AD?", 4, "Address list")

if intAddressList = 6 then

	strAddresslistFile = CStr(inputbox("Please enter the file name for the address list, excluding extension"))
	
end if

'Open file containing list of names
dim fs,objTextFile
set fs=CreateObject("Scripting.FileSystemObject")
set objTextFile = fs.OpenTextFile("Names.csv")

if intAddressList = 6 then

	dim fsWrite,objWriteTextFile
	set fsWrite=CreateObject("Scripting.FileSystemObject")
	set objWriteTextFile = fsWrite.CreateTextFile(strAddressListFile & ".txt", true)
	
end if

'Split requested numbers of users across countries/departments
intUsersPerCountry = CLng(Fix(intNoUsers / intNoCountry))
intUsersPerOffice = CLng(Fix(intUsersPerCountry / intNoOffice))
intUsersPerDept = CLng(Fix(intUsersPerOffice/ intNoDept))
intNatUsers = CLng(intUsersPerDept * intNoOffice * intNoDept * intNoCountry)
intAddUsers = CLng(intNoUsers - intNatUsers)

WScript.Echo "Users per country: " & intUsersPerCountry & vbcrlf & "Users per office: " & intUsersPerOffice & vbcrlf & "Users per department: " & intUsersPerDept & vbcrlf & "Natural users: " & intNatUsers & vbcrlf & "Additional users: " & intAddUsers

'Get domain
Set objRoot = GetObject("LDAP://rootDSE")
objDomain = objRoot.Get("defaultNamingContext")
Set objDomain = GetObject("LDAP://" & objDomain)
strLocPart = objRoot.Get("defaultNamingContext")
strLocPart = replace(strLocPart, "DC=", "")
strLocPart = replace(strLocPart, ",", ".")
strDomain = objRoot.Get("defaultNamingContext")

WScript.Echo "Domain local part: " & strLocPart

'Define user password/userAccountControl value
strPassword = "Rainbow1"
intAccValue = 66112
intPwdValue = -1

'Create company root OU
strOUContainerRoot = "OU=" & strCompany

Set objOU = objDomain.Create("organizationalUnit", strOUContainerRoot)
objOU.Put "Description", strCompany
objOU.SetInfo

WScript.Echo "New Company OU created: " & strOUContainerRoot

'Group constants
'Const ADS_PROPERTY_APPEND = 3
'Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &H8

'Create company root Distribution Group
set objCompanyDG = objOU.Create("group", "cn=" & strCompany)
objCompanyDG.Put "sAMAccountName", strCompany
objCompanyDG.Put "displayName", strCompany
objCompanyDG.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP
objCompanyDG.SetInfo

if intAddressList = 6 then

	objWriteTextFile.WriteLine(strCompany & "|" & strCompany & "@" & strLocPart & "|SMTP")
	
end if

WScript.Echo "New Company DG created: " & strCompany

'Create country OUs
for i = 1 to intNoCountry

	strOUContainer = "OU=" & arrCountry(i) & "," & strOUContainerRoot
	Set objOU = objDomain.Create("organizationalUnit", strOUContainer)
	objOU.Put "Description", arrCountry(i)
	objOU.SetInfo
	
	WScript.Echo "New Country OU created: " & strOUContainer
	
	'Create country root Distribution Group
	set objCountryDG = objOU.Create("group", "cn=" & arrCountry(i))
	objCountryDG.Put "sAMAccountName", arrCountry(i)
	objCountryDG.Put "displayName", arrCountry(i)
	objCountryDG.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP
	objCountryDG.SetInfo
	
	objCompanyDG.PutEx ADS_PROPERTY_APPEND, "member", Array("cn=" & arrCountry(i) & "," & strOUContainer & "," & strDomain)
	objCompanyDG.SetInfo
	
	if intAddressList = 6 then
	
		objWriteTextFile.WriteLine(arrCountry(i) & "|" & arrCountry(i) & "@" & strLocPart & "|SMTP")
	
	end if
	
	WScript.Echo "New Country DG created: " & arrCountry(i)
	
	'Create Office OUs as children of country OU
	for j = 1 to intNoOffice
	
		strOUContainer1 = "OU=" & arrOffice(j) & "," & strOUContainer
		set objOU = objDomain.Create("organizationalUnit", strOUContainer1)
		objOU.Put "Description", arrOffice(j)
		objOU.SetInfo
		
		wscript.echo "New Office OU created: " & strOUContainer1
		
		'Create Office head user account
		arrName = Split(objTextFile.ReadLine, ",")
		strOfficeHeadInitial = arrName(1)
		strOfficeHeadInitial = replace(strOfficeHeadInitial, ".", "")
		strOfficeHeadInitial = trim(strOfficeHeadInitial)
			
		Dim strOfficeHeadSAN
		Dim objNewOfficeHead
		
		if len(strOfficeHeadInitial) = 0 then
		
			strOfficeHeadSAN = trim(arrName(0)) & "." & trim(arrName(2))
			
		else
		
			strOfficeHeadSAN = trim(arrName(0)) & "." & strOfficeHeadInitial & "." & trim(arrName(2))
			
		end if
		
		Set objNewOfficeHead = objOU.Create("User", "cn=" & strOfficeHeadSAN)
		
		objNewOfficeHead.Put "userPrincipalName", strOfficeHeadSAN & "@" & strLocPart
		
		if len(strOfficeHeadSAN) > 20 then
		
			objNewOfficeHead.Put "sAMAccountName", mid(strOfficeHeadSAN, 1, 20)
			
		else
		
			objNewOfficeHead.Put "sAMAccountName", strOfficeHeadSAN
			
		end if
		
		objNewOfficeHead.Put "givenName", trim(arrName(0))
		
		if len(strOfficeHeadInitial) > 0 then
		
			objNewOfficeHead.Put "initials", strOfficeHeadInitial
			
		end if
		
		objNewOfficeHead.Put "SN", trim(arrName(2))
		
		if len(strOfficeHeadInitial) = 0 then
	
			strOfficeHeadDN = trim(arrName(0)) & " " & trim(arrName(2))
		
		else
	
			strOfficeHeadDN = trim(arrName(0)) & " " & strOfficeHeadInitial & ". " & trim(arrName(2))
		
		end if
	
		objNewOfficeHead.Put "displayName", strOfficeHeadDN
		objNewOfficeHead.Put "title", arrOffice(j) & " Manager"
		objNewOfficeHead.Put "mail", strOfficeHeadSAN & "@" & strLocPart
		'objNewOfficeHead.Put "Office", arrOffice(j)
		objNewOfficeHead.Put "company", strCompany
		
		objNewOfficeHead.SetInfo
		
		objNewOfficeHead.SetPassword strPassword
		objNewOfficeHead.SetInfo
		
		objNewOfficeHead.Put "pwdLastSet", intPwdValue
		objNewOfficeHead.SetInfo
		
		objNewOfficeHead.Put "userAccountControl", intAccValue
		objNewOfficeHead.SetInfo
		
		if intAddressList = 6 then
		
			objWriteTextFile.WriteLine(strOfficeHeadDN & "|" & strOfficeHeadSAN & "@" & strLocPart & "|SMTP")
		
		end if
		
		WScript.Echo "New Office Head: " & strOfficeHeadSAN & " created"
		
		'Create Office Distribution Group
		set objOfficeDG = objOU.Create("group", "cn=" & arrOffice(j) & "-" & arrCountry(i))
		objOfficeDG.Put "sAMAccountName", arrOffice(j) & "-" & arrCountry(i)
		objOfficeDG.Put "displayName", arrOffice(j) & "-" & arrCountry(i)
		objOfficeDG.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP
		objOfficeDG.Put "managedBy", "cn=" & strOfficeHeadSAN & "," & strOUContainer1 & "," & strDomain
		objOfficeDG.SetInfo
		
		objOfficeDG.PutEx ADS_PROPERTY_APPEND, "member", Array("cn=" & strOfficeHeadSAN & "," & strOUContainer1 & "," & strDomain)
		objOfficeDG.SetInfo
		
		objCountryDG.PutEx ADS_PROPERTY_APPEND, "member", Array("cn=" & arrOffice(j) & "-" & arrCountry(i) & "," & strOUContainer1 & "," & strDomain)
		objCountryDG.SetInfo
		
		if intAddressList = 6 then
		
			objWriteTextFile.WriteLine(arrOffice(j) & "-" & arrCountry(i) & "|" & arrOffice(j) & "-" & arrCountry(i) & "@" & strLocPart & "|SMTP")
		
		end if
		
		WScript.Echo "New Office DG created: " & arrOffice(j)
		
			'Create department OUs as children of office OU
			for k = 1 to intNoDept
	
			strOUContainer2 = "OU=" & arrDept(k) & "," & strOUContainer1
			set objOU = objDomain.Create("organizationalUnit", strOUContainer2)
			objOU.Put "Description", arrDept(k)
			objOU.SetInfo
		
			wscript.echo "New Department OU created: " & strOUContainer2
		
				'Create department head user account
				arrName = Split(objTextFile.ReadLine, ",")
				strDeptHeadInitial = arrName(1)
				strDeptHeadInitial = replace(strDeptHeadInitial, ".", "")
				strDeptHeadInitial = trim(strDeptHeadInitial)
				
				Dim strDeptHeadSAN
				Dim objNewDeptHead
		
				if len(strDeptHeadInitial) = 0 then
		
					strDeptHeadSAN = trim(arrName(0)) & "." & trim(arrName(2))
			
				else
		
					strDeptHeadSAN = trim(arrName(0)) & "." & strDeptHeadInitial & "." & trim(arrName(2))
			
				end if
		
				Set objNewDeptHead = objOU.Create("User", "cn=" & strDeptHeadSAN)
		
				objNewDeptHead.Put "userPrincipalName", strDeptHeadSAN & "@" & strLocPart
		
				if len(strDeptHeadSAN) > 20 then
		
					objNewDeptHead.Put "sAMAccountName", mid(strDeptHeadSAN, 1, 20)
			
				else
		
					objNewDeptHead.Put "sAMAccountName", strDeptHeadSAN
			
				end if
		
				objNewDeptHead.Put "givenName", trim(arrName(0))
		
					if len(strDeptHeadInitial) > 0 then
		
						objNewDeptHead.Put "initials", strDeptHeadInitial
			
					end if
		
					objNewDeptHead.Put "SN", trim(arrName(2))
		
					if len(strDeptHeadInitial) = 0 then
	
						strDeptHeadDN = trim(arrName(0)) & " " & trim(arrName(2))
		
					else
	
						strDeptHeadDN = trim(arrName(0)) & " " & strDeptHeadInitial & ". " & trim(arrName(2))
		
					end if
	
					objNewDeptHead.Put "displayName", strDeptHeadDN
					objNewDeptHead.Put "title", "Senior " & arrDeptTitle(k)
					objNewDeptHead.Put "mail", strDeptHeadSAN & "@" & strLocPart
					objNewDeptHead.Put "department", arrDept(k)
					objNewDeptHead.Put "company", strCompany
		
					objNewDeptHead.SetInfo
		
					objNewDeptHead.SetPassword strPassword
					objNewDeptHead.SetInfo
		
					objNewDeptHead.Put "pwdLastSet", intPwdValue
					objNewDeptHead.SetInfo
		
					objNewDeptHead.Put "userAccountControl", intAccValue
					objNewDeptHead.SetInfo
		
					if intAddressList = 6 then
		
						objWriteTextFile.WriteLine(strDeptHeadDN & "|" & strDeptHeadSAN & "@" & strLocPart & "|SMTP")
		
					end if
		
					WScript.Echo "New Department Head: " & strDeptHeadSAN & " created"
		
					'Create department Distribution Group
					set objDeptDG = objOU.Create("group", "cn=" & arrDept(k) & "-" & arrOffice(j) & "-" & arrCountry(i))
					objDeptDG.Put "sAMAccountName", arrDept(k) & "-" & arrOffice(j) & "-" & arrCountry(i)
					objDeptDG.Put "displayName", arrDept(k) & "-" & arrOffice(j) & "-" & arrCountry(i)
					objDeptDG.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP
					objDeptDG.Put "managedBy", "cn=" & strDeptHeadSAN & "," & strOUContainer2 & "," & strDomain
					objDeptDG.SetInfo
		
					objDeptDG.PutEx ADS_PROPERTY_APPEND, "member", Array("cn=" & strDeptHeadSAN & "," & strOUContainer2 & "," & strDomain)
					objDeptDG.SetInfo
		
					objCountryDG.PutEx ADS_PROPERTY_APPEND, "member", Array("cn=" & arrDept(k) & "-" & arrOffice(j) & "-" & arrCountry(i) & "," & strOUContainer2 & "," & strDomain)
					objCountryDG.SetInfo
		
					if intAddressList = 6 then
		
						objWriteTextFile.WriteLine(arrDept(k) & "-" & arrOffice(j) & "-" & arrCountry(i) & "|" & arrDept(k) & "-" & arrOffice(j) & "-" & arrCountry(i) & "@" & strLocPart & "|SMTP")
		
					end if
		
					WScript.Echo "New Department DG created: " & arrDept(k)
		
					'Create department employees
						for l = 1 to (intUsersPerDept - 1)
		
						NewUser()
		
					next
		
				'Add 1 additional user to department if required
				if intAddUsers > 0 then
		
					NewUser()
					intAddUsers = intAddUsers - 1
			
				end if
			next
			
	next
next

objTextFile.Close
set objTextFile = Nothing
set fs = Nothing

if intAddressList = 6 then

	objWriteTextFile.Close
	set objWriteTextFile = Nothing
	set fsWrite = Nothing
	
end if

WScript.Quit

'Create user account and set password/useraccountcontrol
sub NewUser()

	arrName = Split(objTextFile.ReadLine, ",")
	strUserInitial = arrName(1)
	strUserInitial = replace(strUserInitial, ".", "")
	strUserInitial = trim(strUserInitial)
	
	Dim strUserSAN
	Dim objNewUser
	
	if len(strUserInitial) = 0 then
	
		strUserSAN = trim(arrName(0)) & "." & trim(arrName(2))
		
	else
	
		strUserSAN = trim(arrName(0)) & "." & strUserInitial & "." & trim(arrName(2))
		
	end if
	
	Set objNewUser = objOU.Create("User", "cn=" & strUserSAN)
	
	objNewUser.Put "userPrincipalName", strUserSAN & "@" & strLocPart
	
	if len(strUserSAN) > 20 then
	
		objNewUser.Put "sAMAccountName", mid(strUserSAN, 1, 20)
	
	else
	
		objNewUser.Put "sAMAccountName", strUserSAN
		
	end if
	
	objNewUser.Put "givenName", trim(arrName(0))
	
	if len(strUserInitial) > 0 then
	
		objNewUser.Put "initials", strUserInitial
		
	end if
	
	objNewUser.Put "SN", trim(arrName(2))
	
	if len(strUserInitial) = 0 then
	
		strUserDN = trim(arrName(0)) & " " & trim(arrName(2))
		
	else
	
		strUserDN = trim(arrName(0)) & " " & strUserInitial & ". " & trim(arrName(2))
		
	end if
	
	objNewUser.Put "displayName", strUserDN
	objNewUser.Put "title", arrDeptTitle(j)
	objNewUser.Put "mail", strUserSAN & "@" & strLocPart
	objNewUser.Put "department", arrDept(k)
	objNewUser.Put "company", strCompany
	objNewUser.Put "manager", "cn=" & strDeptHeadSAN & "," & strOUContainer2 & "," & strDomain
	
	objNewUser.SetInfo
	
	objNewUser.SetPassword strPassword
	objNewUser.SetInfo
	
	objNewUser.Put "pwdLastSet", intPwdValue
	objNewUser.SetInfo
	
	objNewUser.Put "userAccountControl", intAccValue
	objNewUser.SetInfo
	
	objDeptDG.PutEx ADS_PROPERTY_APPEND, "member", Array("cn=" & strUserSAN & "," & strOUContainer2 & "," & strDomain)
	objDeptDG.SetInfo
	
	if intAddressList = 6 then
	
		objWriteTextFile.WriteLine(strUserDN & "|" & strUserSAN & "@" & strLocPart & "|SMTP")
	
	end if
	
	WScript.Echo "New User: " & strUserSAN & " created"

end sub