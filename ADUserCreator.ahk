#SingleInstance force
#NoEnv

domain := "domain.com" ;Specifies UserPrincipalName
groupList := "|test1|test2"
departmentList := "|Administration|Marketing|Sales"
;Get-ADUser -Filter '*' -Property department | select department | sort-object department -unique

ini := A_ScriptDir "\ini.ini"
IniRead, ouListFromIni, % ini, ouList
Loop, Parse, ouListFromIni, `n
	ouList .= StrSplit(A_LoopField,"=")[1] "|"
if !ouList
{
	MsgBox, 16, , % "Error getting OU list.`nPlease check the ini file."
	ExitApp
}
;Get-ADOrganizationalUnit -Filter '*' -SearchBase 'ou=IT,dc=domain,dc=com' | Sort-Object -Property Name,DistinguishedName | Format-Table Name,DistinguishedName -A


editBoxes := "firstNameEng|lastNameEng|firstNameHeb|lastNameHeb|displayName|commandBox|phone|logonName|password"
ddlAndCombo := "department|group|ouPath"
editBoxesUpperCase := "firstNameEng|lastNameEng|firstNameHeb|lastNameHeb"



Gui, Margin, 20, 20
right := "+e0x400000"

;English
Gui, Font, s10 norm, Arial
Gui, Add, GroupBox, w450 h100, English
Gui, Font, s12 bold
Gui, Add, Text, xp+15 yp+25 Section, ★ First name
Gui, Add, Edit, y+10 w200 vfirstNameEng gUpdateCommand
Gui, Add, Text, x+20 ys, ★ Last name
Gui, Add, Edit, y+10 w200 vlastNameEng gUpdateCommand

;Hebrew
Gui, Font, s10 norm, Arial
Gui, Add, GroupBox, xm y+28 w450 h100, Hebrew
Gui, Font, s12 bold
Gui, Add, Text, xp+15 yp+25 Section, שם פרטי
Gui, Add, Edit, y+10 w200 vfirstNameHeb gUpdateCommand %right%
Gui, Add, Text, x+20 ys, שם משפחה
Gui, Add, Edit, y+10 w200 vlastNameHeb gUpdateCommand %right%

;Display Name
Gui, Font, s10 norm, Arial
Gui, Add, GroupBox, xm y+28 w450 h100,
Gui, Font, s12 bold
Gui, Add, Text, xp+15 yp+25 Section, Display Name (final)
Gui, Add, Edit, y+10 w420 vdisplayName gUpdateCommand hwndhDisplayName

;Command Box
Gui, Font, s10 norm, Arial
Gui, Add, GroupBox, xm y+28 w450 h257 Section, Command
Gui, Add, Edit, xp+15 yp+25 w420 r13 ReadOnly vcommandBox

;Right
Gui, Font, s10 norm, Arial
Gui, Add, GroupBox, x+30 ym w230 h458 Section, Details
Gui, Font, s12 bold
Gui, Add, Text, xp+15 yp+25, ★ Department
Gui, Add, ComboBox, y+10 w200 vdepartment gUpdateCommand, % departmentList
Gui, Add, Text, y+15, AD Group
Gui, Add, DDL, y+10 w200 vgroup gUpdateCommand, % groupList
Gui, Add, Text, y+15, ★ OU path
Gui, Add, DDL, y+10 w200 vouPath gUpdateCommand, % ouList
Gui, Add, Text, y+15, Phone
Gui, Add, Edit, y+10 w200 vphone gUpdateCommand

;Logon name
Gui, Add, Text, y+10 w200 h2 0x7 ;Vertical Line
Gui, Add, Text, y+8, ★ Logon Name
Gui, Add, Edit, y+10 w200 vlogonName gUpdateCommand
Gui, Add, Text, y+15, Password
Gui, Add, Edit, y+10 w200 vpassword gUpdateCommand

;Finish
Gui, Font, s12 bold
Gui, Add, GroupBox, xs y+28 w230 h115, Finish
Gui, Font, s15 bold
Gui, Add, Button, xp+15 yp+25 w200 h75 vcreateButton gCreate Disabled, Create user

;Status Bar
Gui, Font, s10 norm
Gui, Add, StatusBar, vStatusBar -Theme, Ready

WM_CTLCOLOREDIT = 0x133
OnMessage( WM_CTLCOLOREDIT, "WM_CTLCOLOR" )
WM_COMMAND := 0x111
OnMessage( WM_COMMAND, "WM_CTRLFUCUS" )

Gui, Show
return


UpdateCommand:
Gui, Submit, NoHide
gosub, UpdateGui

;Clean vars
creatADUserCommand := firstNameEngCommand := lastNameEngCommand := departmentCommand := ouPathCommand := phoneCommand := logonNameCommand := nameCommand := displayNameEngCommand := updateADUserDetailsCommand := firstNameHebCommand := lastNameHebCommand := groupCommand := ""

;New-ADUser command
if firstNameEng || lastNameEng
	creatADUserCommand := "New-ADUser -Enabled $true -ChangePasswordAtLogon $true -AccountPassword (ConvertTo-SecureString """ (password?password:"Aa123456") """ -AsPlainText -Force) "

if firstNameEng
	firstNameEngCommand := "-GivenName """ firstNameEng """ "
if lastNameEng
	lastNameEngCommand := "-Surname """ lastNameEng """ "
nameCommand := "-Name """ firstNameEng " " lastNameEng """ "
displayNameEngCommand := "-DisplayName """ lastNameEng " " firstNameEng """ "

if department
	departmentCommand := "-Department """ department """ "
if ouPath
{
	IniRead, ouPath, % ini, ouList, % ouPath
	ouPathCommand := "-Path """ ouPath """ "
}
if phone
	phoneCommand := "-OfficePhone """ phone """ "
if logonName
	logonNameCommand := "-SamAccountName """ logonName """ -UserPrincipalName """ logonName "@" domain """ -ScriptPath ""NETLOGON.BAT"" "

if creatADUserCommand
	creatADUserCommand .= firstNameEngCommand . lastNameEngCommand . nameCommand . displayNameEngCommand . departmentCommand . ouPathCommand . phoneCommand . logonNameCommand . "`n`n"


;Set-ADUser command
if logonName
{
	if firstNameHeb
		firstNameHebCommand := "-GivenName """ firstNameHeb """ "
	if lastNameHeb
		lastNameHebCommand := "-Surname """ lastNameHeb """ "
	if displayName && (displayName != lastNameEng " " firstNameEng)
		updateADUserDetailsCommand := "Set-ADUser -Identity """ logonName """ -DisplayName """ displayName """ " . firstNameHebCommand . lastNameHebCommand "`n`n"
}

;Add-ADGroupMember command
if logonName && group
	groupCommand := "Add-ADGroupMember -Identity """ group """ -Members """ logonName """"


finalCommand := creatADUserCommand . updateADUserDetailsCommand . groupCommand
GuiControl,, commandBox, % finalCommand

if finalCommand && logonName && ouPath && department ;Can be updated as required
	GuiControl, Enable, createButton
else
	GuiControl, Disable, createButton
return


UpdateGui:
;Uppercase / lowercase letters
Loop, Parse, editBoxesUpperCase, "|"
	%A_LoopField% := str(%A_LoopField%)
logonName := str(logonName,0)

if A_GuiControl in firstNameEng,lastNameEng,firstNameHeb,lastNameHeb
{
	if firstNameHeb || lastNameHeb
	{
		displayName := lastNameHeb " " firstNameHeb
		displayNameLang := "Heb"
	}
	else
	{
		displayName := lastNameEng " " firstNameEng
		displayNameLang := "Eng"
	}
	GuiControl,, displayName, % displayName
}
return


Create:
GuiControl, Disable, createButton
Gui, Submit, NoHide

if ADUserExiat(logonName)
{
	;~ MsgBox, 16, , % "'" logonName "' already exists."
	SB_SetText("'" logonName "' already exists.")
	GuiControl, +BackgroundRed, StatusBar
	GuiControl, Focus, logonName
	return
}
if InStr(logonName, A_Space) || RegExMatch(logonName, "[^a-zA-Z0-9\s]")
{
	;~ MsgBox, 16, , % "Invalid username."
	SB_SetText("Invalid username.")
	GuiControl, +BackgroundRed, StatusBar
	GuiControl, Focus, logonName
	return
}
password := (password?password:"Aa123456")
if !Password_Complexity(password)
{
	;~ MsgBox, 16, , % "Password does not meet complexity requirements."
	SB_SetText("Password does not meet complexity requirements.")
	GuiControl, +BackgroundRed, StatusBar
	GuiControl, Focus, password
	return
}

SB_SetText("Creating...")
GuiControl, -Background, StatusBar

powershellFilePath := temp "\CreateADUserTempFile.ps1"
IfExist, % powershellFilePath
	FileDelete, % powershellFilePath
Sleep 50
FileAppend, % finalCommand, % powershellFilePath, UTF-8
RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %powershellFilePath%,, Hide
Sleep 50
FileDelete, % powershellFilePath

if ADUserExiat(logonName)
{
	;~ MsgBox, 64, , User created successfully.
	SB_SetText("User '" logonName "' was successfully created.")
	GuiControl, +BackgroundGreen, StatusBar
	Loop, Parse, editBoxes, |
		GuiControl,, % A_LoopField
	Loop, Parse, ddlAndCombo, |
		GuiControl, Choose, % A_LoopField, 0
}
else
{
	;~ MsgBox, 16, , % "Error creating user."
	SB_SetText("Error creating user.")
	GuiControl, +BackgroundRed, StatusBar
}
return

GuiClose:
ExitApp


WM_CTLCOLOR( wParam, lParam, msg, hwnd ) 
{
	Static hBrush

	BG_COLOR := 0xFFFFFF
	hBrush := DllCall( "CreateSolidBrush", UInt,BG_COLOR )
	
	BG_COLOR := 0x00FFFF	; Changes only when focused.
	FG_COLOR := 0x000000
	
	DllCall( "SetTextColor", UInt,wParam, UInt,FG_COLOR )
	DllCall( "SetBkColor", UInt,wParam, UInt,BG_COLOR )
	DllCall( "SetBkMode", UInt,wParam, UInt,2 )
	Return hBrush
}

WM_CTRLFUCUS(wParam,lParam,msg,hwnd)
{
	global
	if (lParam = hDisplayName)
	{
		currenKeyboardLayout := DllCall("GetKeyboardLayout", "UInt", 0, "UInt")
		if (currenKeyboardLayout = 67961869) || (currenKeyboardLayout = 4030530573) ;HEB
		{
			GuiControl, +e0x002000, displayName ;WS_EX_RTLREADING
			GuiControl, +Right, displayName
		}
		else ;if (currenKeyboardLayout = 67699721) ;ENG
		{
			GuiControl, -e0x002000, displayName
			GuiControl, -Right, displayName
		}
	}
}

ADUserExiat(_Item)
{
	objRootDSE := ComObjGet("LDAP://rootDSE")
	strDomain := objRootDSE.Get("defaultNamingContext")
	strADPath := "LDAP://" . strDomain
	objDomain := ComObjGet(strADPath)
	objConnection := ComObjCreate("ADODB.Connection")
	objConnection.Open("Provider=ADsDSOObject")
	objCommand := ComObjCreate("ADODB.Command")
	objCommand.ActiveConnection := objConnection

	objCommand.CommandText := "<" . strADPath . ">;(|(sAMAccountName=" . _Item . "));Name;subtree"
	
	objRecordSet := objCommand.Execute
	objRecordCount := objRecordSet.RecordCount
	
	objConnection.Close
	objRelease(objRootDSE)
	objRelease(objDomain)
	objRelease(objConnection)
	objRelease(objCommand)

	return objRecordCount
}

Password_Complexity(var) {
	RegExReplace(var,"[A-Z]","$0",UC)
	RegExReplace(var,"[a-z]","$0",LC)
	RegExReplace(var,"\d","$0",NC)
	;I still haven't found a good method for finding special characters
	return StrLen(var) >= 8 && UC >= 1 && LC >= 1 && NC >= 1
}

str(var,upper:=1) {
	if upper
		var := Regexreplace(var ,"(\b\w)(.*?)","$U1$2")
	else
		var := Format("{:L}", var)
	var = %var%
	return var
}