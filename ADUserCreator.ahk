#SingleInstance force
#NoEnv

domain := "domain.com" ;Specifies UserPrincipalName


ini := A_ScriptDir "\ini.ini"
IfNotExist, % ini
{ ;create ini template
	IniWrite, Example, % ini, departmentList
	IniWrite, =, % ini, ouList
	Sleep 100 ;To be sure the blank line will be first
	IniWrite, OU=Example`,DC=domain`,DC=com, % ini, ouList, Example
	IniWrite, Example, % ini, groupList, 1
}

;Get department list from ini
;Get-ADUser -Filter '*' -Property department | select department | sort-object department -unique
IniRead, departmentListFromIni, % ini, departmentList
Loop, Parse, departmentListFromIni, `n
	departmentList .= StrSplit(A_LoopField,"=")[1] "|"
if !departmentList
{
	MsgBox, 16, , % "Error getting department list.`nPlease check the ini file."
	ExitApp
}

;Get OU list from ini
;Get-ADOrganizationalUnit -Filter '*' -SearchBase 'ou=IT,dc=domain,dc=com' | Sort-Object -Property Name,DistinguishedName | Format-Table Name,DistinguishedName -A
IniRead, ouListFromIni, % ini, ouList
Loop, Parse, ouListFromIni, `n
	ouList .= StrSplit(A_LoopField,"=")[1] "|"
if !ouList
{
	MsgBox, 16, , % "Error getting OU list.`nPlease check the ini file."
	ExitApp
}

;Get group list from ini
IniRead, groupListFromIni, % ini, groupList
Loop, Parse, groupListFromIni, `n
{
	totalGroups++
	group%A_Index%name := StrSplit(A_LoopField,"=")[2]
	group%A_Index%state := false
}



;Create GUI

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
Gui, Font, s9 norm, Arial
Gui, Add, GroupBox, xm y+28 w450 h219 Section, Command
Gui, Add, Edit, xp+15 yp+30 w420 r11 ReadOnly vcommandBox

;Details
Gui, Font, s10 norm, Arial
Gui, Add, GroupBox, x+30 ym w230 h245 Section, Details
Gui, Font, s12 bold
Gui, Add, Text, xp+15 yp+25, ★ Department
Gui, Add, ComboBox, y+10 w200 vdepartment gUpdateCommand, % departmentList
Gui, Add, Text, y+15, ★ OU path
Gui, Add, DDL, y+10 w200 vouPath gUpdateCommand, % ouList
Gui, Add, Text, y+15, Phone
Gui, Add, Edit, y+10 w200 vphone gUpdateCommand number

;Logon name
Gui, Font, s10 norm, Arial
Gui, Add, GroupBox, xs y+28 w230 h175, Logon
Gui, Font, s12 bold
Gui, Add, Text, xp+15 yp+25, ★ Logon Name
Gui, Add, Edit, y+10 w200 vlogonName gUpdateCommand 0x10
Gui, Add, Text, y+15, Password
Gui, Add, Edit, y+10 w200 vpassword gUpdateCommand

;Finish
Gui, Font, s12 bold
Gui, Add, GroupBox, xs y+28 w230 h115, Finish
Gui, Font, s15 bold
Gui, Add, Button, xp+15 yp+25 w200 h75 vcreateButton gCreate Disabled, Create user

;AD-Groups ListBox
if groupListFromIni
{
	Gui, Font, s10 norm, Arial
	Gui, Add, GroupBox, x+30 ym w250 h546 Section, AD-Groups
	Gui, Add, Checkbox, xp+160 yp+15 vallGroupSelect gUpdateCommand, Select all
	Gui, Font, s12 bold
	Gui, Add, ListView, xs+15 ys+33 w220 h496 +LV0x14000 -Hdr Checked Grid AltSubmit -Multi vGroupLV gUpdateCommand, Group
		LV_ModifyCol(1, 199) ;to prevent hScroll
	Loop, % totalGroups
		LV_Add(group%A_Index%state, group%A_Index%name)
	;~ Gui, Font, s10 bold
	;~ Gui, Add, Button, xs-100 y4 h22 w80, Groups >>
}

;Status Bar
Gui, Font, s10 norm
Gui, Add, StatusBar, vStatusBar -Theme, Ready

WM_CTLCOLOREDIT = 0x133
OnMessage( WM_CTLCOLOREDIT, "WM_CTLCOLOR" )
WM_COMMAND := 0x111
OnMessage( WM_COMMAND, "WM_CTRLFUCUS" )

Gui, Show,, % "ADUserCreator - " domain " - " A_UserName
return
;~ Esc::ExitApp



UpdateCommand:
Gui, Submit, NoHide
gosub, UpdateGui
gosub, UpdateGroupList

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
		updateADUserDetailsCommand := "Set-ADUserrrr -Identity """ logonName """ -DisplayName """ displayName """ " . firstNameHebCommand . lastNameHebCommand "`n`n"
}

;Add-ADGroupMember command
if logonName && groupListCommand
	groupCommand := groupListCommand


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

UpdateGroupList:
if (A_GuiControl = "allGroupSelect")
{
	GuiControlGet, allGroupSelected,, allGroupSelect
	if (allGroupSelected = 1)
		LV_Modify("","Check")
	else
		LV_Modify("","-Check")
}

;Get selected groups
selectedGroupList := ""
while(i := LV_GetNext(i, "C")) ;Checked
{
	LV_GetText(selectedGroup, i)
	selectedGroupList .= (selectedGroupList?"|":"") selectedGroup
}

groupListCommand := ""
Loop, Parse, selectedGroupList, |
	groupListCommand .= "Add-ADGroupMember -Identity """ A_LoopField """ -Members """ logonName """`n"
return


Create:
GuiControl, Disable, createButton
Gui, Submit, NoHide

if ADUserExist(logonName)
{
	MsgBox, 16, , % "'" logonName "' already exists."
	GuiControl, Focus, logonName
	return
}
if InStr(logonName, A_Space) || RegExMatch(logonName, "[^a-zA-Z0-9\s]")
{
	MsgBox, 16, , % "Invalid username."
	GuiControl, Focus, logonName
	return
}
password := (password?password:"Aa123456")
if !Password_Complexity(password)
{
	MsgBox, 16, , % "Password does not meet complexity requirements."
	GuiControl, Focus, password
	return
}

SB_SetText("Start creating user...")
Sleep 100

SB_SetText("Deletes files...")
powershellFilePath := A_Temp "\CreateADUserTempFile.ps1"
powershellLogFilePath := A_Temp "\CreateADUserLogFile.log"
IfExist, % powershellFilePath
	FileDelete, % powershellFilePath
IfExist, % powershellLogFilePath
	FileDelete, powershellLogFilePath
Sleep 50

SB_SetText("Creates files...")
powershellCommand := "&{`n" finalCommand "`n} *> " powershellLogFilePath
FileAppend, % powershellCommand, % powershellFilePath, UTF-8

SB_SetText("Running command...")
RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %powershellFilePath%,, Hide
Sleep 50

SB_SetText("Deletes files...")
FileDelete, % powershellFilePath

SB_SetText("Verifying data...")
adUserDetails := ""
adUserDetails := ADUserExist(logonName)

if !adUserDetails
{
	errMsgStr := "Error creating user."
	gosub, MsgBoxError
	return
}

if updateADUserDetailsCommand
{
	if ( (firstNameHeb) && (firstNameHeb != adUserDetails.givenName) ) || ( (lastNameHeb) && (lastNameHeb != adUserDetails.sn) ) || ( ((displayName) && (displayName != lastNameEng " " firstNameEng)) && (displayName != adUserDetails.DisplayName) )
	{
		errMsgStr := "User created successfully but there was a problem updating user information."
		gosub, MsgBoxError
		return
	}
}

if selectedGroupList
{
	;Insert the group names into a new array - for future development
	adUserGroupNames := []
	for i, v in adUserDetails.memberOf
		adUserGroupNames.Push( SubStr( StrSplit(i,",")[1] , 4) )
	
	;Convert the group names to a string
	for i, v in adUserGroupNames
		adUserGroupNamesStr .= "|" v "|"
	
	;Check if the user is a member of the selected groups
	Loop, Parse, selectedGroupList, |
	{
		if !InStr(adUserGroupNamesStr, "|" A_LoopField "|")
		{
			errMsgStr := "User created successfully but failed to add it to group."
			gosub, MsgBoxError
			return
		}
	}
}



SB_SetText("User '" logonName "' was successfully created.")

Loop, Parse, editBoxes, |
	GuiControl,, % A_LoopField
Loop, Parse, ddlAndCombo, |
	GuiControl, Choose, % A_LoopField, 0
LV_Modify("","-Check")
return


MsgBoxError:
MsgBox, 8212, ADUserCreator, % errMsgStr "`nDo you want to view the log file?"
IfMsgBox, Yes
	Run, % powershellLogFilePath
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

ADUserExist(_Item)
{
	Enabled := ComObjError() , ComObjError(0)
	
	objRootDSE := ComObjGet("LDAP://rootDSE")
	strDomain := objRootDSE.Get("defaultNamingContext")
	strADPath := "LDAP://" . strDomain
	objDomain := ComObjGet(strADPath)
	objConnection := ComObjCreate("ADODB.Connection")
	objConnection.Open("Provider=ADsDSOObject")
	objCommand := ComObjCreate("ADODB.Command")
	objCommand.ActiveConnection := objConnection
	
	objCommand.CommandText := "<" . strADPath . ">;(|(sAMAccountName=" . _Item . "));distinguishedName;subtree"
	;~ objCommand.CommandText := "SELECT * FROM 'LDAP://" strDomain "' WHERE sAMAccountName='" _Item "'"
	objRecordSet := objCommand.Execute
	objRecordCount := objRecordSet.RecordCount

	While !objRecordSet.EOF
	{
		FullDNUserName := objRecordSet.Fields.Item("distinguishedName").value
		objRecordSet.MoveNext
	}
	
	if objRecordCount
	{
		objUser := ComObjGet("LDAP://" FullDNUserName)
		objUserDetails := {}
		objUserDetails.distinguishedName 	:= objUser.Get("distinguishedName")
		objUserDetails.Name					:= objUser.Get("Name") ;Name at creation
		objUserDetails.givenName 			:= objUser.Get("givenName")
		objUserDetails.sn 					:= objUser.Get("sn")   
		objUserDetails.DisplayName 			:= objUser.Get("DisplayName")
		objUserDetails.Description 			:= objUser.Get("Description")
		objUserDetails.TelephoneNumber 		:= objUser.Get("TelephoneNumber")
		objUserDetails.Department 			:= objUser.Get("Department")
		objUserDetails.Mail 				:= objUser.Get("Mail")
		objUserDetails.scriptPath 			:= objUser.Get("scriptPath")
		objUserDetails.memberOf				:= objUser.Get("memberOf") ;return array
	}
	else
	{
		objUserDetails := false
	}
	
	objRelease(objRootDSE)
	objRelease(objDomain)
	objRelease(objConnection)
	objRelease(objCommand)
	
	ComObjError(Enabled)
	
	return objUserDetails
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
