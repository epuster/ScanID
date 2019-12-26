;	Written by Eric Puster, epuster@gmail.com

Opt("MustDeclareVars", 1)
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <GuiTab.au3>
#include <IE.au3>
#include <Array.au3>
#include <ColorConstants.au3>
#include <WindowsConstants.au3>
#include <Crypt.au3>


Local		$sIDList = ""			; A list of all ID numbers scanned
Local		$sNameList = ""		; A list of all conference attendees
Local		$sMissedNames			; A list of all names not present on New Innovations
Local		$sIdName	= ""			; A string containing the attendee name
Local		$bEnter					; Button to absorb {ENTER} keystrokes from scanner
Local		$bSend					; Button to submit list to New-Innovations
Local		$lSub						; Label to show that submission is in process
Local		$ttl = "ScanID Ver 0.3"
Local		$eId						; ID edit
Local		$hNameList				; GUI control to show current count
Local		$Id = 0					; ID scanned
Local		$stmp						; Temp string
Local		$hFile					; File handle if needed
Local		$x,$y						; Control coordinate anchor point
Global	$eMsg						; An error message
Global	$sNameFile = @ScriptDir & "\ScanID.lib"				; File in which to store database
Global	$sIniFile = @ScriptDir & "\ScanID.ini"					; INI file for program

;========================================================================================================
;
; Read .ini file
;
;========================================================================================================

Global $sContact = IniRead($sIniFile, "Global", "contact", "your Coordinator.")
Global $sloginform = IniRead($sIniFile, "Login page", "loginform", "")
Global $sinstitutionctl = IniRead($sIniFile, "Login page", "institutionctl", "")
Global $susernamectl = IniRead($sIniFile, "Login page", "usernamectl", "")
Global $spasswordctl = IniRead($sIniFile, "Login page", "passwordctl", "")
Global $sloginbtn = IniRead($sIniFile, "Login page", "loginbtn", "")

Global $sconflist = IniRead($sIniFile, "Conference List", "conflist", "http://www.google.com")
Global $sconfform = IniRead($sIniFile, "Conference List", "confform", "")
Global $sconftable = IniRead($sIniFile, "Conference List", "conftable", "")
Global $sconflink = IniRead($sIniFile, "Conference List", "conflink", "")

Global $srosterform = IniRead($sIniFile, "Roster", "rosterform", "")
Global $srostersizectl = IniRead($sIniFile, "Roster", "rostersizectl", "")
Global $srostersize = IniRead($sIniFile, "Roster", "rostersize", "")
Global $srostertable = IniRead($sIniFile, "Roster", "rostertable", "")
Global $srostercheckstart = IniRead($sIniFile, "Roster", "rostercheckstart", "")
Global $srostercheckend = IniRead($sIniFile, "Roster", "rostercheckend", "")
Global $srostersave = IniRead($sIniFile, "Roster", "rostersave", "")
Global $srosterpopuptitle = IniRead($sIniFile, "Roster", "rosterpopuptitle", "")

Global $semailweb = IniRead($sIniFile, "Email", "emailweb", "http://www.google.com")
Global $semailform = IniRead($sIniFile, "Email", "emailform", "")
Global $semailto = IniRead($sIniFile, "Email", "emailto", "")
Global $semailsubject = IniRead($sIniFile, "Email", "emailsubject", "")
Global $semailbody = IniRead($sIniFile, "Email", "emailbody", "")
Global $semailsend = IniRead($sIniFile, "Email", "emailsend", "")

; Cleanup/sort database file, consider holding open until done TODO

GUICreate($ttl,450,560)
$x = 12
$y = 10
GUICtrlCreateLabel("Scan badge barcodes and submit the list." & @CRLF & "If you entered your name incorrectly, contact " & $sContact,$x+20,$y+3)
$y += 40
GUICtrlCreateLabel("ID:",$x,$y+3)
$eID = GUICtrlCreateEdit("",$x+20,$y,70,22, $ES_NUMBER)
$bEnter = GUICtrlCreateButton("Enter",$x+115,$y-3,75,30,$BS_DEFPUSHBUTTON)
GUICtrlCreateLabel("IDs scanned:",$x,$y+40)
$hNameList = GUICtrlCreateEdit("",$x,$y+70,425,390,BitOr($GUI_SS_DEFAULT_EDIT, $ES_READONLY, $WS_VSCROLL))
$bSend = GUICtrlCreateButton("Send Conference Attendance",$x,520,425,30)
$lSub = GUICtrlCreateLabel("",$x+200,$y+3,150,30)
GUICtrlSetState($eID,$GUI_FOCUS)
GUISetState(@SW_SHOW)							; All set up, now display main window

;========================================================================================================
;
; Main Message Loop
;
;========================================================================================================

While True
	Switch GUIGetMsg()								; Did user do anything yet?
	Case $GUI_EVENT_CLOSE							; Red X checked
		If MsgBox(BitOR($MB_SYSTEMMODAL,$MB_YESNO,$MB_DEFBUTTON2,$MB_TOPMOST),$ttl,"Are you sure you wish to quit?") = 6 Then ExitLoop
	Case $bSend											; Send button pressed
		If $sIDList = "" Then
			MsgBox($MB_SYSTEMMODAL,$ttl,"Scan an ID first.")
		Else
			GUICtrlSetData($lSub,"Submitting, please wait . . .")
			GUISetState(@SW_DISABLE)				; Make window inactive while running _WebSend
			$sMissedNames = _WebSend($sNameList)		; Submit names to website
			If $eMsg Then MsgBox($MB_SYSTEMMODAL,$ttl,"Error in name submission:" & @CRLF & $eMsg)
			If $sMissedNames Then

				$hFile = FileOpen(@MyDocumentsDir & "\Missed.txt",$FO_OVERWRITE)
				If @error Then
					$sTmp = "Some names not entered in New Innovations." & @CRLF
					$sTmp &= "These names could not be saved because" & @CRLF
					$sTmp &= "My Documents\Missed.txt could not be opened." & @CRLF
					$sTmp &= "Usually this means the file is in use" & @CRLF
					$sTmp &= "by another program or you do not have access."
				Else
					FileWrite($hFile,@MON & "-" & @MDAY & "-" & @YEAR & " " & @HOUR & ":" & @MIN & @CRLF)
					FileWrite($hFile,$sMissedNames)
					FileClose($hFile)
					$sTmp = "Some names not entered in New Innovations." & @CRLF
					$sTmp &= "These names have been saved in a file called: " & @CRLF
					$sTmp &= "Missed.txt in the Documents folder."
				EndIf
			Else
				$sTmp = "All names submitted successfully." & @CRLF
				$sTmp &= "You may close the scanner or add additional names and resubmit."
			EndIf
			MsgBox($MB_SYSTEMMODAL,$ttl,$sTmp)
			GUISetState(@SW_ENABLE)
			GUICtrlSetData($lSub,"")
		EndIf
		GUICtrlSetState($eID,$GUI_FOCUS)			; Set the focus back to the ID field
	Case $eId											; Something typed in Id
		$Id = GUICtrlRead($eID)						; Grab the text that was entered
		Select
		Case StringLen($Id)<>10						; ID should be 10 characters (digits enforced by control style)
			GUICtrlSetBkColor($eID, $COLOR_RED)	; Flash Red
			Sleep(200)
			GUICtrlSetBkColor($eID, $COLOR_WHITE)
		Case StringInStr($sIDList,$Id)			; Check to see if already scanned
			GUICtrlSetBkColor($eID, $COLOR_RED)	; Flash Red
			Sleep(200)
			GUICtrlSetBkColor($eID, $COLOR_WHITE)
		Case Else										; New Attendee badge scanned
			GUISetState(@SW_DISABLE)				; Make window inactive while running _NameWrite
			$sIdName = _NameWrite($Id)				; Cross-reference the ID number, or add name if new
			If $sIdName = "" Then
				MsgBox($MB_SYSTEMMODAL, $ttl, "Error in name entry:" & @CRLF & $eMsg)
			Else
				$sNameList &= $sIdName & @CRLF	; Add name to list
				$sIDList &= $Id & ","				; Add ID number to list
				GUICtrlSetData($hNameList, $sNameList)
			EndIf
			GUISetState(@SW_ENABLE)
		EndSelect
		GUICtrlSetData($eID,"")						; Erase the number from the input box
		GUICtrlSetState($eId,$GUI_FOCUS) 		; Place focus back on input box
	EndSwitch
Wend

;========================================================================================================
;
; _NameWrite - Checks to see if ID is in database, if so, returns the name.  If not, asks for the name,
;						enters it in the database, and returns the name.
;
; Return Value - Success - A string in the format "LastName, Firstname"
;					  Failure - Empty string and set $eMsg with some useful information
;
;========================================================================================================

Func _NameWrite($Id)
	Local $sLine												; A line from the file
	Local $sColumn												; Columns from the line
	Local $hAsk_GUI											; GUI to ask for name
	Local $hFirstInput, $hLastInput						; Fields for first and last name
	Local $sFirstName, $sLastName							; Strings for first and last name
	Local $sName = ""											; Name in format "Lastname, Firstname"
	Local $iHash												; Hash of ID number
	Local $hSubmit												; Submit Button

	$eMsg = ""													; Reset Error message
	Local $hFile = FileOpen($sNameFile,$FO_Read)		; Open the name database
	If @error Then
		$eMsg = "Access denied to database file."
		$eMsg &= @CRLF & "Usually this means the file is in use"
		$eMsg &= @CRLF & "by another program or you do not have access."
	EndIf

	$iHash = _Crypt_HashData($Id + 53000,$CALG_SHA1)	; Scramble by SHA1, obscure ID

	; Check for the ID number in the file, grab the name if present
	$sLine = FileReadLine($hFile)							; Read the first data line
	While @error = False And $eMsg = "" And $sName = ""	; While not at EOF, no errors and no match found
		$sColumn = StringSplit($sLine,@TAB)				; Split the line at tabs
		If $iHash = $sColumn[1] Then						; If ID number matches
			FileClose($hFile)									; Close the database file
			$sName = $sColumn[2]								; Match found, set name
		EndIf
		$sLine = FileReadLine($hFile)						; Read a data line
	WEnd
	FileClose($hFile)

	If $sName = "" And $eMsg = "" Then	; If ID number not present and no error, then ask for and enter name
		$hAsk_GUI = GUICreate($ttl, 330, 140, -1, -1, BitOR($MB_SYSTEMMODAL,$MB_TOPMOST))
		GUICtrlCreateLabel("Please enter your name exactly as it appears in New Innovations.", 10, 10)
		GUICtrlCreateLabel("First:", 10, 30)
		GUICtrlCreateLabel("Last:", 175, 30)
		$hFirstInput = GUICtrlCreateEdit("", 10, 50, 145, 25, 0)
		$hLastInput = GUICtrlCreateEdit("", 175, 50, 145, 25, 0)
		$hSubmit = GUICtrlCreateButton("Submit", 10, 80, 310, 25, $BS_DEFPUSHBUTTON)
		GUISetState(@SW_SHOW)

		$hFile = FileOpen($sNameFile,$FO_Append)				; Open the name file to add the name
		If @error Then
			$eMsg = "Access denied to database file."
			$eMsg &= @CRLF & "Usually this means the file is in use"
			$eMsg &= @CRLF & "by another program or you do not have access."
		EndIf

		While $eMsg = ""												; As long as there is no error
			Switch GUIGetMsg()										; Did user do anything yet?
			Case $GUI_EVENT_CLOSE									; Red X checked
				$eMsg = "Name discarded."
				GUIDelete($hAsk_GUI)
				FileClose($hFile)
			Case $hSubmit												; Submit button pressed
				$sFirstName = GUICtrlRead($hFirstInput)		; Capture first name
				$sLastName = GUICtrlRead($hLastInput)			; Capture last name

				Select
				Case $sFirstName = ""								; Check if first name blank
					$eMsg = "First name is blank.  Please enter a first name."
				Case StringRegExp($sFirstName,"[^a-zA-Z\'\x20\-]")	; Allow only a-Z, -, ', and space
					$eMsg = "First name is invalid.  Please reenter."
					GUICtrlSetData($hFirstInput,"")
					GUICtrlSetState($hFirstInput,$GUI_FOCUS)
				Case $sLastName = ""									; Check if last name blank
					$eMsg = "Last name is blank.  Please enter a last name."
				Case StringRegExp($sLastName,"[^a-zA-Z\'\x20\-]")	; Allow only a-Z, -, ', and space
					$eMsg = "Last name is invalid.  Please reenter."
					GUICtrlSetData($hLastInput,"")
					GUICtrlSetState($hLastInput,$GUI_FOCUS)
				Case Else												; First and last name are present and alpha
					FileWrite($hFile,$iHash & @TAB & $sLastName & ", " & $sFirstName & @CRLF)
					$eMsg = "The name " & $sLastName & ", " & $sFirstName & " was successfully written."
					GUIDelete($hAsk_GUI)
					FileClose($hFile)
					$sName  = $sLastName & ", " & $sFirstName
				EndSelect

				If $sName = "" Then									; If we don't have a name
					MsgBox($MB_SYSTEMMODAL,$ttl,$eMsg)			; Show the error message
					$eMsg = ""											; Clear error message
				EndIf
			EndSwitch
		Wend
	EndIf

	Return $sName														; Return matching name, or "" if none
EndFunc

;===========================================================================================================
;
; _WebSend(String ListofNames, Local Number of names) - A function to submit the collected
;		attendance data to New Innovations
;
; Returns a list of all names not matched.  If error occurs, useful information stored in $eMsg
;
;===========================================================================================================

Func _WebSend($sNameList)

	const $sNI = "https://www.new-innov.com/Login/Login.aspx"
	Const $sNow = @HOUR & @MIN & @SEC								; Current time in format HHMMSS
;	Local $sNowDate = 20180219123000														; TESTING

	Const $sNowDate = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC		; Time to compare against
	Const $sTime = @HOUR & ":" & @MIN & ":" & @SEC				; Current time in format HH:MM:SS
	Const $sInstitutionname = "fakeinstname"						; REPLACE fakeinstname with NI INSTITUTION NAME
	const $sUsername = "fakeusername"								; REPLACE fakeusername with NI USERNAME
	const $sPassword = "fakepassword"								; REPLACE fakepassword with NI PASSWORD
	Local $sConfName,$sConfTime					; Conference info
	Local $sConfLinks									; List of links to conference rosters
	Local $oForm, $oText1, $oBox					; Temporary values for finding controls in a webpage
	Local $oText2, $oText3, $oText4
	Local $sTable										; Table of strings from web page
	Local $sLine, $sLine2							; Line of text in a cell of a table
	Local $sDate										; A string containing a date
	Local $sNameArray									; Array of names to be matched
	Local $sMissedNames								; A list of names not found
	Local $sNamesMarked = ""						; Names recorded successfully
	Local $hBrowser									; Handle for the browser window
	Local $i, $j										; Counters
	Local $iStatus = 0								; Status of routine: 0 = Nominal, 1 = Login failed, 2 = Attendance failed

	$sNameArray = StringSplit($sNameList,@CRLF,3)						; Place names in an array
	_ArraySort($sNameArray)														; Sort names alphabetically with Quicksort
	_ArrayDelete($sNameArray,0)												; Delete the blank entry
	_IEErrorHandlerRegister()													; Prevent IE errors from crashing script

	;-----------------------------
	;  Login Screen
	;-----------------------------

	$hBrowser = _IECreate($sNI,0,0)
	;$hBrowser = _IECreate($sNI)												; Create window and go to NI (visible for TESTING)
	_IELoadWait($hBrowser)														; Wait for window to fully load

	$oForm = _IEFormGetObjByName($hBrowser, $sloginform)				; Grab the form

	If @error Then																	; Go no further if form not found
		$eMsg = "Could not find expected web form on Login page"
	Else
		$oText1 = _IEFormElementGetObjByName($oForm, $sinstitutionctl)	; Find the field for institution
		If @error Then $eMsg &= "Could not find institution field on Login page" & @CRLF
		$oText2 = _IEFormElementGetObjByName($oForm, $susernamectl)	; Find the field for Username
		If @error Then	$eMsg &= "Could not find username field on Login page" & @CRLF
		$oText3 = _IEFormElementGetObjByName($oForm, $spasswordctl)	; Find the field for Password
		If @error Then	$eMsg &= "Could not find password field on Login page" & @CRLF
		$oText4 = _IEFormElementGetObjByName($oForm, $sloginbtn)		; Find the Login button
		If @error Then $eMsg &= "Could not find Login button on Login page" & @CRLF
	EndIf

	If $eMsg <> "" Then														; If there was an error
		$sMissedNames = $sNameList											; Mark all names as missed
		$iStatus = 1														; Set status to "login failed"
	Else
		_IEFormElementSetValue($oText1, $sInstitutionname)					; Write institution name for login
		_IEFormElementSetValue($oText2, $sUsername)							; Write username to web form
		_IEFormElementSetValue($oText3, $sPassword)							; Write password to web form
		_IEAction($oText4,"click")											; Click the Login button
		_IELoadWait($hBrowser)
	EndIf

	;-----------------------------
	;  Conference Listing Screen
	;-----------------------------
	If $iStatus = 0 Then													; If error in login, skip this section
		_IENavigate($hBrowser, $sconflist)									; Navigate to listing page
		_IELoadWait($hBrowser)												; Wait for it to load

		$oForm = _IEFormGetObjByName($hBrowser, $sconfform)					; Grab the form
		If @error Then														; If error, go no further
			$eMsg = "Could not find expected web form on conference listing page"
			$eMsg &= @CRLF & "This could also be due to the web address being incorrect."
		Else
			$oText1 = _IEGetObjById($hBrowser,$sconftable)					; Grab the table
			If @error Then													; Stop if table not present
				$eMsg = "Could not find expected table of conferences"
			Else
			; Step through the table one field at a time, locate the time
			; Column 0 = Repeat flag, 1 = Name of conference (link), 2 = Date and time of conference
			; Consider improving to search for date and time and match rather than just taking the last
			; Place links in an array $sLink[], correlate position in table
				$sTable = _IETableWriteToArray($oText1)						; Grab the text from the table

				For $i = 1 to UBound($sTable,2) - 2							; Cycle through all rows, last two rows hold nothing
					$sLine = StringSplit($sTable[2][$i]," ")				; Split whole time field on spaces
					$sLine2 = StringSplit($sLine[1],"/")					; Split date on slashes
					$sLine2[1] = StringRight("0" & $sLine2[1],2)			; If month < 10 prepend 0
					$sLine2[2] = StringRight("0" & $sLine2[2],2)			; If day < 10 prepend 0
					$sConfTime = $sLine2[3] & $sLine2[1] & $sLine2[2]		; Write to $sConfTime in format YYYYMMDD
					$sLine2 = StringSplit($sLine[2],":")					; Split time on colons
					If $sLine2[1] = 12 Then $sLine2[1] = 0					; Set all 12's to 0
					If $sLine[3] == "PM" Then $sLine2[1] += 12				; Convert to military time
					$sLine2[1] = StringRight("0" & $sLine2[1],2)			; If Hour < 10 prepend 0
					$sConfTime &= $sLine2[1] & $sLine2[2] & "00"			; Add time in HHMMSS format to conference time

					; Check for current time within two hours of start time of a meeting
					Select
					Case $sNowDate < $sConfTime								; If the meeting has not happened yet, skip
						Case $sNowDate - $sConfTime < 20000					; If we are less than two hours after, choose

						$sConfName = $sTable[1][$i]							; If found, grab the name
						$j = $i												; And the index in the table
						$sConfName = StringRegExpReplace($sConfName,"^\*\x20","")	; If "* " at beginning, remove it
					EndSelect
				Next

				If $sConfName = "" Then										; If no match was found
					$eMsg = "No conference found within two hours."
					$sConfTime = ""											; Clear Conference Time as well
				Else														; If a match was found
					$sLine = StringSplit($sTable[2][$j]," ")				; Grab the time of the conference again
					$sConfTime = $sLine[1]
					$sLine = ""												; Reinitialize $sLine

					$sTable = _IELinkGetCollection($oForm)					; Grab all the links on the page
					; Collect a list of the links in order, grab only those which involve attendance taking
					For $sLink in $sTable
						If StringLeft($sLink.href, 57) == $sconflink Then $sLine &= $sLink.href & @TAB
					Next
					$sConfLinks = StringSplit($sLine,@TAB)					; Place the links into an array
					If @error Then											; If no links were found
						$eMsg &= "No links to conferences found in table."  & @CRLF
						$eMsg &= "This could be from a bad table ID, or else a conflink" & @CRLF
						$eMsg &= "string that is outdated in the .ini file."
					EndIf
				EndIf
			EndIf
		EndIf

		If $eMsg Then														; If there was an error
			$sMissedNames = $sNameList										; Mark all names as missed
			$iStatus = 2													; Set status to "attendance failed"
		EndIf
	EndIf

	;-----------------------------
	;  Roster Screen
	;-----------------------------

	If $iStatus = 0 Then													; Only move forward if nominal
		; Go to the link that lay in the same row with the correct time
		_IENavigate($hBrowser, $sConfLinks[$j])
		_IELoadWait($hBrowser)

		$oText1 = _IEGetObjById($hBrowser,$srostersizectl)					; Grab the page size tool
		If @error Then	$eMsg = "Could not find the table size drop down menu" & @CRLF
		_IEFormElementSetValue($oText1,$srostersize)						; Set page size to hold maximum number of names
		_IELoadWait($hBrowser)												; Wait for page to reload

		$oForm = _IEFormGetObjByName($hBrowser, $srosterform)				; Grab the form
		If @error Then														; If form not present, go no further
			$eMsg &= "Could not find expected web form on roster page"
		Else
			$oText3 = _IEGetObjById($hBrowser,$srostersave)					; Grab the "Save" link by name
			If @error Then $eMsg &= "Could not find Save button on roster page." & @CRLF
			$oText4 = _IEGetObjById($oForm,$srostertable)					; Grab the table
			If @error Then	$eMsg &= "Could not find expected table of attendees" & @CRLF

			If $eMsg = "" Then														; If no errors, go ahead
				$sTable = _IETableWriteToArray($oText4)						; Convert table to array
				$sLine = ""																; Reset for use as row number in control name
				If @error Then MsgBox($MB_SYSTEMMODAL,"","_IETableWriteToArray error " & @error)
				For $i = 0 to UBound($sNameArray) - 1							; Cycle through each name on the list of scanned badges
					For $j = 1 to UBound($sTable,2) - 1							; For each of those names, check it against the roster
						If $sNameArray[$i] = $sTable[3][$j] Then				; If we find the badge scanned in the roster
							$sLine = $j + 1											; The given number for the control on the webpage is row + 1
							If $sLine < 10 Then										; add a zero in front if needed
							Else
								$sLine = $j + 1
							EndIf

							$oBox = _IEGetObjById($oText4,$srostercheckstart & $sLine & $srostercheckend)	; Grab the checkbox by name
							If @error And $eMsg = "" Then							; If this is the first missed checkbox
								$eMsg = "At least one attendee listed "		; Write an error message
								$eMsg &= "with no ""Present"" checkbox found in roster table." & @CRLF
								$eMsg &= " This may be because the page was not fully loaded."
								$sMissedNames &= $sNameArray[$i] & @CRLF		; Add the name to the list of the missed
								SetError(0)												; Reset error
							ElseIf @error Then										; If not the first miss . . .
								$sMissedNames &= $sNameArray[$i] & @CRLF		; Add the name to the list of the missed
								SetError(0)												; Reset error
							Else
								$oBox.checked = True									; Mark the "Present" checkbox
								$sNamesMarked &= $sNameArray[$i] & @CRLF		; Add to the list of recorded names
							EndIf
						EndIf
					Next
					If $sLine = "" Then												; If name was not found (no row number set)
						$sMissedNames &= $sNameArray[$i] & @CRLF				; Add the name to the list of the missed
					Else
						$sLine = ""														; Reset row number if used
					EndIf
				Next

				If $eMsg <> ""	Then $iStatus = 2									; Note that a matched name missed

				; Klugey bug fix, somehow the list of names gets double line feeds
				$sMissedNames = StringReplace($sMissedNames,@CRLF & @CRLF, @CRLF)

				$sLine = _IEPropertyGet($hBrowser, "hwnd")					; Load the name of the webpage
				_IEAction($oText3, "focus")										; Focus on the Save Link
				ControlSend($sLine, "", "", "{Enter}")							; Execute a click on the Save Link (_IEAction() causes script to freeze)
				WinWait($srosterpopuptitle, "")									; Wait for popup to appear, then click on the OK button
				ControlClick($srosterpopuptitle, "", "[CLASS:Button; TEXT:OK; Instance:1;]")

				_IELoadWait($hBrowser)												; Wait for roster to be saved
				Sleep(2000)																; Wait an extra second
			EndIf
		EndIf

		If $eMsg And $sMissedNames = "" Then									; If there was an error and no missed names recorded
			$sMissedNames = $sNameList												; Mark all names as missed
			$iStatus = 2																; Set status to "attendance failed"
		EndIf

	EndIf

	;-----------------------------
	;  Email Screen
	;-----------------------------

	If $iStatus <> 1 Then															; Go ahead as long as login worked
		SetError(0)																		; Reset error flag
		$sLine = ""																		; Reset $sLine

		_IENavigate($hBrowser,$semailweb)										; Go to email page
		_IELoadWait($hBrowser)														; Wait for E-mail page to load
		$oForm = _IEFormGetObjByName($hBrowser,$semailform)				; Grab the form
		If @error Then																	; If form not present go no further
			$eMsg &= "Could not find expected web form on the Email page" & @CRLF
		Else
			$oText1 = _IEGetObjById($hBrowser,$semailto)						; Grab the To: field
			If @error Then $sLine &= "Could not find To: field on the Email page" & @CRLF
			$oText2 = _IEGetObjById($hBrowser,$semailsubject)				; Grab the Subject field
			If @error Then	$sLine &= "Could not find Subject field on the Email page" & @CRLF
			$oText3 = _IEGetObjById($hBrowser,$semailbody)					; Grab the Body section of email
			If @error Then	$sLine &= "Could not find Body field on the Email page" & @CRLF
			$oText4 = _IEGetObjById($hBrowser,$semailsend)					; Grab the Send button
			If @error Then	$sLine &= "Could not find Send button on the Email page" & @CRLF

			If $sLine = "" Then														; If controls loaded
				; Set To: field to /dev/null, a confirmation e-mail is already sent to owner of account
				_IEFormElementSetValue($oText1,"donotreply@ascension.org")
				_IEFormElementSetValue($oText2,"Attendance " & $sConfTime)	; Set subject, add Conference Time if found

				; Assemble body of e-mail
				If $sConfName Then													; If we found a name for the meeting
					$sLine2 = $sConfName												; Use it
				Else
					$sLine2 = "a didactic session"								; Otherwise set a generic name
				EndIf

				$sLine = "Residency Coordinator," & @CRLF
				$sLine &= "    " & $sLine2 & " took place " & $sConfTime & @CRLF
				If $sMissedNames <> "" Then
					$sLine &= "    Please find below a list of missed entries:"
					$sLine &= @CRLF & @CRLF & $sMissedNames & @CRLF
				EndIf
				If $sNamesMarked <> "" Then
					$sLine &= "    The following names were entered for attendance:" & @CRLF
					$sLine &= $sNamesMarked & @CRLF
				EndIf
				If $eMsg <> "" Then
					$sLine &= "    The following error was also produced:" & @CRLF
					$sLine &= $eMsg & @CRLF
				EndIf
				$sLine &= "Sincerely," & @CRLF & @TAB & $ttl

				_IEFormElementSetValue($oText3,$sLine)							; Insert message in Body field
				$sLine2 = _IEPropertyGet($hBrowser, "hwnd")					; Load the name of the webpage
				_IEAction($oText4, "focus")										; Focus on the Save Link
				ControlSend($sLine2, "", "", "{Enter}")						; Execute a click on the Save Link (_IEAction() causes script to freeze)
				_IELoadWait($hBrowser)
				Sleep(20000)															; Wait for "Sending" tooltip to finish
			Else
				$eMsg &= $sLine														; If controls didn't load, update error
			EndIf
		EndIf

		If $sLine <> "" Or @error Then
		$eMsg &= "In this case, attendance was recorded, but" & @CRLF
		$eMsg &= "no confirmatory e-mail could be sent." & @CRLF & @CRLF
		$eMsg &= "These names were not found in New-Innovations:" & @CRLF
		$eMsg &= $sMissedNames
		EndIf

	EndIf

	_IEQuit($hBrowser)																; Close the browser

	Return $sMissedNames
EndFunc