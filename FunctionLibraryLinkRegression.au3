#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         Joshua Walton

 Script Function:
	Function Library for Link Regression

#ce ----------------------------------------------------------------------------
#include <File.au3>
#include <Array.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>

Global $windowtitleinuse = "Remote Link"
Global $aOperatorUsernames[604]
Global $aUserLastNames[1000]
Global $aUserFirstNames[1000]
Global $LogFilePath = ""
Global $MessageToWrite = ""
Global $FilePath = "P:\JW\Auto It Scripts\Remote Link\Link Automatic Regression\Remote Link Automatic Regression\TextFiles"

Func DefaultOperatorLinkLogin() ;Login into Remote Link using default login information
	Local $filepathstring = "C:\Link\Link.exe"
	Local $windowtitleinuse = "Remote Link"
	Local $operatorusername = "new"
	Local $operatorpassword = "new"

	Run($filepathstring)
	Sleep(5000)
	WinWaitActive($windowtitleinuse, "", 30)

	If WinActive($windowtitleinuse) Then
		ControlSetText($windowtitleinuse, "", "TEdit2", "")
		Sleep(250)
		ControlSend($windowtitleinuse, '', "TEdit2", $operatorusername)
		Sleep(250)
		ControlSend($windowtitleinuse, '', "TEdit1", $operatorpassword)
		Sleep(250)
		ControlClick($windowtitleinuse, "OK", "TButton2")
	EndIf

	$windowtitleinuse = "Remote Link"
	WinWaitActive($windowtitleinuse, "Panel Information", 5)
	If WinActive($windowtitleinuse, "Panel Information") Then
		$MessageToWrite = "Operator " & $operatorusername & " logged in with a password of " & $operatorpassword
		LogFileWrite($LogFile, $MessageToWrite)
	EndIf
EndFunc
Func CustomOperatorLinkLogin($operatorusername, $iterationpassword) ;Login into Remote Link using custom login information
	Local $filepathstring = "C:\Link\Link.exe"
	Local $windowtitleinuse = "Remote Link"

	Run($filepathstring)
	Sleep(5000)
	WinWaitActive($windowtitleinuse, "", 30)

	If WinActive($windowtitleinuse) Then
		ControlSetText($windowtitleinuse, "", "TEdit2", "")
		Sleep(250)
		ControlSend($windowtitleinuse, '', "TEdit2", $operatorusername)
		Sleep(250)
		ControlSend($windowtitleinuse, '', "TEdit1", $iterationpassword)
		Sleep(250)
		ControlClick($windowtitleinuse, "OK", "TButton2")
	EndIf

	$windowtitleinuse = "Remote Link"
	WinWaitActive($windowtitleinuse, "Panel Information", 5)
	If WinActive($windowtitleinuse, "Panel Information") Then
		$MessageToWrite = "Operator " & $operatorusername & " logged in with a password of " & $iterationpassword
		LogFileWrite($LogFile, $MessageToWrite)
	Else
		MsgBox(4112, "Failure!", "Operator " & $operatorusername & " could not be logged in with his password.  Please check for control characters.")
		$MessageToWrite = "Operator " & $operatorusername & " could not be logged in with a password of " & $iterationpassword
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf
EndFunc
Func CreateListUsernames(ByRef $aOperatorUsernames) ;Creates a List of Valid User Names in Array
	Local $aUsernameList[604]
	Local $file = FileOpen("P:\JW\Auto It Scripts\Remote Link\Link Automatic Regression\Database Text Files\ListofUsernames.txt", 0)
	Local $i = 0

	If $file = -1 Then
		MsgBox(0, "Error", "Unable to open file.")
		Exit
	EndIf

	While $i < 604
		Local $line = FileReadLine($file)
		If @error = -1 Then ExitLoop
		$aUsernameList[$i] = $line
		$aOperatorUsernames[$i] = $aUsernameList[$i]
		$i += 1
	WEnd

	FileClose($file)
	$aOperatorUsernames = $aUsernameList
	;MsgBox(4096, "Test", "Sample Username Values: " & $aOperatorUsernames[14] & " " & $aOperatorUsernames[122], 5)
EndFunc
Func CreateListLastNames(ByRef $aUserLastNames) ;Creates a List of Valid Last Names in Array
	Local $aUsernameList[1000]
	Local $file = FileOpen("P:\JW\Auto It Scripts\Remote Link\Link Automatic Regression\Database Text Files\1000LastNames.txt", 0)
	Local $i = 0

	If $file = -1 Then
		MsgBox(0, "Error", "Unable to open file.")
		Exit
	EndIf

	While $i < 1000
		Local $line = FileReadLine($file)
		If @error = -1 Then ExitLoop
		$aUsernameList[$i] = $line
		$aUserLastNames[$i] = $aUsernameList[$i]
		$i += 1
	WEnd

	FileClose($file)
	$aUserLastNames = $aUsernameList
	Return $aUserLastNames
	;MsgBox(4096, "Test", "Sample Last Name Values: " & $aUserLastNames[714] & " and " & $aUserLastNames[522], 5)
EndFunc
Func CreateListFirstNames(ByRef $aUserFirstNames) ;Creates a List of Valid First Names in Array
	Local $aUsernameList[1000]
	Local $file = FileOpen("P:\JW\Auto It Scripts\Remote Link\Link Automatic Regression\Database Text Files\1000FirstNamesFemale.txt", 0)
	Local $i = 0

	If $file = -1 Then
		MsgBox(0, "Error", "Unable to open file.")
		Exit
	EndIf

	While $i < 1000
		Local $line = FileReadLine($file)
		If @error = -1 Then ExitLoop
		$aUsernameList[$i] = $line
		$aUserFirstNames[$i] = $aUsernameList[$i]
		$i += 1
	WEnd

	FileClose($file)
	$aUserFirstNames = $aUsernameList
	Return $aUserFirstNames
	;MsgBox(4096, "Test", "Sample First Name Values: " & $aUserFirstNames[714] & " " & $aUserFirstNames[522], 5)
EndFunc
Func RandomCharPasswordGenerator($max_len) ;Borrowed random character password generator (see http://pastebin.com/TmUkBZRS)
   Local $aPassword[$max_len]

   For $i = 0 To UBound($aPassword) - 1
	   $aPassword[$i] = chr(Random(33, 126, 1))
   Next

   $aPassword = StringReplace(_ArrayToString($aPassword, @TAB, 0, $max_len), @TAB, "")
   Return $aPassword

EndFunc
Func LogFileWrite($LogFile,$MessageToWrite) ;Function to write to specified log file with specific message
	Local $location = $LogFile
	Local $message = $MessageToWrite
	Local $iFlag = 0

	Local $log = FileOpen(@ScriptDir & $location, 1)
	Local $result = _FileWriteLog($log, $message, $iFlag = -1)

	If $result = 0 Then
		If @error = 1 Then
			MsgBox(4096, "Failure!", "Could not open specified file at " & $location & @LF & "Please confirm location is valid.", 30)
			Exit
		EndIf
		If @error = 2 Then
			MsgBox(4096, "Failure!", "File located at " & $location & " could not be written to." & @LF & "Please confirm location is valid.", 30)
			Exit
		EndIf
	Else
		FileClose($location)
	EndIf
EndFunc
Func FindCurrentPanelFile(ByRef $panelaccountundertestnum, ByRef $panelaccountundertestname, ByRef $panelaccountundertestrec)	;Find Panel File to be used from Database
	$lastpanelaccnum = 1
	$lastpanelrecnum = 1
	ControlClick($windowtitleinuse, "Panel Information", 65280)	;ControlID for Win7 - TPanelInformationForm1
	Sleep(250)
	ControlClick($windowtitleinuse, "", "TwwDBGrid1")	;Control ID for Win7 - 263510
	Sleep(250)
	Do
		$selectedpanelfilerecnum = ControlGetText($windowtitleinuse, "Panel Information", 198218) ;ControlID for Win7 - TDBEdit13
		$selectedpanelfileacctnum = ControlGetText($windowtitleinuse, "Panel Information", 198158) ;ControlID for Win7 - TDBEdit14
		If $selectedpanelfilerecnum <> $panelaccountundertestrec Or $selectedpanelfileacctnum <> $panelaccountundertestnum Then
			If $selectedpanelfileacctnum = $lastpanelaccnum And $selectedpanelfilerecnum = $lastpanelrecnum Then
				ControlSend($windowtitleinuse, "Panel Information", "TwwDBGrid1", "{UP}", 0)
			Else
				ControlSend($windowtitleinuse, "Panel Information", "TwwDBGrid1", "{DOWN}", 0)
			EndIf
		EndIf
		$lastpanelaccnum = $selectedpanelfileacctnum
		$lastpanelrecnum = $selectedpanelfilerecnum
	Until $selectedpanelfileacctnum = $panelaccountundertestnum And $selectedpanelfilerecnum = $panelaccountundertestrec
	$selectedpanelfilename = ControlGetText($windowtitleinuse, "Panel Information", 198156) ;ControlID for Win7 - TDBEdit15
	If $selectedpanelfilename <> $panelaccountundertestname Then
		MsgBox(4096, "Error with Panel File!", "The Panel Name field contains unexpected information, this is not the panels we are looking for!", 10)
		Exit
	Else
		ControlClick($windowtitleinuse, "Panel Information", "&OK")	;ControlID for Win7 - TButton5 or 198192
		MsgBox(4096, "Panel File Found!", "Will Now Collect Panel Programming for Verification", 2)
	EndIf
EndFunc
Func EnterPanelProgrammingMenu()
	Sleep(500)
	Send("!r")	;Enter Panel Programming
	Sleep(250)
EndFunc
Func PanelProgrammingLog($sfilename,$programmingtowrite)
	$sfilelocation = "P:\JW\Auto It Scripts\Remote Link\Link Automatic Regression\Remote Link Automatic Regression\TextFiles"
	Local Const $sFilePath = $sfilelocation & $sfilename

	Local $hfileopen = FileOpen($sFilePath, 9)
	If $hfileopen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the file.")
		Exit
	EndIf

	Local $result = FileWriteLine($hfileopen, $programmingtowrite)

	If $result = 0 Then
		MsgBox(4096, "Failure!", "Could not write to specified file at " & $sfilename & @LF & "Please confirm text file exists.", 30)
		Exit
	Else
		FileClose($FilePath)
	EndIf
EndFunc
Func RetrieveXRCommunicationPathProgramming($sfilename, $programmingtowrite, ByRef $aXRCommPathXSettings)
	Local $aXRCommPathXSettings[31] = _
		[$xraccountnumber, $xrtransmitdelay, $xrpathnumber, $xrcommunicationtype, $xrcellularstatus, $xrpathtype, $xrtestreport, $xrtestfrequency, _
		$xrfrequencyunit, $xrtesttime, $xrusecheckin, $xrcheckinminutes, $xrfailtimeminutes, $xrfirstphonenumber, $xrsecondphonenumber, $xrencryption, _
		$xrencryptionschema, $xrreceiverip, $xrreceiverport, $xrsendcommtrouble, $xrsendpathinformation, $xrduplicatealarms, $xrfailtesthours, $xr893A, _
		$xrsecondlineprefix, $xralarmswitchover, $xrretryseconds, $xrsubstitutioncode, $xrprotocol, $xrfirstgprsapn, $xrsecondgprsapn]

	;Assuming EnterPanelProgrammingMenu Function has been initiated leaving you with Panel Prgramming drop down open
	Send("{ENTER}") 	;Communication Paths is selected
	Sleep(2000)
	Local $currentprogoptform = ControlGetText($windowtitleinuse, "", 65280)
	If $currentprogoptform <> "Communication Paths" Then
		MsgBox(4096, "Error Entering Panel Programming!", "Script is not in Communication Paths.", 10)
		Exit
	EndIf
	$sfilename = "\CommunicationPaths\BeforeSend\PathOneProgramming.txt"
	;Paths
	$xraccountnumber = ControlGetText($windowtitleinuse, "", "TDBNumericControl6") ;Account Number - No Text, Watch Control Class Instance
	$programmingtowrite = "Account Number: " & $xraccountnumber
	PanelProgrammingLog
	$xrtransmitdelay = ControlGetText($windowtitleinuse, "", "TDBNumericControl5") ;Transmit Delay - No Text, Watch Control Class Instance
	$programmingtowrite = "Transmit Delay: " & $xrtransmitdelay
	PanelProgrammingLog
	$xrpathnumber = ControlGetText($windowtitleinuse, "", "TDBNumericListControl1") ;Path Number - No Text, Watch Control Class Instance
	$programmingtowrite = "Path Number: " & $xrpathnumber
	PanelProgrammingLog
	$xrcommunicationtype = ControlGetText($windowtitleinuse, "", "TDBListControl15") ;Comm Type - No Text, Watch Control Class Instance
	$programmingtowrite = "Communication Type: " & $xrcommunicationtype
	PanelProgrammingLog
	If $xrcommunicationtype = "Cellular Network" Then
		$xrcellularstatus = ControlGetText($windowtitleinuse, "", "TDBCalculatedControl1") ;Cellular Status - No Text, Watch Control Class Instance
		$programmingtowrite = "Cellular Status: " & $xrcellularstatus
		PanelProgrammingLog
	Else
		$xrcellularstatus = ""
		$programmingtowrite = "Cellular status: " & $xrcellularstatus
		PanelProgrammingLog
	EndIf
	$xrpathtype = ControlGetText($windowtitleinuse, "", "TDBListControl14") ;Path Type - No Text, Watch Control Class Instance
	$programmingtowrite = "Path Type: " & $xrpathtype
	PanelProgrammingLog
	;Supervision
	$xrtestreport = ControlGetText($windowtitleinuse, "", "TDBListControl12") ;Test Report - No Text, Watch Control Class Instance
	$programmingtowrite = "Test Report: " & $xrtestreport
	PanelProgrammingLog
	$xrtestfrequency = ControlGetText($windowtitleinuse, "", "TDBNumericControl3") ;Test Frequency - No Text, Watch Control Class Instance
	$programmingtowrite = "Test Frequency: " & $xrtestfrequency
	PanelProgrammingLog
	$xrfrequencyunit = ControlGetText($windowtitleinuse, "", "TDBListControl11") ;Frequency Unit - No Text, Watch Control Class Instance
	$programmingtowrite = "Frequency Unit: " & $xrfrequencyunit
	PanelProgrammingLog
	$xrtesttime = ControlGetText($windowtitleinuse, "", "TDBTimeControl11") ;Test Time - No Text, Watch Control Class Instance
	$programmingtowrite = "Test Time: " & $xrtesttime
	PanelProgrammingLog
	;CheckIn
	If $xrcommunicationtype = "Network" Or "Cellular Network" Then
		$xrusecheckin = ControlGetText($windowtitleinuse, "", "TDBListControl9") ;Use Check-In - No Text, Watch Control Class Instance
		$programmingtowrite = "Check In: " & $xrusecheckin
		PanelProgrammingLog
		$xrcheckinminutes = ControlGetText($windowtitleinuse, "", "TDBNumericControl2") ;Check-in(Minutes) - No Text, Watch Control Class Instance
		$programmingtowrite = "Check In Minutes: " & $xrcheckinminutes
		PanelProgrammingLog
		$xrfailtimeminutes = ControlGetText($windowtitleinuse, "", "TDBNumericControl1") ;Fail Time(Minutes) - No Text, Watch Control Class Instance
		$programmingtowrite = "Fail Time Minutes: " & $xrfailtimeminutes
		PanelProgrammingLog
	Else
		$xrusecheckin = ""
		$programmingtowrite = "Check In: " & $xrusecheckin
		PanelProgrammingLog
		$xrcheckinminutes = ""
		$programmingtowrite = "Check In Minutes: " & $xrcheckinminutes
		PanelProgrammingLog
		$xrfailtimeminutes = ""
		$programmingtowrite = "Fail Time Minutes: " & $xrfailtimeminutes
		PanelProgrammingLog
	EndIf
	;Comm Type Details
	Switch $xrcommunicationtype
		Case "None"
			$programmingtowrite = "There are no Comm Type Details."
			PanelProgrammingLog
		Case "Digital Dialer"
			$xrfirstphonenumber = ControlGetText($windowtitleinuse, "", "TDBStringControl2") ;1st Phone No. - No Text, Watch Control Class Instance
			$programmingtowrite = "First Phone Number: " & $xrfirstphonenumber
			PanelProgrammingLog
			$xrsecondphonenumber = ControlGetText($windowtitleinuse, "", "TDBStringControl1") ;2nd Phone No. - No Text, Watch Control Class Instance
			$programmingtowrite = "Second Phone Number: " & $xrsecondphonenumber
			PanelProgrammingLog
		Case "Network"
			$xrencryptionenabled = ControlCommand($windowtitleinuse, "", "Encryption", "IsVisisble", "")
			If $xrencryptionenabled = 1 Then
				$xrencryption = ControlCommand($windowtitleinuse, "", "Encryption", "IsChecked", "") ;Encryption - ClassInstanceNN: TDBBooleanControl1
				If $xrencryption = 1 Then
					$programmingtowrite = "Encryption Enabled"
					PanelProgrammingLog
					$xrencryptionschema = ControlGetText($windowtitleinuse, "", "TDBListControl13") ;Encryption Schema - No Text, Watch Control Class Instance
					$programmingtowrite = "Encryption Scheme " & $xrencryptionschema
					PanelProgrammingLog
				Else
					$programmingtowrite = "Encryption Disabled"
					PanelProgrammingLog
					$xrencryptionschema = ""
					$programmingtowrite = "Encryption Scheme " & $xrencryptionschema
					PanelProgrammingLog
				EndIf
			Else
				$programmingtowrite = "No Encryption in this Panel"
				PanelProgrammingLog
			EndIf
			$xrreceiverip = ControlGetText($windowtitleinuse, "", "TDBStringControl3") ;Receiver IP - No Text, Watch Control Class Instance
			$programmingtowrite = "Receiver IP: " & $xrreceiverip
			PanelProgrammingLog
			$xrreceiverport = ControlGetText($windowtitleinuse, "", "TDBStringControl3") ;Receiver Port - No Text, Watch Control Class Instance
			$programmingtowrite = "Receiver Port: " & $xrreceiverport
			PanelProgrammingLog
		Case "Contact ID"
			$xrfirstphonenumber = ControlGetText($windowtitleinuse, "", "TDBStringControl2") ;1st Phone No. - No Text, Watch Control Class Instance
			$programmingtowrite = "First Phone Number: " & $xrfirstphonenumber
			PanelProgrammingLog
			$xrsecondphonenumber = ControlGetText($windowtitleinuse, "", "TDBStringControl1") ;2nd Phone No. - No Text, Watch Control Class Instance
			$programmingtowrite = "Second Phone Number: " & $xrsecondphonenumber
			PanelProgrammingLog
		Case "Cellular Network"
			$xrencryptionenabled = ControlCommand($windowtitleinuse, "", "Encryption", "IsVisible", "")
			If $xrencryptionenabled = 1 Then
				$xrencryption = ControlCommand($windowtitleinuse, "", "Encryption", "IsChecked", "") ;Encryption - ClassInstanceNN: TDBBooleanControl1
				If $xrencryption = 1 Then
					$programmingtowrite = "Encryption Enabled"
					PanelProgrammingLog
					$xrencryptionschema = ControlGetText($windowtitleinuse, "", "TDBListControl13") ;Encryption Schema - No Text, Watch Control Class Instance
					$programmingtowrite = "Encryption Scheme " & $xrencryptionschema
					PanelProgrammingLog
				Else
					$programmingtowrite = "Encryption Disabled"
					PanelProgrammingLog
					$xrencryptionschema = ""
					$programmingtowrite = "Encryption Scheme " & $xrencryptionschema
					PanelProgrammingLog
				EndIf
			Else
				$programmingtowrite = "No Encryption in this Panel"
				PanelProgrammingLog
			EndIf
			$xrreceiverip = ControlGetText($windowtitleinuse, "", "TDBStringControl3") ;Receiver IP - No Text, Watch Control Class Instance
			$programmingtowrite = "Receiver IP: " & $xrreceiverip
			PanelProgrammingLog
			$xrreceiverport = ControlGetText($windowtitleinuse, "", "TDBStringControl3") ;Receiver Port - No Text, Watch Control Class Instance
			$programmingtowrite = "Receiver Port: " & $xrreceiverport
			PanelProgrammingLog
	EndSwitch
	ControlClick($windowtitleinuse, "Operator Configuration", "TPageControl1") ;Activates first Tab "Path"
	Sleep(250)
	ControlSend($windowtitleinuse, "Operator Configuration", "TPageControl1", "{RIGHT}", 0) ; Moves to third tab for Advanced
	Sleep(250)
	;Advanced Tab
	Switch $xrcommunicationtype
		Case "None"
			;Details
			$xrsendcommtrouble = ControlGetText($windowtitleinuse, "", "TDBListControl1") ;Send Comm Trouble - No Text, Watch Control Class Instance
			$programmingtowrite = "Send Comm Trouble: " & $xrsendcommtrouble
			PanelProgrammingLog
			If $xrpathtype = "Backup" Then
				$xrduplicatealarms = ControlCommand($windowtitleinuse, "", "Duplicate Alarms", "IsChecked", "") ;Duplicate Alarms(Backup Path Only)
				If $xrduplicatealarms = 1 Then
					$programmingtowrite = "Duplicate Alarms Selected"
					PanelProgrammingLog
				Else
					$programmingtowrite = "Duplicate Alarms Not Selected"
					PanelProgrammingLog
				EndIf
				$xrfailtesthours = ControlGetText($windowtitleinuse, "", "TDBNumericControl1") ;Fail Test Hrs(Backup Path Only) - No Text, Watch Control Class Instance
				$programmingtowrite = "Fail Test Hours: " & $xrfailtesthours
				PanelProgrammingLog
			Else
				$xrduplicatealarms = "Duplicate Alarms Not Avaialbe for Primary Paths"
				$programmingtowrite = $xrduplicatealarms
				PanelProgrammingLog
				$xrfailtesthours = "Fail Test Hours Not Available for Primary Paths"
				$programmingtowrite = $xrfailtesthours
				PanelProgrammingLog
			EndIf
			;Reports
			If $xrpathtype = "Primary" Then
				$xralarmreports = ControlGetText($windowtitleinuse, "", "TDBListControl8") ;Alarm - No Text, Watch Control Class Instance
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ControlGetText($windowtitleinuse, "", "TDBListControl7") ;Supv/Trouble - No Text, Watch Control Class Instance
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ControlGetText($windowtitleinuse, "", "TDBListControl6") ;O/C_User - No Text, Watch Control Class Instance
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ControlGetText($windowtitleinuse, "", "TDBListControl5") ;Door Access - No Text, Watch Control Class Instance
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ControlGetText($windowtitleinuse, "", "TDBListControl4") ;Panic Test - No Text, Watch Control Class Instance
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			Else
				$xralarmreports = ""
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ""
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ""
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ""
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ""
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			EndIf
		Case "Digital Dialer"
			;Details
			$xrsendcommtrouble = ControlGetText($windowtitleinuse, "", "TDBListControl1") ;Send Comm Trouble - No Text, Watch Control Class Instance
			$programmingtowrite = "Send Comm Trouble: " & $xrsendcommtrouble
			PanelProgrammingLog
			$xrsendpathinformation = ControlCommand($windowtitleinuse, "", "Send Path Information", "IsChecked", "") ;Duplicate Alarms(Backup Path Only) - No Text, Watch Control Class Instance
				If $xrsendpathinformation = 1 Then
					$programmingtowrite = "Send Path Information Selected"
					PanelProgrammingLog
				Else
					$programmingtowrite = "Send Path Information Not Selected"
					PanelProgrammingLog
				EndIf
			If $xrpathtype = "Backup" Then
				$xrduplicatealarms = ControlCommand($windowtitleinuse, "", "Duplicate Alarms", "IsChecked", "") ;Duplicate Alarms(Backup Path Only) - No Text, Watch Control Class Instance
				If $xrduplicatealarms = 1 Then
					$programmingtowrite = "Duplicate Alarms Selected"
					PanelProgrammingLog
				Else
					$programmingtowrite = "Duplicate Alarms Not Selected"
					PanelProgrammingLog
				EndIf
				$xrfailtesthours = ControlGetText($windowtitleinuse, "", "TDBNumericControl1") ;Fail Test Hrs(Backup Path Only) - No Text, Watch Control Class Instance
				$programmingtowrite = "Fail Test Hours: " & $xrfailtesthours
				PanelProgrammingLog
			Else
				$xrduplicatealarms = "Duplicate Alarms Not Avaialbe for Primary Paths"
				$programmingtowrite = $xrduplicatealarms
				PanelProgrammingLog
				$xrfailtesthours = "Fail Test Hours Not Available for Primary Paths"
				$programmingtowrite = $xrfailtesthours
				PanelProgrammingLog
			EndIf
			;DD/CID Details
			$xr893A = ControlCommand($windowtitleinuse, "", "893 A", "IsChecked", "") ;893A
			If $xr893A = 1 Then
				$programmingtowrite = "893A Option is Enabled"
				PanelProgrammingLog
			Else
				$programmingtowrite = "893A Option is Not Enabled"
				PanelProgrammingLog
			EndIf
			$xrsecondlineprefix = ControlGetText($windowtitleinuse, "", "TDBStringControl3") ;Second Line Prefix - No Text, Watch Control Class Instance
			$programmingtowrite = $xrsecondlineprefix
			PanelProgrammingLog
			$xralarmswitchover = ControlGetText($windowtitleinuse, "", "TDBNumericControl3") ;Alarm Switchover - No Text, Watch Control Class Instance
			$programmingtowrite = $xralarmswitchover
			PanelProgrammingLog
			;Reports
			If $xrpathtype = "Primary" Then
				$xralarmreports = ControlGetText($windowtitleinuse, "", "TDBListControl8") ;Alarm - No Text, Watch Control Class Instance
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ControlGetText($windowtitleinuse, "", "TDBListControl7") ;Supv/Trouble - No Text, Watch Control Class Instance
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ControlGetText($windowtitleinuse, "", "TDBListControl6") ;O/C_User - No Text, Watch Control Class Instance
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ControlGetText($windowtitleinuse, "", "TDBListControl5") ;Door Access - No Text, Watch Control Class Instance
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ControlGetText($windowtitleinuse, "", "TDBListControl4") ;Panic Test - No Text, Watch Control Class Instance
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			Else
				$xralarmreports = ""
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ""
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ""
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ""
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ""
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			EndIf
		Case "Network"
			;Details
			$xrretryseconds = ControlGetText($windowtitleinuse, "", "TDBNumericControl2") ;Retry Seconds - No Text, Watch Control Class Instance
			$programmingtowrite = $xrretryseconds
			PanelProgrammingLog
			$xrsubstitutioncode = ControlGetText($windowtitleinuse, "", "TDBListControl3") ;Retry Seconds - No Text, Watch Control Class Instance
			$programmingtowrite = $xrsubstitutioncode
			PanelProgrammingLog
			$xrprotocol = ControlGetText($windowtitleinuse, "", "TDBListControl2") ;Retry Seconds - No Text, Watch Control Class Instance
			$programmingtowrite = $xrprotocol
			PanelProgrammingLog
			$xrsendcommtrouble = ControlGetText($windowtitleinuse, "", "TDBListControl1") ;Send Comm Trouble - No Text, Watch Control Class Instance
			$programmingtowrite = "Send Comm Trouble: " & $xrsendcommtrouble
			PanelProgrammingLog
			$xrsendpathinformation = ControlCommand($windowtitleinuse, "", "Send Path Information", "IsChecked", "") ;Duplicate Alarms(Backup Path Only) - No Text, Watch Control Class Instance
				If $xrsendpathinformation = 1 Then
					$programmingtowrite = "Send Path Information Selected"
					PanelProgrammingLog
				Else
					$programmingtowrite = "Send Path Information Not Selected"
					PanelProgrammingLog
				EndIf
			If $xrpathtype = "Backup" Then
				$xrduplicatealarms = ControlCommand($windowtitleinuse, "", "Duplicate Alarms", "IsChecked", "") ;Duplicate Alarms(Backup Path Only) - No Text, Watch Control Class Instance
				If $xrduplicatealarms = 1 Then
					$programmingtowrite = "Duplicate Alarms Selected"
					PanelProgrammingLog
				Else
					$programmingtowrite = "Duplicate Alarms Not Selected"
					PanelProgrammingLog
				EndIf
				$xrfailtesthours = ControlGetText($windowtitleinuse, "", "TDBNumericControl1") ;Fail Test Hrs(Backup Path Only) - No Text, Watch Control Class Instance
				$programmingtowrite = "Fail Test Hours: " & $xrfailtesthours
				PanelProgrammingLog
			Else
				$xrduplicatealarms = "Duplicate Alarms Not Avaialbe for Primary Paths"
				$programmingtowrite = $xrduplicatealarms
				PanelProgrammingLog
				$xrfailtesthours = "Fail Test Hours Not Available for Primary Paths"
				$programmingtowrite = $xrfailtesthours
				PanelProgrammingLog
			EndIf
			;Reports
			If $xrpathtype = "Primary" Then
				$xralarmreports = ControlGetText($windowtitleinuse, "", "TDBListControl8") ;Alarm - No Text, Watch Control Class Instance
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ControlGetText($windowtitleinuse, "", "TDBListControl7") ;Supv/Trouble - No Text, Watch Control Class Instance
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ControlGetText($windowtitleinuse, "", "TDBListControl6") ;O/C_User - No Text, Watch Control Class Instance
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ControlGetText($windowtitleinuse, "", "TDBListControl5") ;Door Access - No Text, Watch Control Class Instance
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ControlGetText($windowtitleinuse, "", "TDBListControl4") ;Panic Test - No Text, Watch Control Class Instance
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			Else
				$xralarmreports = ""
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ""
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ""
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ""
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ""
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			EndIf
		Case "Contact ID"
			;Details
			$xrsendcommtrouble = ControlGetText($windowtitleinuse, "", "TDBListControl1") ;Send Comm Trouble - No Text, Watch Control Class Instance
			$programmingtowrite = "Send Comm Trouble: " & $xrsendcommtrouble
			PanelProgrammingLog
			If $xrpathtype = "Backup" Then
				$xrduplicatealarms = ControlCommand($windowtitleinuse, "", "Duplicate Alarms", "IsChecked", "") ;Duplicate Alarms(Backup Path Only) - No Text, Watch Control Class Instance
				If $xrduplicatealarms = 1 Then
					$programmingtowrite = "Duplicate Alarms Selected"
					PanelProgrammingLog
				Else
					$programmingtowrite = "Duplicate Alarms Not Selected"
					PanelProgrammingLog
				EndIf
				$xrfailtesthours = ControlGetText($windowtitleinuse, "", "TDBNumericControl1") ;Fail Test Hrs(Backup Path Only) - No Text, Watch Control Class Instance
				$programmingtowrite = "Fail Test Hours: " & $xrfailtesthours
				PanelProgrammingLog
			Else
				$xrduplicatealarms = "Duplicate Alarms Not Avaialbe for Primary Paths"
				$programmingtowrite = $xrduplicatealarms
				PanelProgrammingLog
				$xrfailtesthours = "Fail Test Hours Not Available for Primary Paths"
				$programmingtowrite = $xrfailtesthours
				PanelProgrammingLog
			EndIf
			;DD/CID Details
			$xr893A = ControlCommand($windowtitleinuse, "", "893 A", "IsChecked", "") ;893A
			If $xr893A = 1 Then
				$programmingtowrite = "893A Option is Enabled"
				PanelProgrammingLog
			Else
				$programmingtowrite = "893A Option is Not Enabled"
				PanelProgrammingLog
			EndIf
			$xrsecondlineprefix = ControlGetText($windowtitleinuse, "", "TDBStringControl3") ;Second Line Prefix - No Text, Watch Control Class Instance
			$programmingtowrite = $xrsecondlineprefix
			PanelProgrammingLog
			$xralarmswitchover = ControlGetText($windowtitleinuse, "", "TDBNumericControl3") ;Alarm Switchover - No Text, Watch Control Class Instance
			$programmingtowrite = $xralarmswitchover
			PanelProgrammingLog
			;Reports
			If $xrpathtype = "Primary" Then
				$xralarmreports = ControlGetText($windowtitleinuse, "", "TDBListControl8") ;Alarm - No Text, Watch Control Class Instance
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ControlGetText($windowtitleinuse, "", "TDBListControl7") ;Supv/Trouble - No Text, Watch Control Class Instance
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ControlGetText($windowtitleinuse, "", "TDBListControl6") ;O/C_User - No Text, Watch Control Class Instance
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ControlGetText($windowtitleinuse, "", "TDBListControl5") ;Door Access - No Text, Watch Control Class Instance
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ControlGetText($windowtitleinuse, "", "TDBListControl4") ;Panic Test - No Text, Watch Control Class Instance
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			Else
				$xralarmreports = ""
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ""
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ""
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ""
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ""
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			EndIf
		Case "Cellular Network"
			;Details
			$xrsubstitutioncode = ControlGetText($windowtitleinuse, "", "TDBListControl3") ;Retry Seconds - No Text, Watch Control Class Instance
			$programmingtowrite = $xrsubstitutioncode
			PanelProgrammingLog
			$xrfirstgprsapn = ControlGetText($windowtitleinuse, "", "TDBStringControl2") ;Retry Seconds - No Text, Watch Control Class Instance
			$programmingtowrite = $xrfirstgprsapn
			PanelProgrammingLog
			$xrsecondgprsapn = ControlGetText($windowtitleinuse, "", "TDBStringControl1") ;Retry Seconds - No Text, Watch Control Class Instance
			$programmingtowrite = $xrsecondgprsapn
			PanelProgrammingLog
			$xrsendcommtrouble = ControlGetText($windowtitleinuse, "", "TDBListControl1") ;Send Comm Trouble - No Text, Watch Control Class Instance
			$programmingtowrite = "Send Comm Trouble: " & $xrsendcommtrouble
			PanelProgrammingLog
			$xrsendpathinformation = ControlCommand($windowtitleinuse, "", "Send Path Information", "IsChecked", "") ;Duplicate Alarms(Backup Path Only) - No Text, Watch Control Class Instance
				If $xrsendpathinformation = 1 Then
					$programmingtowrite = "Send Path Information Selected"
					PanelProgrammingLog
				Else
					$programmingtowrite = "Send Path Information Not Selected"
					PanelProgrammingLog
				EndIf
			If $xrpathtype = "Backup" Then
				$xrduplicatealarms = ControlCommand($windowtitleinuse, "", "Duplicate Alarms", "IsChecked", "") ;Duplicate Alarms(Backup Path Only) - No Text, Watch Control Class Instance
				If $xrduplicatealarms = 1 Then
					$programmingtowrite = "Duplicate Alarms Selected"
					PanelProgrammingLog
				Else
					$programmingtowrite = "Duplicate Alarms Not Selected"
					PanelProgrammingLog
				EndIf
				$xrfailtesthours = ControlGetText($windowtitleinuse, "", "TDBNumericControl1") ;Fail Test Hrs(Backup Path Only) - No Text, Watch Control Class Instance
				$programmingtowrite = "Fail Test Hours: " & $xrfailtesthours
				PanelProgrammingLog
			Else
				$xrduplicatealarms = "Duplicate Alarms Not Avaialbe for Primary Paths"
				$programmingtowrite = $xrduplicatealarms
				PanelProgrammingLog
				$xrfailtesthours = "Fail Test Hours Not Available for Primary Paths"
				$programmingtowrite = $xrfailtesthours
				PanelProgrammingLog
			EndIf
			;Reports
			If $xrpathtype = "Primary" Then
				$xralarmreports = ControlGetText($windowtitleinuse, "", "TDBListControl8") ;Alarm - No Text, Watch Control Class Instance
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ControlGetText($windowtitleinuse, "", "TDBListControl7") ;Supv/Trouble - No Text, Watch Control Class Instance
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ControlGetText($windowtitleinuse, "", "TDBListControl6") ;O/C_User - No Text, Watch Control Class Instance
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ControlGetText($windowtitleinuse, "", "TDBListControl5") ;Door Access - No Text, Watch Control Class Instance
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ControlGetText($windowtitleinuse, "", "TDBListControl4") ;Panic Test - No Text, Watch Control Class Instance
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			Else
				$xralarmreports = ""
				$programmingtowrite = "Alarm Reports: " & $xralarmreports
				PanelProgrammingLog
				$xrsupvtroublereports = ""
				$programmingtowrite = "Supervisory Trouble Reports: " & $xrsupvtroublereports
				PanelProgrammingLog
				$xroc_userreports = ""
				$programmingtowrite = "Open and Close Reports: " & $xroc_userreports
				PanelProgrammingLog
				$xrdooraccessreports = ""
				$programmingtowrite = "Door Access Reports: " & $xrdooraccessreports
				PanelProgrammingLog
				$xrpanictestreports = ""
				$programmingtowrite = "Panic Test Reports: " & $xrpanictestreports
				PanelProgrammingLog
			EndIf
	EndSwitch
	ControlClick($windowtitleinuse, "Operator Configuration", "&OK") ;Closes Communication Paths Window
EndFunc

