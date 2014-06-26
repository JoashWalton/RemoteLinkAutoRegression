#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         Joshua Walton

 Script Function:
	Load with 604 Operators and Test System>Operator Configuration Menus

#ce ----------------------------------------------------------------------------

Opt("WinTitleMatchMode", 2)

#include <File.au3>
#include <Array.au3>
#include "P:\JW\Auto It Scripts\Remote Link\Link Automatic Regression\Remote Link Automatic Regression\FunctionLibraryLinkRegression.au3"
#include "P:\JW\Auto It Scripts\Remote Link\Link Automatic Regression\Remote Link Automatic Regression\LogOperatorConfigurationLoadandMenuTest.txt"

Local $LogFile = "\LogOperatorConfigurationLoadandMenuTest.txt"
Local $filepathstring = "C:\Link\Link.exe"
Local $windowtitleinuse = "Remote Link"
Local $i = 0 ;iterator for loops
Local $a = 0 ;iterator for first element in arrays
Local $a1 = 0 ;iterator for second element in arrays
Local $max_len = 32
Local $iterationpassword = ""
Local $viewonlyenabled = 0
Local $programmenuoption = 0
Local $panelaccountnumber = 13579
Local $scrolldown = 0
Local $currentprogoptform = ""

Local $aOpConfigSpecPerm[7][2] = _
	[["TDBCheckBox5","Administrator"],["TDBCheckBox6","Remote Update"],["TDBCheckBox8","Allow Export"],["TDBCheckBox7","Allow Import"], _
	["TDBCheckBox4","Cellular Activations"],["TDBCheckBox3","Advanced Filtering"],["TDBCheckBox2","Allow Trap"]]

Local $aPanelProgTabOpts[23][2] = _
	[["TDBCheckBox23","Communications"],["TDBCheckBox6","Network Options"],["TDBCheckBox22","Device Setup"],["TDBCheckBox1","Z-Wave"], _
	["TDBCheckBox21","Remote Options"],["TDBCheckBox20","System Reports"],["TDBCheckBox19","System Options"],["TDBCheckBox18","Output Options"], _
	["TDBCheckBox17","Output Groups"], ["TDBCheckBox16","Output Names"],["TDBCheckBox5","Keyfobs"],["TDBCheckBox15","Menu Display"], _
	["TDBCheckBox14","Status List"],["TDBCheckBox13","Printer Reports"],["TDBCheckBox7","PC Log Rprts"],["TDBCheckBox12","Area Information"], _
	["TDBCheckBox11","Zone Information"],["TDBCheckBox10","Access Code"],["TDBCheckBox9","Panel Send"],["TDBCheckBox8","Panel Retrieve"], _
	["TDBCheckBox4","Bell Outputs"],["TDBCheckBox3","Communication Paths"],["TDBCheckBox2","Messaging Setup"]]

Local $aUserStatusProgTabOpts[19][2] = _
	[["TDBCheckBox7","User Codes"],["TDBCheckBox4","Codes Visible"],["TDBCheckBox8","Profiles"],["TDBCheckBox9", "Schedules"], _
	["TDBCheckBox10","Holiday Dates"],["TDBCheckBox6","Forgive User"],["TDBCheckBox5","Panel Retrieve"],["TDBCheckBox17","Arm/Disarm"], _
	["TDBCheckBox16","Alarm Silence"],["TDBCheckBox15","Sensor Reset"],["TDBCheckBox13","Set Time Date"],["TDBCheckBox12","LX-Bus Diag."], _
	["TDBCheckBox19","Zone Status"],["TDBCheckBox18","Output Status"],["TDBCheckBox14","Send Message"],["TDBCheckBox11","List Z-Wave"], _
	["TDBCheckBox3","Alarm Ack"],["TDBCheckBox2","Alarm Remove"],["TDBCheckBox1", "Alarm Disable"]]

Local $aReceiverTabOpts[19][2] = _
	[["TDBCheckBox7","System Options"],["TDBCheckBox6","Hosts"],["TDBCheckBox5","Line Cards"],["TDBCheckBox4","Status"], _
	["TDBCheckBox3", "Printer"], ["TDBCheckBox2", "Serial Ports"], ["TDBCheckBox1", "Diagnostics"]]

Local $aProgamMenuFormTitles550[26] = _
	["Communication Paths", "Network Options", "Messaging Setup", "Device Setup", "Z-Wave Setup", "Favorites", "Remote Optiions", "System Reports", _
	"System Options", "Bell Options", "Output Options", "Output Information", "Output Groups", "Menu Display", "Status List", "PC Log Reports", _
	"Area Information", "Zone Information", "Keyfobs", "Holiday Dates", "Time Schedules", "Area Schedules", "Output/Door/Favorite Schedules", _
	"Profiles", "User Codes", "Access Code"]



Func AccountAccessRestrictOperation($windowtitleinuse,$LogFile,$MessageToWrite)
	ControlClick($windowtitleinuse, "", "TGroupButton1") ;Toggle Account Access>Restrict On
	Sleep(250)
	$MessageToWrite = "Operator " & $operatorusername & " has Account Access set to Restrict."
	LogFileWrite($LogFile, $MessageToWrite)
	Local $accaccessall = ControlCommand ($windowtitleinuse, "", "TGroupButton2", "IsChecked", "") ;Check AccountAccess>All option is deselected when Restrict is on
	If $accaccessall = 1 Then
		MsgBox(4112, "Failure!", "Not the Expected Result, Account Access> All should not be selected.'.")
		$MessageToWrite = "Failure - not the expexted result, Account Access> All should not be selected."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf
	ControlClick($windowtitleinuse, "", "TGroupButton2") ;Toggle Account Access>All On
	Sleep(250)
	Local $accaccessall = ControlCommand ($windowtitleinuse, "", "TGroupButton2", "IsChecked", "") ;Check AccountAccess>All option is deselected when Restrict is on
	If $accaccessall = 1 Then
		;MsgBox(4144, "Success!", "Account Access> All is now selected for " & $operatorusername & ".", 2)
		$MessageToWrite = "Operator " & $operatorusername & " has Account Access> All selected."
		LogFileWrite($LogFile, $MessageToWrite)
	EndIf
EndFunc
Func ViewOnlySpecialPermissionsOptions($windowtitleinuse,$LogFile,$MessageToWrite,$viewonlyenabled)
	ControlClick($windowtitleinuse, "", "TDBCheckBox1") ;Toggle User Restrictions>View Only On
	Sleep(250)
	$MessageToWrite = "Operator " & $operatorusername & " has Restrictions> View Only on, only specific Special Permissions are avaialble."
	LogFileWrite($LogFile, $MessageToWrite)
	Send("{TAB}")
	Sleep(500)
	Local $controltext = ControlGetText($windowtitleinuse, "", "TDBCheckBox3") ;Check Next Valid Option is Special Permissions>Advanced Filtering with View Only On
	If $controltext <> "Advanced Filtering" Then
		MsgBox(4112, "Failure!", "Not the Expected Result, next option after Toggling View Only is Advanced Filtering.")
		$MessageToWrite = "Next available Special Permission, with View Only On, is SUPPOSED to be Advanced Filtering"
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	Else
		ControlClick($windowtitleinuse, "", "TDBCheckBox3") ;Toggle SpecialPermissions>Advanced Filtering On
		Sleep(250)
		$MessageToWrite = "Operator " & $operatorusername & " has Special Peremission> Advanced Filtering turned on."
		LogFileWrite($LogFile, $MessageToWrite)
		ControlClick($windowtitleinuse, "", "TDBCheckBox2") ;Toggle SpecialPermissions>Allow Trap On
		Sleep(250)
		$MessageToWrite = "Operator " & $operatorusername & " has Special Peremission> Allow Trap turned on."
		LogFileWrite($LogFile, $MessageToWrite)
		ControlClick($windowtitleinuse, "", "TButton2") ;Click Clear Button to set Special Permissions>Advanced Filtering and Allow Trap to Off
		Sleep(250)
		$MessageToWrite = "All Special Permissions, with View Only On, have been turned off with Clear button."
		LogFileWrite($LogFile, $MessageToWrite)
		Local $advfilter = ControlCommand($windowtitleinuse, "", "TDBCheckBox3", "IsChecked", "")
		If $advfilter = 1 Then
			MsgBox(4112, "Failure!", "Not the Expected Result, Special Permissions> Advanced Filtering should not be selected.")
			$MessageToWrite = "Failure - not the expected result, Special Permissions> Advanced Filtering should not be selected."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			$MessageToWrite = "Special Permissions>Advanced Filtering has been turned off with Clear button."
			LogFileWrite($LogFile, $MessageToWrite)
		EndIf
		Local $allowtrap = ControlCommand($windowtitleinuse, "", "TDBCheckBox2", "IsChecked", "")
		If $allowtrap = 1 Then
			MsgBox(4112, "Failure!", "Not the Expected Result, Special Permissions> Allow Trap should not be selected.")
			$MessageToWrite = "Failure - not the expected result, Special Permissions> Allow Trap should not be selected."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			$MessageToWrite = "Special Permissions> Allow Trap has been turned off with Clear button."
			LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	EndIf
	ControlClick($windowtitleinuse, "", "TDBCheckBox1") ;Toggle User Restrictions>View Only Off
	Sleep(250)
	$viewonlyenabled = 0
	$MessageToWrite = "Operator " & $operatorusername & " has Restrictions> View Only off, all specific Special Permissions are avaialble."
	LogFileWrite($LogFile, $MessageToWrite)
EndFunc
Func SpecialPermissonsAllClearFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	ControlClick($windowtitleinuse, "", "TButton2") ;Toggles Clear button for Special Permissions - all off
	Sleep(250)
	$MessageToWrite = "Operator " & $operatorusername & " has their Special Permissions cleared with Clear button."
	LogFileWrite($LogFile, $MessageToWrite)
	For $a = 0 To 6 Step 1
		Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aOpConfigSpecPerm[$a][$a1], "IsChecked", "")
		If $checkedstatus = 1 Then
			MsgBox(4112, "Failure!", "Not the Expected Result, " & $aOpConfigSpecPerm[$a][$a1+1] & " should not be selected.")
			$MessageToWrite = "Failure, no Special Permissions should be selected on."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			$MessageToWrite = "Success - Special Permission " & $aOpConfigSpecPerm[$a][$a1+1] & " is turned off."
			LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	Next
		ControlClick($windowtitleinuse, "", "TButton3") ;Toggles All button for Special Permissions - all on
		Sleep(250)
		$MessageToWrite = "Operator " & $operatorusername & " has their Special Permissions enabled with All button."
		LogFileWrite($LogFile, $MessageToWrite)
		For $a = 0 To 6 Step 1
			Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aOpConfigSpecPerm[$a][$a1], "IsChecked", "")
			Sleep(250)
			If $checkedstatus = 0 Then
				MsgBox(4112, "Failure!", "Not the Expected Result, " & $aOpConfigSpecPerm[$a][$a1+1] & " should be selected.")
				$MessageToWrite = "Failure, Special Permissions " &  $aOpConfigSpecPerm[$a][$a1+1] & " should be selected."
				LogFileWrite($LogFile, $MessageToWrite)
				Exit
			Else
				$MessageToWrite = "Success - Special Permission " & $aOpConfigSpecPerm[$a][$a1+1] & " is turned on."
				LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	Next
EndFunc
Func PanelProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	ControlClick($windowtitleinuse, "", "TPageControl1") ;Activates first Tab "Operator Information"
	Sleep(250)
	ControlSend($windowtitleinuse, "Operator Configuration", "TPageControl1", "{RIGHT}", 0) ; Moves to third tab for Panel Programming
	Sleep(250)
	Local $sheettext = ControlGetText($windowtitleinuse, "", "TTabSheet1")
	;MsgBox(4144, "Test", "Will now begin testing " & $sheettext, 2)
	If $sheettext = "Panel Programming" Then
		Local $controltext = ControlGetText($windowtitleinuse, "", "TDBCheckBox23")
		If $controltext <> "Communications" Then
			MsgBox(4112, "Failure!", "Not the Expected Result, first option in 'Panel Programming' tab is supposed to be 'Communications'.")
			$MessageToWrite = "Failure - Query as to first checkbox on Panel Programming is supposed to be Communications."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			For $a = 0 To 22 Step 1
				Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aPanelProgTabOpts[$a][$a1], "IsChecked", "")
				If $checkedstatus = 0 Then
					;MsgBox(4144, "Success!", "As Expected, " & $aPanelProgTabOpts[$a][$a1+1] & " should not be selected on.", 1)
					$MessageToWrite = $aPanelProgTabOpts[$a][$a1+1] & " as expected is not selected on by default."
					LogFileWrite($LogFile, $MessageToWrite)
				Else
					MsgBox(4112, "Failure!", "Not the Expected Result, " & $aPanelProgTabOpts[$a][$a1+1] & " should not be selected.")
					$MessageToWrite = "Failure! " & $aPanelProgTabOpts[$a][$a1+1] & " is not turned on in a default situation."
					LogFileWrite($LogFile, $MessageToWrite)
					Exit
				EndIf
			Next
				ControlClick($windowtitleinuse, "", "TButton2") ;Toggles All button for Panel Programming options - all on
				$MessageToWrite = "Operator " & $operatorusername & " has their Panel Programming options enabled with All button."
				LogFileWrite($LogFile, $MessageToWrite)
				For $a = 0 To 22 Step 1
					Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aPanelProgTabOpts[$a][$a1], "IsChecked", "")
					If $checkedstatus = 1 Then
						;MsgBox(4144, "Success!", "As Expected, " & $aPanelProgTabOpts[$a][$a1+1] & " should be selected on.", 1)
						$MessageToWrite = $aPanelProgTabOpts[$a][$a1+1] & " as expected is now selected on."
						LogFileWrite($LogFile, $MessageToWrite)
					Else
						MsgBox(4112, "Failure!", "Not the Expected Result, " & $aPanelProgTabOpts[$a][$a1+1] & " should be selected.")
						$MessageToWrite = "Failure! " & $aPanelProgTabOpts[$a][$a1+1] & " is not turned on as expected."
						LogFileWrite($LogFile, $MessageToWrite)
						Exit
					EndIf
				Next
					ControlClick($windowtitleinuse, "", "TButton1") ;Toggles Clear button for Panel Programming options - all off
					$MessageToWrite = "Operator " & $operatorusername & " has their Panel Programing options disabled with the CLEAR button."
					LogFileWrite($LogFile, $MessageToWrite)
					For $a = 0 To 22 Step 1
						Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aPanelProgTabOpts[$a][$a1], "IsChecked", "")
						If $checkedstatus = 0 Then
							;MsgBox(4144, "Success!", "As Expected, " & $aPanelProgTabOpts[$a][$a1+1] & " should not be selected on.", 1)
							$MessageToWrite = $aPanelProgTabOpts[$a][$a1+1] & " as expected is not selected on by default."
							LogFileWrite($LogFile, $MessageToWrite)
						Else
							MsgBox(4112, "Failure!", "Not the Expected Result, " & $aPanelProgTabOpts[$a][$a1+1] & " should not be selected.")
							$MessageToWrite = "Failure! " & $aPanelProgTabOpts[$a][$a1+1] & " is not turned on in a default situation."
							LogFileWrite($LogFile, $MessageToWrite)
							Exit
						EndIf
					Next
						ControlClick($windowtitleinuse, "", "TButton2") ;Toggles All button for Panel Programming options - all on
						$MessageToWrite = "Operator " & $operatorusername & " has their Panel Programming options enabled with All button."
						LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	Else
		$MessageToWrite = "Query as to correct tab returned unexpected result - was supposed to be Panel Programming."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf
EndFunc
Func UserStatusProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	ControlClick($windowtitleinuse, "Operator Configuration", "TPageControl1") ;Activates first Tab "Operator Information"
	Sleep(250)
	ControlSend($windowtitleinuse, "Operator Configuration", "TPageControl1", "{RIGHT}", 0) ; Moves to third tab for User and Status Programming
	Sleep(250)
	Local $sheettext = ControlGetText($windowtitleinuse, "", "TTabSheet1")
	Sleep(250)
	;MsgBox(4144, "Test", "Will now begin testing " & $sheettext, 5)
	If $sheettext = "User and Status Programming" Then
		Local $controltext = ControlGetText($windowtitleinuse, "", "TDBCheckBox7")
		If $controltext <> "User Codes" Then
			MsgBox(4112, "Failure!", "Not the Expected Result, first option in 'User and Status Programming' tab is supposed to be 'User Codes'.")
			$MessageToWrite = "Query as to first checkbox on User and Status Programming is supposed to be User Codes."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			For $a = 0 To 18 Step 1
				Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aUserStatusProgTabOpts[$a][$a1], "IsChecked", "")
				If $checkedstatus = 0 Then
					;MsgBox(4144, "Success!", "As Expected, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should not be selected on.", 2)
					$MessageToWrite = $aUserStatusProgTabOpts[$a][$a1+1] & " as expected is not selected on by default."
					LogFileWrite($LogFile, $MessageToWrite)
				Else
					If $aUserStatusProgTabOpts[$a][$a1+1] = "Panel Retrieve" Then
						;MsgBox(4144, "Acceptable", "This is the expected Result, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should be selected.", 2)
						$MessageToWrite = "Acceptable! " & $aUserStatusProgTabOpts[$a][$a1+1] & " is turned on when Panel Programming options are selected."
						LogFileWrite($LogFile, $MessageToWrite)
					Else
						MsgBox(4112, "Failure!", "Not the Expected Result, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should not be selected.")
						$MessageToWrite = "Failure! " & $aUserStatusProgTabOpts[$a][$a1+1] & " is not turned on in a default situation."
						LogFileWrite($LogFile, $MessageToWrite)
						Exit
					EndIf
				EndIf
			Next
				ControlClick($windowtitleinuse, "", "TButton4") ;Toggles All button for User Programming - all on for that section only
				ControlClick($windowtitleinuse, "", "TButton6") ;Toggles All button for Panel Stastus - all on for that section only
				ControlClick($windowtitleinuse, "", "TButton1") ;Toggles All button for Alarm List - all on for that section only
				$MessageToWrite = "Operator " & $operatorusername & " has all their User and Status Programming options enabled with respective All buttons."
				LogFileWrite($LogFile, $MessageToWrite)
				For $a = 0 To 18 Step 1
					Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aUserStatusProgTabOpts[$a][$a1], "IsChecked", "")
					If $checkedstatus = 1 Then
						;MsgBox(4144, "Success!", "As Expected, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should be selected on.", 2)
						$MessageToWrite = $aUserStatusProgTabOpts[$a][$a1+1] & " as expected is now selected on."
						LogFileWrite($LogFile, $MessageToWrite)
					Else
						MsgBox(4112, "Failure!", "Not the Expected Result, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should be selected.")
						$MessageToWrite = "Failure! " & $aUserStatusProgTabOpts[$a][$a1+1] & " is not turned on as expected."
						LogFileWrite($LogFile, $MessageToWrite)
						Exit
					EndIf
				Next
					ControlClick($windowtitleinuse, "", "TButton3") ;Toggles Clear button for User Programming - all off
					ControlClick($windowtitleinuse, "", "TButton5") ;Toggles Clear button for Panel Status - all off
					ControlClick($windowtitleinuse, "", "TButton2") ;Toggles Clear button for Alarm List - all off
					$MessageToWrite = "Operator " & $operatorusername & " has all their User and Status Programing options disabled with respective CLEAR button."
					LogFileWrite($LogFile, $MessageToWrite)
					For $a = 0 To 18 Step 1
						Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aUserStatusProgTabOpts[$a][$a1], "IsChecked", "")
						If $checkedstatus = 0 Then
							;MsgBox(4144, "Success!", "As Expected, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should not be selected on.", 2)
							$MessageToWrite = $aUserStatusProgTabOpts[$a][$a1+1] & " as expected is not selected on."
							LogFileWrite($LogFile, $MessageToWrite)
						Else
							If $aUserStatusProgTabOpts[$a][$a1+1] = "Panel Retrieve" Then
								MsgBox(4144, "Acceptable", "This is the expected Result, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should be selected.", 2)
								$MessageToWrite = "Acceptable! " & $aUserStatusProgTabOpts[$a][$a1+1] & " is turned on when Panel Programming options are selected."
								LogFileWrite($LogFile, $MessageToWrite)
							Else
								MsgBox(4112, "Failure!", "Not the Expected Result, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should not be selected.")
								$MessageToWrite = "Failure! " & $aUserStatusProgTabOpts[$a][$a1+1] & " is not supposed to be turned on as expected."
								LogFileWrite($LogFile, $MessageToWrite)
								Exit
							EndIf
						EndIf
					Next
						ControlClick($windowtitleinuse, "", "TButton4") ;Toggles All button for User Programming - all on for that section only
						ControlClick($windowtitleinuse, "", "TButton6") ;Toggles All button for Panel Stastus - all on for that section only
						ControlClick($windowtitleinuse, "", "TButton1") ;Toggles All button for Alarm List - all on for that section only
						$MessageToWrite = "Operator " & $operatorusername & " has all their User and Status Programming options enabled with respective All buttons."
						LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	Else
		$MessageToWrite = "Query as to correct tab returned unexpected result - was supposed to be User and Status Programming."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf
EndFunc
Func ReceiverProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	ControlClick($windowtitleinuse, "", "TPageControl1") ;Activates first Tab "Operator Information"
	Sleep(250)
	ControlSend($windowtitleinuse, "Operator Configuration", "TPageControl1", "{RIGHT}", 0) ; Moves to third tab for Receiver Programming
	Sleep(250)
	Local $sheettext = ControlGetText($windowtitleinuse, "", "TTabSheet1")
	;MsgBox(4144, "Test", "Will now begin testing " & $sheettext, 5)
	If $sheettext = "Receiver" Then
		Local $controltext = ControlGetText($windowtitleinuse, "", "TDBCheckBox7")
		If $controltext <> "System Options" Then
			MsgBox(4112, "Failure!", "Not the Expected Result, first option in the 'Receiver' tab is supposed to be 'System Options'.")
			$MessageToWrite = "Query as to first checkbox on Receiver is supposed to be System Options."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			For $a = 0 To 6 Step 1
				Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aReceiverTabOpts[$a][$a1], "IsChecked", "")
				If $checkedstatus = 0 Then
					;MsgBox(4144, "Success!", "As Expected, " & $aReceiverTabOpts[$a][$a1+1] & " should not be selected on.", 2)
					$MessageToWrite = $aReceiverTabOpts[$a][$a1+1] & " as expected is not selected on by default."
					LogFileWrite($LogFile, $MessageToWrite)
				Else
					MsgBox(4112, "Failure!", "Not the Expected Result, " & $aReceiverTabOpts[$a][$a1+1] & " should not be selected.")
					$MessageToWrite = "Failure! " & $aReceiverTabOpts[$a][$a1+1] & " is not turned on in a default situation."
					LogFileWrite($LogFile, $MessageToWrite)
					Exit
				EndIf
			Next
				ControlClick($windowtitleinuse, "", "TButton2") ;Toggles All button for Receiver Programming - all on for that section only
				$MessageToWrite = "Operator " & $operatorusername & " has all their Receiver Programming options enabled with the All button."
				LogFileWrite($LogFile, $MessageToWrite)
				For $a = 0 To 6 Step 1
					Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aReceiverTabOpts[$a][$a1], "IsChecked", "")
					If $checkedstatus = 1 Then
						;MsgBox(4144, "Success!", "As Expected, " & $aReceiverTabOpts[$a][$a1+1] & " should be selected on.", 2)
						$MessageToWrite = $aReceiverTabOpts[$a][$a1+1] & " as expected is now selected on."
						LogFileWrite($LogFile, $MessageToWrite)
					Else
						MsgBox(4112, "Failure!", "Not the Expected Result, " & $aReceiverTabOpts[$a][$a1+1] & " should be selected.")
						$MessageToWrite = "Failure! " & $aReceiverTabOpts[$a][$a1+1] & " is not turned on as expected."
						LogFileWrite($LogFile, $MessageToWrite)
						Exit
					EndIf
				Next
					ControlClick($windowtitleinuse, "", "TButton1") ;Toggles Clear button for User Programming - all off
					$MessageToWrite = "Operator " & $operatorusername & " has all their Receiver Programing options disabled with the CLEAR button."
					LogFileWrite($LogFile, $MessageToWrite)
					For $a = 0 To 6 Step 1
						Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aReceiverTabOpts[$a][$a1], "IsChecked", "")
						If $checkedstatus = 0 Then
							;MsgBox(4144, "Success!", "As Expected, " & $aReceiverTabOpts[$a][$a1+1] & " should not be selected on.", 2)
							$MessageToWrite = $aReceiverTabOpts[$a][$a1+1] & " as expected is not selected on."
							LogFileWrite($LogFile, $MessageToWrite)
						Else
							MsgBox(4112, "Failure!", "Not the Expected Result, " & $aReceiverTabOpts[$a][$a1+1] & " should not be selected.")
							$MessageToWrite = "Failure! " & $aReceiverTabOpts[$a][$a1+1] & " is not supposed to be turned on as expected."
							LogFileWrite($LogFile, $MessageToWrite)
							Exit
						EndIf
					Next
						ControlClick($windowtitleinuse, "", "TButton2") ;Toggles All button for Receiver Programming - all on for that section only
						$MessageToWrite = "Operator " & $operatorusername & " has all their Receiver Programming options enabled with the All button."
						LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	Else
		$MessageToWrite = "Query as to correct tab returned unexpected result - was supposed to be Receiver."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf
EndFunc

Func CheckAccountAccessRestrictOperation($windowtitleinuse,$LogFile,$MessageToWrite)
	Local $restrictaccess = ControlCommand($windowtitleinuse, "", "TGroupButton1", "IsChecked", "")
	If $restrictaccess = 1 Then
		MsgBox(4112, "Failure!", "This Operator should not have Restrict selected in Account Access.")
		$MessageToWrite = "Operator " & $operatorusername & " does not have Account Access set to Restrict as expected."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	Else
		MsgBox(4144, "Success!", "This Operator does not have Restrict selected in Account Access.", 3)
		$MessageToWrite = "Operator " & $operatorusername & " does not have Account Access set to Restrict as expected."
		LogFileWrite($LogFile, $MessageToWrite)
	EndIf
EndFunc
Func CheckViewOnlySpecialPermissionsOptions($windowtitleinuse,$LogFile,$MessageToWrite,$viewonlyenabled)
	If $viewonlyenabled = 1 Then
		$MessageToWrite = "Operator " & $operatorusername & " has authority to View Only.  They cannot check their own operator configuration."
		LogFileWrite($LogFile, $MessageToWrite)
	EndIf
EndFunc
Func CheckSpecialPermissonsAllClearFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	For $a = 0 To 6 Step 1
		Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aOpConfigSpecPerm[$a][$a1], "IsChecked", "")
		If $checkedstatus = 0 Then
			MsgBox(4112, "Failure!", "Not the Expected Result, " & $aOpConfigSpecPerm[$a][$a1+1] & " should be selected.")
			$MessageToWrite = "Failure, Special Permissions" & $aOpConfigSpecPerm[$a][$a1+1] & " should be selected on."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			;$MessageToWrite = "Success - Special Permission " & $aOpConfigSpecPerm[$a][$a1+1] & " is turned on."
			LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	Next
		$MessageToWrite = "Operator " & $operatorusername & " has their Special Permissions enabled as expected."
		LogFileWrite($LogFile, $MessageToWrite)
EndFunc
Func CheckPanelProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	ControlClick($windowtitleinuse, "", "TPageControl1") ;Activates first Tab "Operator Information"
	Sleep(250)
	ControlSend($windowtitleinuse, "Operator Configuration", "TPageControl1", "{RIGHT}", 0) ; Moves to third tab for Panel Programming
	Sleep(250)
	Local $sheettext = ControlGetText($windowtitleinuse, "", "TTabSheet1")
	;MsgBox(4144, "Test", "Will now begin testing " & $sheettext, 5)
	$MessageToWrite = "Verifying that Operator " & $operatorusername & " has the correct options selected in Panel Programming."
	LogFileWrite($LogFile, $MessageToWrite)
	If $sheettext = "Panel Programming" Then
		Local $controltext = ControlGetText($windowtitleinuse, "", "TDBCheckBox23")
		If $controltext <> "Communications" Then
			MsgBox(4112, "Failure!", "Not the Expected Result, first option in 'Panel Programming' tab is supposed to be 'Communications'.")
			$MessageToWrite = "Query as to first checkbox on Panel Programming is supposed to be Communications."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			For $a = 0 To 22 Step 1
				Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aPanelProgTabOpts[$a][$a1], "IsChecked", "")
				If $checkedstatus = 1 Then
					;MsgBox(4144, "Success!", "As Expected, " & $aPanelProgTabOpts[$a][$a1+1] & " should be selected on.", 1)
					$MessageToWrite = $aPanelProgTabOpts[$a][$a1+1] & " as expected is selected on."
					LogFileWrite($LogFile, $MessageToWrite)
				Else
					MsgBox(4112, "Failure!", "Not the Expected Result, " & $aPanelProgTabOpts[$a][$a1+1] & " should be selected.")
					$MessageToWrite = "Failure! " & $aPanelProgTabOpts[$a][$a1+1] & " is not turned on as was expected."
					LogFileWrite($LogFile, $MessageToWrite)
					Exit
				EndIf
			Next
				$MessageToWrite = "Operator " & $operatorusername & " has their Panel Programming options enabled as expected."
				LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	Else
		$MessageToWrite = "Query as to correct tab returned unexpected result - was supposed to be Panel Programming."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf
EndFunc
Func CheckUserStatusProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	ControlClick($windowtitleinuse, "Operator Configuration", "TPageControl1") ;Activates first Tab "Operator Information"
	Sleep(250)
	ControlSend($windowtitleinuse, "Operator Configuration", "TPageControl1", "{RIGHT}", 0) ; Moves to third tab for User and Status Programming
	Sleep(250)
	Local $sheettext = ControlGetText($windowtitleinuse, "", "TTabSheet1")
	Sleep(250)
	;MsgBox(4144, "Test", "Will now begin testing " & $sheettext, 5)
	$MessageToWrite = "Verifying that Operator " & $operatorusername & " has the correct options selected in User and Status Programming."
	LogFileWrite($LogFile, $MessageToWrite)
	If $sheettext = "User and Status Programming" Then
		Local $controltext = ControlGetText($windowtitleinuse, "", "TDBCheckBox7")
		If $controltext <> "User Codes" Then
			MsgBox(4112, "Failure!", "Not the Expected Result, first option in 'User and Status Programming' tab is supposed to be 'User Codes'.")
			$MessageToWrite = "Query as to first checkbox on User and Status Programming is supposed to be User Codes."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			For $a = 0 To 18 Step 1
				Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aUserStatusProgTabOpts[$a][$a1], "IsChecked", "")
				If $checkedstatus = 1 Then
					;MsgBox(4144, "Success!", "As Expected, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should be selected on.", 1)
					$MessageToWrite = $aUserStatusProgTabOpts[$a][$a1+1] & " as expected is selected on."
					LogFileWrite($LogFile, $MessageToWrite)
				Else
					MsgBox(4112, "Failure!", "Not the Expected Result, " & $aUserStatusProgTabOpts[$a][$a1+1] & " should be selected.")
					$MessageToWrite = "Failure! " & $aUserStatusProgTabOpts[$a][$a1+1] & " is not turned on as expected."
					LogFileWrite($LogFile, $MessageToWrite)
					Exit
				EndIf
			Next
				$MessageToWrite = "Operator " & $operatorusername & " has their User and Status Programming options enabled as expected."
				LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	Else
		$MessageToWrite = "Query as to correct tab returned unexpected result - was supposed to be User and Status Programming."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf
EndFunc
Func CheckReceiverProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	ControlClick($windowtitleinuse, "", "TPageControl1") ;Activates first Tab "Operator Information"
	Sleep(250)
	ControlSend($windowtitleinuse, "Operator Configuration", "TPageControl1", "{RIGHT}", 0) ; Moves to third tab for Receiver Programming
	Sleep(250)
	Local $sheettext = ControlGetText($windowtitleinuse, "", "TTabSheet1")
	;MsgBox(4144, "Test", "Will now begin testing " & $sheettext, 5)
	$MessageToWrite = "Verifying that Operator " & $operatorusername & " has the correct options selected in Receiver Programming."
	LogFileWrite($LogFile, $MessageToWrite)
	If $sheettext = "Receiver" Then
		Local $controltext = ControlGetText($windowtitleinuse, "", "TDBCheckBox7")
		If $controltext <> "System Options" Then
			MsgBox(4112, "Failure!", "Not the Expected Result, first option in the 'Receiver' tab is supposed to be 'System Options'.")
			$MessageToWrite = "Query as to first checkbox on Receiver is supposed to be System Options."
			LogFileWrite($LogFile, $MessageToWrite)
			Exit
		Else
			For $a = 0 To 6 Step 1
				Local $checkedstatus = ControlCommand($windowtitleinuse, "", $aReceiverTabOpts[$a][$a1], "IsChecked", "")
				If $checkedstatus = 1 Then
					;MsgBox(4144, "Success!", "As Expected, " & $aReceiverTabOpts[$a][$a1+1] & " should be selected on.", 2)
					$MessageToWrite = $aReceiverTabOpts[$a][$a1+1] & " as expected is selected on."
					LogFileWrite($LogFile, $MessageToWrite)
				Else
					MsgBox(4112, "Failure!", "Not the Expected Result, " & $aReceiverTabOpts[$a][$a1+1] & " should be selected.")
					$MessageToWrite = "Failure! " & $aReceiverTabOpts[$a][$a1+1] & " is not turned on as expected."
					LogFileWrite($LogFile, $MessageToWrite)
					Exit
				EndIf
			Next
				$MessageToWrite = "Operator " & $operatorusername & " has their Receiver Programming options enabled as expected."
				LogFileWrite($LogFile, $MessageToWrite)
		EndIf
	Else
		$MessageToWrite = "Query as to correct tab returned unexpected result - was supposed to be Receiver."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf
EndFunc

DefaultOperatorLinkLogin()

CreateListUsernames($aOperatorUsernames)
CreateListLastNames($aUserLastNames)
CreateListFirstNames($aUserFirstNames)

Do
	If WinWaitActive($windowtitleinuse, "Panel Information") Then ;Gets script to System> Operator Configuration
		ControlClick($windowtitleinuse, "&Cancel", "TButton4")
		Sleep(250)
		Send("!s")
		Sleep(250)
		Send("o")
		Sleep(250)
		$MessageToWrite = "Entered System> Operator Configuration, will begin testing."
		LogFileWrite($LogFile, $MessageToWrite)
		ControlClick($windowtitleinuse, "", "TButton8") ;Toggles New for new Operator to be configured
		Sleep(500)
	Else ;Reports problem getting to System> Operator Configuration and Exits
		$MessageToWrite = "Unable to enter System> Operator Configuration, will exit program."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf
	Sleep(250)
	$operatorusername = $aOperatorUsernames[Random(0, 603, 1)]
	ControlSend($windowtitleinuse, "Operator Configuration", "TDBEdit4", $operatorusername, 0)
	Sleep(250)
	$MessageToWrite = "Operator " & $i & " username is " & $operatorusername
	LogFileWrite($LogFile, $MessageToWrite)
	ControlClick($windowtitleinuse, "", "TDBEdit3")
	Sleep(250)
	$iterationpassword = RandomCharPasswordGenerator($max_len)
	ControlSend($windowtitleinuse, "Operator Configuration", "TDBEdit3", $iterationpassword, 0)
	Sleep(250)
	$MessageToWrite = "Operator " & $i & " password is " & $iterationpassword
	LogFileWrite($LogFile, $MessageToWrite)
	ControlClick($windowtitleinuse, "", "TEdit1")
	Sleep(250)
	ControlSend($windowtitleinuse, "Operator Configuration", "TEdit1", $iterationpassword, 0)
	Sleep(250)
	$MessageToWrite = "Operator " & $i & " confirmation password is " & $iterationpassword
	LogFileWrite($LogFile, $MessageToWrite)
	ControlClick($windowtitleinuse, "", "TDBEdit2")
	Sleep(250)
	Local $lastname = $aUserLastNames[Random(0, 999, 1)]
	ControlSend($windowtitleinuse, "Operator Configuration", "TDBEdit2", $lastname, 0)
	Sleep(250)
	$MessageToWrite = "Operator " & $i & " last name is " & $lastname
	LogFileWrite($LogFile, $MessageToWrite)
	ControlClick($windowtitleinuse, "", "TDBEdit1")
	Sleep(250)
	Local $firstname = $aUserFirstNames[Random(0, 999, 1)]
	ControlSend($windowtitleinuse, "Operator Configuration", "TDBEdit1", $firstname, 0)
	Sleep(250)
	$MessageToWrite = "Operator " & $i & " first name is " & $firstname
	LogFileWrite($LogFile, $MessageToWrite)

	AccountAccessRestrictOperation($windowtitleinuse,$LogFile,$MessageToWrite)
	ViewOnlySpecialPermissionsOptions($windowtitleinuse,$LogFile,$MessageToWrite,$viewonlyenabled)
	SpecialPermissonsAllClearFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	PanelProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	UserStatusProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	ReceiverProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)

	Sleep(250)
	ControlClick($windowtitleinuse, "Operator Configuration", "TButton16")
	Sleep(250)
	WinClose($windowtitleinuse, "")
	Sleep(1000)
	WinKill($windowtitleinuse, "")
	Sleep(1000)

	CustomOperatorLinkLogin($operatorusername, $iterationpassword)
	Sleep(500)

	WinActivate($windowtitleinuse, "Remote Link")

	If WinWaitActive($windowtitleinuse, "Panel Information") Then
		ControlClick($windowtitleinuse, "", "TButton10")
		Sleep(250)
		ControlSend($windowtitleinuse, "New Panel", "TEdit3", $panelaccountnumber, 1)
		Sleep(300)
		ControlClick($windowtitleinuse, "", "TButton2")
		Sleep(500)
		Do
			ControlSend($windowtitleinuse, "Panel Information", "TwwDBGrid1", "{DOWN}", 0)
			Sleep(300)
			Local $currentpanel = ControlGetText($windowtitleinuse, "", "TDBEdit13") ;TDBEdit11 on Win7
		Until $currentpanel = $panelaccountnumber
		ControlClick($windowtitleinuse, "", "&OK")
		Sleep(500)
		Send("!r")	;Enter Panel Programming
		Sleep(250)
		Send("{ENTER}") 	;Communications is selected
		Sleep(2000)
		$programmenuoption = 0
		$scrolldown = 0
		Do
			WinActivate($windowtitleinuse, "Remote Link")
			Local $currentprogoptform = ControlGetText($windowtitleinuse, "", 65280)
			MsgBox(4096, "Active Window", "Current Active Window is: " & $currentprogoptform & ".", 1)

			If $currentprogoptform = $aProgamMenuFormTitles550[$programmenuoption] Then
				$MessageToWrite = "Progam Menu Option " & $currentprogoptform & " is currently active as expected."
				LogFileWrite($LogFile, $MessageToWrite)
			Else
				MsgBox(4112, "Failure!", "The current window should be " & $aProgamMenuFormTitles550[$programmenuoption] & " and not " & $currentprogoptform & ".")
				$MessageToWrite = "The current window should be " & $aProgamMenuFormTitles550[$programmenuoption] & " and not " & $currentprogoptform & "."
				LogFileWrite($LogFile, $MessageToWrite)
				Exit
			EndIf

			Sleep(500)
			ControlClick($windowtitleinuse, "", "&Cancel")
			Sleep(1000)
			Send("!r")
			Sleep(1000)
			Send("{DOWN " & $scrolldown + 1 & "}")
			Sleep(1000)
			Send("{ENTER}")
			Sleep(3000)
			$scrolldown += 1
			$programmenuoption += 1
		Until $currentprogoptform = $aProgamMenuFormTitles550[25]
		Send("{ESC}")
		Sleep(250)
		Send("!s")
		Sleep(250)
		Send("o")
		Sleep(250)
		$MessageToWrite = "Entered System> Operator Configuration, will begin testing."
		LogFileWrite($LogFile, $MessageToWrite)
	Else ;Reports problem getting to System> Operator Configuration and Exits
		$MessageToWrite = "Unable to enter System> Operator Configuration, will exit program."
		LogFileWrite($LogFile, $MessageToWrite)
		Exit
	EndIf

	ControlClick($windowtitleinuse, "", "TwwDBGrid1")
	Sleep(250)
	Do
		ControlSend($windowtitleinuse, "Operator Configuration", "TwwDBGrid1", "{DOWN}", 0)
		Sleep(300)
		Local $currentopername = ControlGetText($windowtitleinuse, "", "TDBEdit4")
	Until $currentopername = $operatorusername
	ControlClick($windowtitleinuse, "", "TTabSheet1")
	Sleep(250)

	If $viewonlyenabled = 1 Then
		CheckViewOnlySpecialPermissionsOptions($windowtitleinuse,$LogFile,$MessageToWrite,$viewonlyenabled)
	Else
		CheckAccountAccessRestrictOperation($windowtitleinuse,$LogFile,$MessageToWrite)
		CheckViewOnlySpecialPermissionsOptions($windowtitleinuse,$LogFile,$MessageToWrite,$viewonlyenabled)
		CheckSpecialPermissonsAllClearFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
		CheckPanelProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
		CheckUserStatusProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
		CheckReceiverProgrammingTabFunctions($windowtitleinuse,$LogFile,$MessageToWrite)
	EndIf

	WinClose($windowtitleinuse, "")
	Sleep(1000)
	WinKill($windowtitleinuse, "")
	Sleep(1000)

	DefaultOperatorLinkLogin()

	$i += 1

Until $i = 5

MsgBox(4160, "Done", "Testing is complete.  Log File has results.", 15)
Exit

