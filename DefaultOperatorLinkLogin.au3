#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.8.1
 Author:         Joshua Walton

 Script Function:
	Login into Remote Link using default login information (new, new)

#ce ----------------------------------------------------------------------------

#include <File.au3>
#include <FileConstants.au3>

Local $windowtitleinuse = "Remote Link"
Local $operatorusername = "new"
Local $operatorpassword = "new"

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
