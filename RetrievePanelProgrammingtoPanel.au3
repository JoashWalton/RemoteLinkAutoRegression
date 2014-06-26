#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.10.2
 Author:	Joshua Walton

 Script Function:
	Send Panel Programming from Link to Specific Panel

#ce ----------------------------------------------------------------------------
Opt("WinTitleMatchMode", 2)

#include <File.au3>
#include <FileConstants.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include "P:\JW\Auto It Scripts\Remote Link\Link Automatic Regression\Remote Link Automatic Regression\FunctionLibraryLinkRegression.au3"

Global $panelaccountundertestnum = 1914
Global $panelaccountundertestrec = 1
Global $panelaccountundertestname = ""
Global $panelaccountundertestip = "192.168.64.50"
Global $panelaccountundertestremkey = "12345678"

Local $LogFile = "\LogOperatorConfigurationLoadandMenuTest.txt"
Local $filepathstring = "C:\Link\Link.exe"
Local $windowtitleinuse = "Remote Link"

DefaultOperatorLinkLogin()
FindCurrentPanelFile($panelaccountundertestnum, $panelaccountundertestname, $panelaccountundertestrec)
EnterPanelProgrammingMenu()
RetrieveXRCommunicationPathProgramming($sfilename, $programmingtowrite, ByRef $aXRCommPathXSettings)
