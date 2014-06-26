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

Global $panelaccountundertestnum = 0
Global $panelaccountundertestrec = 1
Global $panelaccountundertestname = ""
Global $panelaccountundertestip = "192.168.64.50"
Global $panelaccountundertestremkey = "12345678"

DefaultOperatorLinkLogin()




