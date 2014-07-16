#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=realsmart-smile.ico
#AutoIt3Wrapper_Outfile=RealSmart-Importer.exe
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.10.2
 Author:         Richard Easton
 Website:		 www.sellostring.co.uk
 Email:			 rich.easton@gmail.com

 Script Function:
	Realsmart SIMS Import tool.

#ce ----------------------------------------------------------------------------


#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Date.au3>

FileInstall("logo.jpg", @ScriptDir & "\logo.jpg", 0)

;must be an admin or have admin rights on server/workstation

;if not then bye bye ;)
If not IsAdmin() Then
	MsgBox(48, "Warning!", "Please run with administrator right")
	Exit
EndIf

Opt("TrayIconHide", 1)


$ini = @scriptdir & "\settings.ini"
;if settings.ini exists get data from there, if not leave inputs blank
if fileexists($ini) Then
	$SUi = iniread($ini, "Details", "User", "")
	$SPi = iniread($ini, "Details", "Pass", "")
	$SKi = iniread($ini, "Details", "Key", "")
EndIf


$RSI = GUICreate("RLSmart Importer", 265, 386, -1, -1)
GUISetBkColor(0xFFFFFF)
GUICtrlCreatePic(@ScriptDir & "\logo.jpg", 0,0, 265,115)
GUICtrlCreateGroup(" RealSmart / SIMS initial Import ", 8, 168, 249, 177)

GUICtrlCreateLabel("SIMS Username", 19, 192, 81, 17)
if fileexists($ini) Then
	$SU = GUICtrlCreateInput($SUi, 16, 212, 233, 21)
Else
	$SU = GUICtrlCreateInput("", 16, 212, 233, 21)
EndIf

GUICtrlCreateLabel("SIMS Password", 18, 240, 78, 17)
if fileexists($ini) Then
	$SP = GUICtrlCreateInput($SPi, 16, 258, 233, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_PASSWORD))
Else
	$SP = GUICtrlCreateInput("", 16, 258, 233, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_PASSWORD))
EndIf

GUICtrlCreateLabel("SIMS Key", 18, 288, 51, 17)
if fileexists($ini) Then
	$SK = GUICtrlCreateInput($SKi, 16, 308, 233, 21)
Else
	$SK = GUICtrlCreateInput("", 16, 308, 233, 21)
EndIf


GUICtrlCreateGroup("", -99, -99, 1, 1)
$save = GUICtrlCreateButton("Save Details", 8, 352, 80, 25)
$task = GUICtrlCreateButton("Create Task", 93, 352, 80, 25)
;checks of realsmart button and enables task button if present
$st = @ScriptDir & "\st.tmp"
runwait(@ComSpec & " /c " & 'schtasks /Query /TN realsmart-import >> ' & $st , "", @SW_SHOW)


if fileexists($st) Then
	$sttest = filereadline($st, 5)
	if StringInStr($st, "Ready") Then
		;disables task button
		guictrlsetstate($task, $GUI_DISABLE)
		FileRecycle($st)
	Else
		guictrlsetstate($task, $GUI_ENABLE)
		FileRecycle($st)
	EndIf
Else
	guictrlsetstate($task, $GUI_ENABLE)
EndIf

$exe = GUICtrlCreateButton("Execute Import", 178, 352, 80, 25)
;check of realsmart log, and disables the execute button if present
if not FileExists(@scriptdir & "\realsmart.log") then
	guictrlsetstate($exe, $GUI_ENABLE)
Else
	GUICtrlSetState($exe, $GUI_DISABLE)
EndIf

GUISetState(@SW_SHOW)

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		case $task
			runwait(@ComSpec & " /c " & 'SchTasks /Query /TN Realsmart-import >> ' & $st , @scriptdir, @SW_hide)
			$sttest = filereadline($st, 5)
			if stringinstr($sttest, "ERROR:") Then
				;creates scheduled task on server/workstation that runs daily at 06:00
				Runwait(@ComSpec & " /c " & 'SchTasks /Create /SC DAILY /RL HIGHEST /RU SYSTEM /TN Realsmart-import /TR "' & @scriptdir & '\realsmart.exe ' & $SUi & ' ' & $SPi & ' ' & $SKi & ' https://www.rlsmart.net/webservices" /ST 06:00' , @scriptdir, @SW_hide)
				;check to see if scheduled task was created correctly
				runwait(@ComSpec & " /c " & 'SchTasks /Query /TN Realsmart-import >> ' & $st , @scriptdir, @SW_hide)
				$sttest = filereadline($st, 5)
				if stringinstr($sttest, "READY") = true Then
					Msgbox(64, "Success", "Realsmart scheduled task has been created and will run at 06:00 everyday")
					FileRecycle($st)
					;disables task button
					guictrlsetstate($task, $GUI_DISABLE)
				Else
					Msgbox(48, "Warning!", "Realsmart scheduled task was not created, please make sure you have the correct rights to create tasks")
					FileRecycle($st)
				endif
			elseif stringinstr($sttest, "READY") = true Then
					Msgbox(48, "Warning!", "Realsmart scheduled task was not created, as it already exists")
					FileRecycle($st)
			EndIf

		case $save
			;read inputs for data
			$SU = guictrlread($SU)
			$SP = guictrlread($SP)
			$SK = guictrlread($SK)
			if $SU > "" and $SP > "" and $SK > "" Then
				;write data to settings.ini
				iniwrite($ini, "Details", "User", $SU)
				iniwrite($ini, "Details", "Pass", $SP)
				iniwrite($ini, "Details", "key", $SK)
				Msgbox(64,"Success!", "Details Saved")
			Else
				MsgBox(48, "Warning!", "Unable to save details as there is a problem with one or more of your entries, please check and try again")
			EndIf
		case $exe
			;read inputs for data
			$SU = guictrlread($SU)
			$SP = guictrlread($SP)
			$SK = guictrlread($SK)
			;check of existance of realsmart.exe and .rptdef
			$RLEXE = FileExists(@ScriptDir & "\realsmart.exe")
			$RLDef = FileExists(@ScriptDir & "\SmartAssess - V301109.RptDef")
			if $SU > "" and $SP > "" and $SK > "" Then
				if $RLEXE = 0 or $rldef = 0 Then
					Msgbox(48,"Warning!", "Realsmart.exe or Report definition are not present," & @CR & "Please make sure they in the same directory you ran this program file")
				Else
					msgbox(64,"Information", "The inital import is about to run, please await a completion confirmation message", 10)
					;disables execute button
					guictrlsetstate($exe, $GUI_DISABLE)
					;runs initial import with provided data
					Runwait(@ComSpec & " /c " & 'realsmart.exe ' & $SU & ' ' & $SP & ' ' & $SK & ' https://www.rlsmart.net/webservices', @scriptdir, @SW_HIDE)
				 	Msgbox(64, "Information", "The initial Import process is now completed, please email the log file which has been created to realsmart")
					if FileExists(@scriptdir & "\realsmart.log") then
						;check of realsmart.log and enables task button
						guictrlsetstate($task, $GUI_ENABLE)
					EndIf


				endif
			Else
				MsgBox(48, "Warning!", "There is a problem with one or more of your entries, please check and try again")
			EndIf
	EndSwitch
WEnd

