#NoEnv, UseUnsetLocal
#Warn
#SingleInstance force
SendMode Event
SetWorkingDir %A_ScriptDir% 
SetBatchLines, -1

; Creates a GUI for quick access to frequently used resources

#w::
Gui, New,,Quick Access Bar
Gui +AlwaysOnTop +ToolWindow
Gui, Color, E54D31

Gui, Font, cWhite s8 w600
Gui, Add, Text,, Websites
Gui, Font

Gui, Add, Button, gPressTimeSheet w150, Timesheet
Gui, Add, Button, gPressKIRC w150, KIRC
Gui, Add, Button, gPressSharePoint w150, SharePoint

Gui, Font, cWhite s8 w600
Gui, Add, Text,, Folders
Gui, Font

Gui, Add, Button, gPressUDrive w150, U: Drive

Gui, Show
Return

; ------------------- GUI CONTROLS PROGRAMMING ---------------------

PressTimeSheet:
run, iexplore.exe https://idcsapepp.corp.hatchglobal.com/irj/portal?NavigationTarget=navurl://a4c6004e4c9a20fd30f8437568091a55
Gui, Destroy
return

PressKIRC:
run, https://1415.sydneyplus.com/Hatch_SE/portal.aspx
Gui, Destroy
return

PressSharePoint:
run, https://hatchengineering.sharepoint.com/
Gui, Destroy
return

PressUDrive:
	IfWinExist, Vibrations (U:)
		WinActivate 
	else 
		run, Explorer U:\_PRJ\_PROJECTS
Gui, Destroy
return