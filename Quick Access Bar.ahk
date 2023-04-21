#NoEnv, UseUnsetLocal
#Warn
#SingleInstance force
SendMode Event
SetWorkingDir %A_ScriptDir% 
SetBatchLines, -1

; Creates a GUI for quick access to frequently used resources

#w::
FileRead, CustomLinksFileContents, CustomLinks.txt

; Create Arrays
GrabAll := []
NamesArray := []
LinksArray := []   

Loop, Parse, CustomLinksFileContents, `n;>, `r%A_Space%%A_Tab%	    		; Parse variable contents line by line
    GrabAll.Push(A_LoopField)    											; Store lines in array
	
Loop, % GrabAll.MaxIndex()
{
	if (mod(A_Index, 2) == 1)
		NamesArray[Floor(A_Index/2) + 1] := GrabAll[A_Index]
	if (mod(A_Index, 2) == 0)
		LinksArray[Floor(A_Index/2)] := GrabAll[A_Index]
}
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
Gui, Add, Text,, U: Drive Folders
Gui, Font

Gui, Add, Button, gPressUDriveProjects w150, U: Projects
Gui, Add, Button, gPressUDriveReference w150, U: Reference

Gui, Font, cWhite s8 w600
Gui, Add, Text,, Dynamic Buttons
Gui, Font

Gui, Add, Button, gPressCopiedLink w150, Copied Address

Loop, 25
Gui, Add, Button, gPressDynamicButton v%A_Index% w150 Hidden, % NamesArray[A_Index]

Loop, % NamesArray.Length()
	GuiControl, Show, % A_Index

Gui, Show, Autosize

Return

; ------------------- GUI CONTROLS PROGRAMMING ---------------------

; ------------------- WEBSITES ---------------------
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

; ------------------- U: DRIVE FOLDERS ---------------------

PressUDriveProjects:
	IfWinExist, Vibrations (U:)
		WinActivate 
	else 
		run, U:\_PRJ\_PROJECTS
Gui, Destroy
return

PressUDriveReference:
	IfWinExist, Vibrations (U:)
		WinActivate 
	else 
		run, U:\_REFERENCE
Gui, Destroy
return

; ------------------- DYNAMIC BUTTONS ---------------------

PressCopiedLink:
run, % Clipboard,, UseErrorLevel
If ErrorLevel = ERROR
{
	MsgBox, 20496, ERROR, There was an error. Ensure copied address is a viable link.`nCopied Address:`n`n%Clipboard%
	return
}
Gui, Destroy
return

PressDynamicButton:
run, % LinksArray[A_GuiControl],, UseErrorLevel
If ErrorLevel = ERROR
{
	MsgBox, 20496, ERROR, % "There was an error. Ensure linked address is functional.`nLinked Address:`n`n" LinksArray[A_GuiControl]
	return
}
Gui, Destroy
return