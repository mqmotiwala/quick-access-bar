#NoEnv, UseUnsetLocal
#Warn
#SingleInstance force
SendMode Event
SetWorkingDir %A_ScriptDir% 
SetBatchLines, -1
AutoTrim, On

; Creates a GUI for quick access to frequently used resources

#w::
; Create Arrays
GrabAll := []
NamesArray := []
LinksArray := []   

FileRead, CustomLinksFileContents, CustomLinks.txt
Loop, Parse, CustomLinksFileContents, `n;, `r	    		; Parse variable contents line by line
    GrabAll.Push(A_LoopField)    				; Store lines in array
	
Loop, % GrabAll.MaxIndex()
{
	if (mod(A_Index, 2) == 1)
		NamesArray[Floor(A_Index/2) + 1] := GrabAll[A_Index]
	if (mod(A_Index, 2) == 0)
		LinksArray[Floor(A_Index/2)] := GrabAll[A_Index]
}

Loop, % LinksArray.MaxIndex()
	MsgBox, % LinksArray[A_Index]
	;LinksArray[A_Index] := RegExReplace(LinksArray[A_Index],"A)\s")

Gui, New,,Quick Access Bar
Gui +AlwaysOnTop +ToolWindow
Gui, Color, E54D31

Gui, Font, cWhite s8 w600
Gui, Add, Text,, Websites
Gui, Font

Gui, Add, Button, gPressTimeSheet w150, Timesheet
Gui, Add, Button, gPressSharePoint w150, SharePoint
Gui, Add, Button, gPressKIRC w150, KIRC

Gui, Font, cWhite s8 w600
Gui, Add, Text,, U: Drive Folders
Gui, Font

Gui, Add, Button, gPressUDriveProjects w150, U:\Projects
Gui, Add, Button, gPressUDriveReference w150, U:\Reference

Gui, Font, cWhite s8 w600
Gui, Add, Text,, Custom Buttons
Gui, Font

Gui, Add, Button, gPressCopiedLink w150, Copied Address
Gui, Add, Button, gPressButton1 w150, % NamesArray[1]
Gui, Add, Button, gPressButton2 w150, % NamesArray[2]
Gui, Add, Button, gPressButton3 w150, % NamesArray[3]
Gui, Add, Button, gPressButton4 w150, % NamesArray[4]
Gui, Add, Button, gPressButton5 w150, % NamesArray[5]
Gui, Add, Button, gPressButton6 w150, % NamesArray[6]
Gui, Add, Button, gPressButton7 w150, % NamesArray[7]
Gui, Add, Button, gPressButton8 w150, % NamesArray[8]

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

PressUDriveProjects:
	IfWinExist, _PROJECTS
		WinActivate 
	else 
		run, Explorer U:\_PRJ\_PROJECTS 
Gui, Destroy
return

PressUDriveReference:
	IfWinExist, _REFERENCE
		WinActivate 
	else 
		run, Explorer U:\_REFERENCE
Gui, Destroy
return

PressCopiedLink:
run, % Clipboard
Gui, Destroy
return

PressButton1:
	run, % LinksArray[1]
Gui, Destroy
return

PressButton2:
	run, % LinksArray[2]
Gui, Destroy
return

PressButton3:
	run, % LinksArray[3]
Gui, Destroy
return

PressButton4:
	run, % LinksArray[4]
Gui, Destroy
return

PressButton5:
	run, % LinksArray[5]
Gui, Destroy
return

PressButton6:
	run, % LinksArray[6]
Gui, Destroy
return

PressButton7:
	run, % LinksArray[7]
Gui, Destroy
return

PressButton8:
	run, % LinksArray[8]
Gui, Destroy
return