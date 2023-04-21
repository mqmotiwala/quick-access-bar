#NoEnv, UseUnsetLocal
#Warn
#SingleInstance force
SendMode Event
SetWorkingDir %A_ScriptDir% 
SetBatchLines, -1
AutoTrim, On

; Creates a GUI for quick access to frequently used resources

#w::
GoSub, ShowQuickAccessGUI
Return

ShowQuickAccessGUI:
IniFilePath = QABInitializationFile.ini
IniRead, ParsedButtonNames, %IniFilePath%, Buttons
IniRead, ParsedLinkAddresses, %IniFilePath%, Links

; Create Arrays
GrabAll := []
NamesArray := []
LinksArray := []   

; Parse .ini file contents line by line 
Loop, Parse, ParsedButtonNames, `n, `r%A_Space%%A_Tab%
	NamesArray.Push(A_LoopField)

Loop, Parse, ParsedLinkAddresses, `n, `r%A_Space%%A_Tab%	
	LinksArray.Push(A_LoopField)

Gui, New,,Quick Access Bar
Gui +AlwaysOnTop +ToolWindow
Gui, Color, E54D31

GeneralButtonWidth = 150
HorizontalMargin = 10
EditLinksButtonY = 10
EditLinksButtonWidth = 120
GapBetweenEditAndAboutButtons = 5
HelpButtonY := EditLinksButtonY
HelpButtonX := EditLinksButtonWidth + HorizontalMargin + GapBetweenEditAndAboutButtons
HelpButtonWidth := GeneralButtonWidth - EditLinksButtonWidth - GapBetweenEditAndAboutButtons

EditLinksButtonName = Edit Dynamic Buttons
Gui, Add, Button, x%HorizontalMargin% y%EditLinksButtonY% w%EditLinksButtonWidth% gEditLinks, %EditLinksButtonName%
Gui, Add, Button, x%HelpButtonX% y%HelpButtonY% w%HelpButtonWidth% gPressHelp, ?

Gui, Font, cWhite s8 w600
Gui, Add, Text, x%HorizontalMargin%, Websites
Gui, Font

Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressTimeSheet, Timesheet
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSharePoint, SharePoint

Gui, Font, cWhite s8 w600
Gui, Add, Text, x%HorizontalMargin%, Folders
Gui, Font

Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressUDriveProjects, U: Projects
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressUDriveReference, U: Reference

Gui, Font, cWhite s8 w600
Gui, Add, Text, x%HorizontalMargin%, Special Buttons
Gui, Font

Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressCopiedLink, Copied Address
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSnippingTool, Snipping Tool

if(NamesArray.Length() != 0)
{
	Gui, Font, cWhite s8 w600
	Gui, Add, Text, x%HorizontalMargin%, Dynamic Buttons
	Gui, Font
}

Loop, 25
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% v%A_Index% gPressDynamicButton Hidden, % NamesArray[A_Index]

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

PressSharePoint:
run, https://hatchengineering.sharepoint.com/
Gui, Destroy
return

; ------------------- FOLDERS ---------------------

PressUDriveProjects:
	IfWinExist, _PRJ
		WinActivate 
	else 
		run, U:\_PRJ\_PROJECTS
Gui, Destroy
return

PressUDriveReference:
	IfWinExist, _REFERENCE
		WinActivate 
	else 
		run, U:\_REFERENCE
Gui, Destroy
return

; ------------------- SPECIAL BUTTONS --------------------
PressCopiedLink:
run, % Clipboard,, UseErrorLevel
If ErrorLevel = ERROR
{
	StrippedClipboard := RegExReplace(clipboard, "[`r`n`t]+$")
	if (StrippedClipboard)
		MsgBox, 4112, ERROR, There was an error. Ensure copied address is a viable link.`nCopied Address:`n`n%Clipboard%
	else
		MsgBox, 4112, ERROR, You don't have anything copied!
	return
}
Gui, Destroy
return

PressSnippingTool:
{
	Gui, Destroy
	run, SnippingTool.exe
	WinWaitActive, Snipping Tool
	Send, ^n
}
Return

; ------------------- DYNAMIC BUTTONS ---------------------
PressDynamicButton:
run, % LinksArray[A_GuiControl],, UseErrorLevel
If ErrorLevel = ERROR
{
	MsgBox, 4112, ERROR, % "There was an error. Ensure linked address is functional.`nLinked Address:`n`n" LinksArray[A_GuiControl]
	return
}
Gui, Destroy
return

; ------------------- ABOUT BUTTON  ---------------------
PressHelp:
	Gui, Destroy
	Title := "About Quick Access Bar"
	Text  := "Press '" . EditLinksButtonName . "' on the Quick Access Bar to get started. You can add up to 25 dynamic buttons.`nAny software, file, folder or website can be linked for quick access.`n`nFor general feedback, bug reports or new feature ideas - "

	FontName     := "Segoe UI"   
	FontSize     := 10       
	Gap          := 20            
	LeftMargin   := 12       
	RightMargin  := 8         
	ButtonWidth  := 88      
	ButtonHeight := 26 
	ButtonOffset := 15   
	SS_WHITERECT := 0x0006      

	Gui, Add, Text, x0 y0 %SS_WHITERECT% vWhiteBox 
	BottomGap := LeftMargin/2
	BottomHeight := ButtonHeight+LeftMargin
	Gui, Font, s%FontSize%, %FontName%              
	Gui, +ToolWindow -MinimizeBox -MaximizeBox         
	Gui, Add, Text, x%LeftMargin% y%Gap% BackgroundTrans vTextBox, %Text%

	Gui, Font, Underline
	Gui, Add, Link, x350 y70, <a href="mailto:mufaddal.motiwala@hatch.com">contact Mufaddal Motiwala</a>
	Gui, Font, s%FontSize%, %FontName%
	Gui, Font, norm
	Gui, Add, Text, x510 y70, .

	GuiControlGet, Size, Pos, TextBox                                       
	GuiWidth := LeftMargin+SizeW+ButtonOffset+RightMargin+1                      
	WhiteBoxHeight := SizeY+SizeH+Gap                                       

	GuiControl, Move, WhiteBox, w%GuiWidth% h%WhiteBoxHeight%   
	ButtonX := GuiWidth-RightMargin-ButtonWidth                 
	ButtonY := WhiteBoxHeight+BottomGap

	Gui, Add, Button, x%ButtonX% y%ButtonY% w%ButtonWidth% h%ButtonHeight% Default, OK
	GuiControl, Focus, OK
	GuiHeight := WhiteBoxHeight+BottomHeight

	Gui, Show, w%GuiWidth% h%GuiHeight%, %Title%                
	Gui, -ToolWindow                                      
Return

ButtonOK:
Gui, Destroy
GoSub, ShowQuickAccessGUI
return

GuiClose:
GuiEscape:
Gui, Destroy
Return

; ------------------- LINKS EDITOR ---------------------
EditLinks:
	Gui, Destroy
	Margin = 5
	TextHeight = 22
	SaveButtonWidth = 50
	SaveButtonHeight = 25
	NamesEditX := Margin + 1
	NamesEditWidth = 150
	LinksEditX := Margin + NamesEditWidth + 3
	LinksEditWidth = 600
	EditsY := Margin + SaveButtonHeight + TextHeight
	MaxButtons = 25
	
	Gui, New,, Add Dynamic Buttons
	Gui +AlwaysOnTop +ToolWindow
	Gui, Font, s9
	Gui, Color, eaeaea
	Gui, Margin, 5, 5
	Gui, Add, Button, x%Margin% y%Margin% w%SaveButtonWidth% h%SaveButtonHeight% gPressSave, Save
	Gui, Add, Text,, %A_Space%Insert Button Names Below        Insert Corresponding Links Below
	Gui, Add, Edit, vLinksEdit x%LinksEditX% y%EditsY% WantTab WantReturn -VScroll -Wrap w%LinksEditWidth% r%MaxButtons%
	Gui, Add, Edit, vNamesEdit x%NamesEditX% y%EditsY% WantTab WantReturn -VScroll -Wrap 0x2 w%NamesEditWidth% r%MaxButtons%

	GuiControl,, NamesEdit, %ParsedButtonNames%
	GuiControl,, LinksEdit, %ParsedLinkAddresses%
	
	Gui, Show, Center AutoSize
return

PressSave:
	GuiControlGet, NamesEdit
	GuiControlGet, LinksEdit
	IniDelete, %IniFilePath%, Buttons
	IniDelete, %IniFilePath%, Links
	IniWrite, %NamesEdit%, %IniFilePath%, Buttons
	IniWrite, %LinksEdit%, %IniFilePath%, Links
	
	Gui, Destroy
	GoSub, ShowQuickAccessGUI
return