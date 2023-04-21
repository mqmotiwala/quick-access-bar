#NoEnv, UseUnsetLocal
#Warn
#SingleInstance force
SendMode Event
SetWorkingDir %A_ScriptDir% 
SetBatchLines, -1
AutoTrim, On

#w::
GoSub, ShowQuickAccessGUI
Return

; ----------------------------------------------------------------------------------------------------------
; ------------------------------------------ QUICK ACCESS BAR ----------------------------------------------
; ----------------------------------------------------------------------------------------------------------

ShowQuickAccessGUI:
; create a shortcut in startup folder
SplitPath, A_ScriptName,,,,FileName
StartUpPath	:= "C:\Users\" . A_Username . "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" . FileName . ".lnk"
IfNotExist, %StartUpPath%
	FileCreateShortcut, %A_ScriptFullPath%, %StartUpPath%

; create arrays
GrabAll := []
NamesArray := []
LinksArray := []   

; parse .ini file contents to populate arrays
IniFilePath = QABInitializationFile.ini
IniRead, ParsedButtonNames, %IniFilePath%, Buttons
IniRead, ParsedLinkAddresses, %IniFilePath%, Links

Loop, Parse, ParsedButtonNames, `n, `r%A_Space%%A_Tab%
	NamesArray.Push(A_LoopField)

Loop, Parse, ParsedLinkAddresses, `n, `r%A_Space%%A_Tab%	
	LinksArray.Push(A_LoopField)

; create Quick Access Bar 
Gui, New,,Quick Access Bar
Gui +AlwaysOnTop +ToolWindow
Gui, Color, E54D31

GeneralButtonWidth 		= 150
HorizontalMargin 		= 10
VerticalMargin 			= 10
EditLinksButtonWidth 	= 120
EditAndAboutButtonGap 	= 5

; calculate About button positioning
AboutButtonY 	:= VerticalMargin
AboutButtonX 	:= EditLinksButtonWidth + HorizontalMargin + EditAndAboutButtonGap
AboutButtonWidth := GeneralButtonWidth - EditLinksButtonWidth - EditAndAboutButtonGap

; add Edit Links and About buttons
EditLinksButtonName := "Edit Dynamic Buttons"
Gui, Add, Button, x%HorizontalMargin% y%VerticalMargin% w%EditLinksButtonWidth% gEditLinks, %EditLinksButtonName%
Gui, Add, Button, x%AboutButtonX% y%AboutButtonY% w%AboutButtonWidth% gPressHelp, ?

; add Website buttons
Gui, Font, cWhite s8 w600
Gui, Add, Text, x%HorizontalMargin%, Websites
Gui, Font

Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressTimeSheet, Timesheet
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSharePoint, SharePoint

; add Folders buttons
Gui, Font, cWhite s8 w600
Gui, Add, Text, x%HorizontalMargin%, Folders
Gui, Font

Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressUDriveProjects, U: Projects
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressUDriveReference, U: Reference

; add Special Buttons
Gui, Font, cWhite s8 w600
Gui, Add, Text, x%HorizontalMargin%, Special Buttons
Gui, Font

Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressCopiedLink, Copied Address
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSnippingTool, Snipping Tool

; add Dynamic Buttons
if(NamesArray.Length() != 0)
{
	Gui, Font, cWhite s8 w600
	Gui, Add, Text, x%HorizontalMargin%, Dynamic Buttons
	Gui, Font
}

Loop, 25
{
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% v%A_Index% gPressDynamicButton Hidden, % NamesArray[A_Index]
}	
Loop, % NamesArray.Length()
	GuiControl, Show, % A_Index

; show Quick Access Bar
Gui, Show, Autosize
Return

; ----------------------------------------------------------------------------------------------------------
; ----------------------------------------- BUTTONS PROGRAMMING --------------------------------------------
; ----------------------------------------------------------------------------------------------------------

; ------------- WEBSITES ---------------

PressTimeSheet:
run, iexplore.exe https://idcsapepp.corp.hatchglobal.com/irj/portal?NavigationTarget=navurl://a4c6004e4c9a20fd30f8437568091a55
Gui, Destroy
return

PressSharePoint:
run, https://hatchengineering.sharepoint.com/
Gui, Destroy
return

; ------------- FOLDERS ---------------

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

; ------------- SPECIAL BUTTONS ---------------

PressCopiedLink:
run, % Clipboard,, UseErrorLevel
If ErrorLevel = ERROR
{
	; remove trailing and leading blank space
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

; ------------- DYNAMIC BUTTONS ---------------

PressDynamicButton:
run, % LinksArray[A_GuiControl],, UseErrorLevel
If ErrorLevel = ERROR
{
	MsgBox, 4112, ERROR, % "There was an error. Ensure linked address is functional.`nLinked Address:`n`n" LinksArray[A_GuiControl]
	return
}
Gui, Destroy
return

; ----------------------------------------------------------------------------------------------------------
; ------------------------------------------ ABOUT BUTTON GUI  ---------------------------------------------
; ----------------------------------------------------------------------------------------------------------

PressHelp:
	; destroy Quick Access Bar
	Gui, Destroy
	
	; declare About GUI texts
	Title 			:= "About Quick Access Bar"
	Sentence1 		:= "Press '" . EditLinksButtonName . "' on the Quick Access Bar to get started. You can add up to 25 dynamic buttons."
	Sentence2 		:= "Any software, file, folder or website can be linked for quick access."
	Sentence3		:= "For general feedback, bug reports or new feature ideas - "
	
	; declare About GUI preferences
	FontName     	:= "Segoe UI"   
	FontSize     	:= 10       
	VerticalMargin 	:= 20            
	LeftMargin   	:= 12       
	RightMargin  	:= 8    
	TextHeight		:= 17
	ButtonWidth  	:= 88      
	ButtonHeight 	:= 26 
	ButtonOffset 	:= 15   
	SS_WHITERECT 	:= 0x0006      

	BottomVerticalMargin := LeftMargin/2
	BottomHeight	:= ButtonHeight + BottomVerticalMargin*2
	Sentence1Y		:= VerticalMargin
	Sentence2Y		:= VerticalMargin + TextHeight
	Sentence3Y		:= VerticalMargin + TextHeight*3
	
	; create a white background
	Gui, Add, Text, x0 y0 %SS_WHITERECT% vWhiteBox
	
	; add About information
	Gui, Font, s%FontSize%, %FontName%              
	Gui, +AlwaysOnTop +ToolWindow -SysMenu      
	Gui, Add, Text, x%LeftMargin% y%Sentence1Y% BackgroundTrans vSentence1, %Sentence1%
	Gui, Add, Text, x%LeftMargin% y%Sentence2Y% BackgroundTrans vSentence2, %Sentence2%
	Gui, Add, Text, x%LeftMargin% y%Sentence3Y% BackgroundTrans vSentence3, %Sentence3%
	
	; acquire sentence1 dimensions for GUI width calculations
	; acquire sentence3 dimensions for link positioning 
	GuiControlGet, Sentence1Dimensions, Pos, Sentence1
	GuiControlGet, Sentence3Dimensions, Pos, Sentence3
	
	; add contact link
	LinkBoxX := LeftMargin + Sentence3DimensionsW
	LinkBoxY := Sentence3DimensionsY

	Gui, Font, Underline
	Gui, Add, Link, x%LinkBoxX% y%LinkBoxY% BackgroundTrans vLinkBox, <a href="mailto:mufaddal.motiwala@hatch.com">contact Mufaddal Motiwala</a>
	Gui, Font, s%FontSize%, %FontName%
	Gui, Font, norm
	
	; acquire link dimensions to position the following period
	GuiControlGet, LinkBoxDimensions, Pos, LinkBox
	PeriodX := LinkBoxDimensionsX + LinkBoxDimensionsW
	PeriodY := LinkBoxDimensionsY
	Gui, Add, Text, x%PeriodX% y%PeriodY% BackgroundTrans, .

	; resize white background to fit all of GUI
	GuiWidth := LeftMargin + Sentence1DimensionsW + ButtonOffset + RightMargin + 1
	WhiteBoxHeight := Sentence3DimensionsY + Sentence3DimensionsH + VerticalMargin                                       

	GuiControl, Move, WhiteBox, w%GuiWidth% h%WhiteBoxHeight%
 
	; place OK button
	ButtonX := GuiWidth - RightMargin - ButtonWidth                 
	ButtonY := WhiteBoxHeight + BottomVerticalMargin
	
	Gui, Add, Button, x%ButtonX% y%ButtonY% w%ButtonWidth% h%ButtonHeight% Default, OK
	GuiControl, Focus, OK
	
	; show About GUI
	GuiHeight := WhiteBoxHeight + BottomHeight
	Gui, Show, w%GuiWidth% h%GuiHeight%, %Title%
Return

; Pressing OK on About GUI will return to Quick Access Bar
ButtonOK:
Gui, Destroy
GoSub, ShowQuickAccessGUI
return

GuiEscape:
Gui, Destroy
Return

; ----------------------------------------------------------------------------------------------------------
; -------------------------------------------- LINKS EDITOR  -----------------------------------------------
; ----------------------------------------------------------------------------------------------------------

EditLinks:
	; destroy Quick Access Bar
	Gui, Destroy
	
	; declare Edit Links GUI preferences
	Offsets				= 5
	TextsGap			= 8
	TextsHeight			= 22
	MaxButtons 			= 25
	NamesEditWidth 		= 150
	LinksEditWidth 		= 600
	SaveButtonWidth 	= 50
	SaveButtonHeight 	= 25
	
	NamesEditX	:= Offsets + 1
	LinksEditX	:= Offsets + NamesEditWidth + 3
	TextsY		:= Offsets + SaveButtonHeight + TextsGap
	EditsY 		:= Offsets + SaveButtonHeight + TextsHeight
	
	; create Edit Links GUI
	Gui, New,, Add Dynamic Buttons
	Gui +AlwaysOnTop +ToolWindow -SysMenu
	Gui, Font, s9
	Gui, Color, eaeaea
	
	; set border margins
	Gui, Margin, 5, 5
	
	; add Save button
	Gui, Add, Button, x%Offsets% y%Offsets% w%SaveButtonWidth% h%SaveButtonHeight% gPressSave, Save
	
	; add description texts edit fields with desired preferences 
	Gui, Add, Text, x%NamesEditX% y%TextsY%, Insert Button Names Below
	Gui, Add, Text, x%LinksEditX% y%TextsY%, Insert Corresponding Links Below

	Gui, Add, Edit, vNamesEdit x%NamesEditX% y%EditsY% WantReturn -VScroll -Wrap 0x2 w%NamesEditWidth% r%MaxButtons%	
	Gui, Add, Edit, vLinksEdit x%LinksEditX% y%EditsY% WantReturn -VScroll -Wrap w%LinksEditWidth% r%MaxButtons%
	
	; insert .ini file parsed contents into the edit fields
	GuiControl,, NamesEdit, %ParsedButtonNames%
	GuiControl,, LinksEdit, %ParsedLinkAddresses%
	
	; show Edit Links GUI
	Gui, Show, Center AutoSize
return

PressSave:
	; save edit fields contents in .ini file
	GuiControlGet, NamesEdit
	GuiControlGet, LinksEdit
	IniDelete, %IniFilePath%, Buttons
	IniDelete, %IniFilePath%, Links
	IniWrite, %NamesEdit%, %IniFilePath%, Buttons
	IniWrite, %LinksEdit%, %IniFilePath%, Links
	
	; return to Quick Access Bar
	Gui, Destroy
	GoSub, ShowQuickAccessGUI
return