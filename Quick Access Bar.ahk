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

QABWidth := GeneralButtonWidth + 2*HorizontalMargin

; calculate About button positioning
AboutButtonY 	:= VerticalMargin
AboutButtonX 	:= EditLinksButtonWidth + HorizontalMargin + EditAndAboutButtonGap
AboutButtonWidth := GeneralButtonWidth - EditLinksButtonWidth - EditAndAboutButtonGap

; add Edit Links and About buttons
EditLinksButtonName := "Edit Dynamic Buttons"
Gui, Add, Button, x%HorizontalMargin% y%VerticalMargin% w%EditLinksButtonWidth% gEditLinks, %EditLinksButtonName%
Gui, Add, Button, x%AboutButtonX% y%AboutButtonY% w%AboutButtonWidth% gPressHelp, ?

; add Hatch specific buttons
Gui, Font, cWhite s8 w600
Gui, Add, Text, x%HorizontalMargin%, Hatch
Gui, Font

Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressTimeSheet, Timesheet
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSharePoint, SharePoint
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressKIRC, KIRC
Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressVPN, VPN

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
ActiveMonitor 	:= GetCurrentMonitorIndex()
QABX 			:= CoordXCenterScreen(QABWidth, ActiveMonitor)

Gui, Show, x%QABX% yCenter Autosize
Return

CloseQAB:
IniRead, CurrentCheckbox1Status, %IniFilePath%, Checkboxes, CheckBox1
if (CurrentCheckbox1Status)
	Gui, Destroy
return

; ----------------------------------------------------------------------------------------------------------
; ----------------------------------------- BUTTONS PROGRAMMING --------------------------------------------
; ----------------------------------------------------------------------------------------------------------

; ------------- HATCH SPECIFIC ---------------

PressTimeSheet:
run, iexplore.exe https://idcsapepp.corp.hatchglobal.com/irj/portal?NavigationTarget=navurl://a4c6004e4c9a20fd30f8437568091a55,, UseErrorLevel
If ErrorLevel = ERROR
	GoSub, GenericErrorLevel
GoSub, CloseQAB
return

PressSharePoint:
run, https://hatchengineering.sharepoint.com/,, UseErrorLevel
If ErrorLevel = ERROR
	GoSub, GenericErrorLevel
GoSub, CloseQAB
return

PressKIRC:
run, https://1415.sydneyplus.com/Hatch_SE/portal.aspx,, UseErrorLevel
If ErrorLevel = ERROR
	GoSub, GenericErrorLevel
GoSub, CloseQAB
return

PressVPN:
run, C:\Program Files (x86)\F5 VPN\f5fpclientW.exe,, UseErrorLevel
If ErrorLevel = ERROR
	GoSub, GenericErrorLevel
GoSub, CloseQAB
return

GenericErrorLevel:
	MsgBox, 4116, ERROR, Uh oh, there was an unusual error.`n`nWould you like to send me an email? I'll sort this out for ya!
	IfMsgBox, Yes
		run, mailto:mufaddal.motiwala@hatch.com
	GoSub, ShowQuickAccessGUI
return

; ------------- FOLDERS ---------------

PressUDriveProjects:
	IfWinExist, _PRJ
		WinActivate 
	else 
		{
			run, U:\_PRJ\_PROJECTS,, UseErrorLevel
			If ErrorLevel = ERROR
				{
					MsgBox, 4112, ERROR, Uh oh, I think you need to connect to Hatch network first.
					GoSub, PressVPN
				}
		}
GoSub, CloseQAB
return

PressUDriveReference:
	IfWinExist, _REFERENCE
		WinActivate 
	else 
		{
			run, U:\_REFERENCE,, UseErrorLevel
			If ErrorLevel = ERROR
				{
					MsgBox, 4112, ERROR, Uh oh, I think you need to connect to Hatch network first.
					GoSub, PressVPN
				}
		}
GoSub, CloseQAB
return

; ------------- SPECIAL BUTTONS ---------------

^+c::
{
	IniRead, CurrentCheckbox2Status, %IniFilePath%, Checkboxes, CheckBox2
	if (CurrentCheckbox2Status)
		GoSub, PressCopiedLink
}
return

PressCopiedLink:
run, % Clipboard,, UseErrorLevel
If ErrorLevel = ERROR
{
	; remove trailing and leading blank space
	StrippedClipboard := RegExReplace(clipboard, "[`r`n`t\s]+$")
	
	; if clipboard exists and is not a direct link, Google search clipboard contents
	if (StrippedClipboard)
		run, https://www.google.com/#q=%clipboard%,, UseErrorLevel
	else
		{
			MsgBox, 4112, ERROR, You don't have anything copied!
			return
		}
}
GoSub, CloseQAB
return

^+s::
{
	IniRead, CurrentCheckbox3Status, %IniFilePath%, Checkboxes, CheckBox3
	if (CurrentCheckbox3Status)
		GoSub, PressSnippingTool
}
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
GoSub, CloseQAB
return

; ----------------------------------------------------------------------------------------------------------
; ------------------------------------------ ABOUT BUTTON GUI  ---------------------------------------------
; ----------------------------------------------------------------------------------------------------------

PressHelp:
	; destroy Quick Access Bar
	Gui, Destroy
	
	; declare About GUI preferences
	FontName     	:= "Segoe UI"   
	FontSize     	:= 10       
	VerticalMargin 	:= 10            
	LeftMargin   	:= 12       
	RightMargin  	:= 12    
	TextHeight		:= 17
	ButtonWidth  	:= 88      
	ButtonHeight 	:= 26 
	ButtonOffset 	:= 15   
	SS_WHITERECT 	:= 0x0006      

	ExtraRowSpacing	:= 2*TextHeight
	BottomVerticalMargin := LeftMargin/2
	BottomHeight	:= ButtonHeight + BottomVerticalMargin*2
	Title 			:= "Quick Access Bar"
	
	; create a white background
	Gui, Add, Text, x0 y0 %SS_WHITERECT% vWhiteBox

; --- ABOUT QUICK ACCESS BAR [AREA 1]
	; section heading
	Area1Y := VerticalMargin
	Gui, Font, s%FontSize% w600, %FontName%
	Gui, Add, Text, x%LeftMargin% y%Area1Y% BackgroundTrans, About Quick Access Bar
	Gui, Font, Norm

	; add About information
	Sentence1Y		:= Area1Y + TextHeight
	Sentence2Y		:= Sentence1Y + TextHeight
	Sentence3Y		:= Sentence2Y + ExtraRowSpacing
	
	Sentence1 		:= "Press '" . EditLinksButtonName . "' on the Quick Access Bar to get started. You can add up to 25 dynamic buttons."
	Sentence2 		:= "Any software, file, folder or website can be linked for quick access."
	Sentence3		:= "For general feedback, bug reports or new feature ideas - "
                
	Gui, Add, Text, x%LeftMargin% y%Sentence1Y% BackgroundTrans vSentence1, %Sentence1%
	Gui, Add, Text, x%LeftMargin% y%Sentence2Y% BackgroundTrans vSentence2, %Sentence2%
	Gui, Add, Text, x%LeftMargin% y%Sentence3Y% BackgroundTrans vSentence3, %Sentence3%

	; add developer contact information
	GuiControlGet, Sentence3Dimensions, Pos, Sentence3
	
	LinkBoxX := LeftMargin + Sentence3DimensionsW
	LinkBoxY := Sentence3DimensionsY

	Gui, Font, Underline
	Gui, Add, Link, x%LinkBoxX% y%LinkBoxY% BackgroundTrans vLinkBox, <a href="mailto:mufaddal.motiwala@hatch.com">contact Mufaddal Motiwala</a>
	Gui, Font, s%FontSize%, %FontName% 
	Gui, Font, Norm

	; acquire link dimensions to position the following period
	GuiControlGet, LinkBoxDimensions, Pos, LinkBox
	PeriodX := LinkBoxDimensionsX + LinkBoxDimensionsW
	PeriodY := LinkBoxDimensionsY
	Gui, Add, Text, x%PeriodX% y%PeriodY% BackgroundTrans, .
	
; --- SPECIAL BUTTONS EXPLANATION [AREA 2]
	; section heading
	Area2Y := PeriodY + ExtraRowSpacing
	Gui, Font, s%FontSize% w600, %FontName%
	Gui, Add, Text, x%LeftMargin% y%Area2Y% BackgroundTrans, Special Buttons Explanation
	Gui, Font, Norm

	; add Explanations
	Sentence4Y 	:= Area2Y + TextHeight
	Sentence5Y	:= Sentence4Y + TextHeight + ExtraRowSpacing
	
	Sentence4 	:= "Copied Address will route to the currently copied link.`nIf the copied content is not a link, Copied Address will perform a Google search."
	Sentence5	:= "Snipping Tool will begin a new snip with your default snip settings." 

	Gui, Add, Text, x%LeftMargin% y%Sentence4Y% BackgroundTrans vSentence4, %Sentence4%
	Gui, Add, Text, x%LeftMargin% y%Sentence5Y% BackgroundTrans vSentence5, %Sentence5%

; --- USER SETTINGS [AREA 3]
	; section heading
	Area3Y := Sentence5Y + ExtraRowSpacing
	Gui, Font, s%FontSize% w600, %FontName%
	Gui, Add, Text, x%LeftMargin% y%Area3Y% BackgroundTrans, User Settings
	Gui, Font, Norm

	; add checkboxes
	CheckBox1Y := Area3Y + TextHeight
	CheckBox2Y := CheckBox1Y + TextHeight
	CheckBox3Y := CheckBox2Y + TextHeight
	
	IniRead, CurrentCheckbox1Status, %IniFilePath%, Checkboxes, CheckBox1
	IniRead, CurrentCheckbox2Status, %IniFilePath%, Checkboxes, CheckBox2
	IniRead, CurrentCheckbox3Status, %IniFilePath%, Checkboxes, CheckBox3
	
	Gui, Add, CheckBox, x%LeftMargin% y%CheckBox1Y% Checked%CurrentCheckBox1Status% vCheckBox1Button gToggleCheckBox
	Gui, Add, CheckBox, x%LeftMargin% y%CheckBox2Y% Checked%CurrentCheckBox2Status% vCheckBox2Button gToggleCheckBox
	Gui, Add, CheckBox, x%LeftMargin% y%CheckBox3Y% Checked%CurrentCheckBox3Status% vCheckBox3Button gToggleCheckBox
	
	; set checkboxes width to match their height
	GuiControlGet, CheckBoxDim, Pos, CheckBox1Button
	GuiControl, Move, CheckBox1Button, w%CheckBoxDimH%
	GuiControl, Move, CheckBox2Button, w%CheckBoxDimH%
	GuiControl, Move, CheckBox3Button, w%CheckBoxDimH%
	
	; add checkbox description texts
	; set initial text widths sufficiently large
	CheckBoxTextX 	:= LeftMargin + CheckBoxDimH + 10
	Gui, Add, Text, x%CheckboxTextX% y%CheckBox1Y% w1000 BackgroundTrans vCheckbox1Text
	Gui, Add, Text, x%CheckboxTextX% y%CheckBox2Y% w1000 BackgroundTrans vCheckbox2Text
	Gui, Add, Text, x%CheckboxTextX% y%CheckBox3Y% w1000 BackgroundTrans vCheckbox3Text

	; alter checkbox text values as per their current status
	CheckBox1Status1Text 	:= "Quick Access Bar will close automatically after each use."
	CheckBox1Status0Text 	:= "Quick Access Bar will remain visible after use."
	
	CheckBox2Status1Text 	:= "Ctrl + Shift + C will trigger Copied Address special button."
	CheckBox2Status0Text 	:= "No shortcut set for Copied Address special button."	 

	CheckBox3Status1Text 	:= "Ctrl + Shift + S will trigger Snipping Tool special button."
	CheckBox3Status0Text 	:= "No shortcut set for Snipping Tool special button."
	
	if (CurrentCheckbox1Status)
		GuiControl, Text, Checkbox1Text, %CheckBox1Status1Text%
	else
		GuiControl, Text, Checkbox1Text, %CheckBox1Status0Text%	
		
	if (CurrentCheckbox2Status)
		GuiControl, Text, Checkbox2Text, %CheckBox2Status1Text%
	else
		GuiControl, Text, Checkbox2Text, %CheckBox2Status0Text%	
		
	if (CurrentCheckbox3Status)
		GuiControl, Text, Checkbox3Text, %CheckBox3Status1Text%
	else
		GuiControl, Text, Checkbox3Text, %CheckBox3Status0Text%	
	
; --- OVERALL ABOUT GUI CONFIGURATION & OK BUTTON [AREA 4]
	; acquire relevant control dimensions for GUI width and height calculations
	GuiControlGet, WidestControlDim, Pos, Sentence1
	GuiControlGet, LowestControlDim, Pos, CheckBox3Button
	
	; resize white background to fit all of GUI
	WhiteBoxHeight 	:= LowestControlDimY + LowestControlDimH + VerticalMargin                                       
	GuiWidth 		:= LeftMargin + WidestControlDimW + RightMargin
	
	GuiControl, Move, WhiteBox, w%GuiWidth% h%WhiteBoxHeight%
 
	; place OK button
	ButtonX := GuiWidth - RightMargin - ButtonWidth
	ButtonY := WhiteBoxHeight + BottomVerticalMargin
	
	Gui, Add, Button, x%ButtonX% y%ButtonY% w%ButtonWidth% h%ButtonHeight% Default, OK
	GuiControl, Focus, OK
	
	; show About GUI
	GuiHeight := WhiteBoxHeight + BottomHeight
	
	Gui, +AlwaysOnTop +ToolWindow -SysMenu 
	
	ActiveMonitor 	:= GetCurrentMonitorIndex()
	AboutGuiX		:= CoordXCenterScreen(GuiWidth, ActiveMonitor)

	Gui, Show, x%AboutGuiX% yCenter w%GuiWidth% h%GuiHeight%, %Title%
Return

; Pressing OK on About GUI will return to Quick Access Bar
ButtonOK:
	Gui, Destroy
	GoSub, ShowQuickAccessGUI
return

GuiEscape:
Gui, Destroy
Return

ToggleCheckbox:
	; update all control variables
	Gui, Submit, NoHide
	
	; dynamically update checkbox text
	if (CheckBox1Button)
		GuiControl, Text, CheckBox1Text, %CheckBox1Status1Text%
	else
		GuiControl, Text, CheckBox1Text, %CheckBox1Status0Text%
	
	if (CheckBox2Button)
		GuiControl, Text, CheckBox2Text, %CheckBox2Status1Text%
	else
		GuiControl, Text, CheckBox2Text, %CheckBox2Status0Text%
	
	if (CheckBox3Button)
		GuiControl, Text, CheckBox3Text, %CheckBox3Status1Text%
	else
		GuiControl, Text, CheckBox3Text, %CheckBox3Status0Text%
	
	; write Checkbox status to .ini file
	IniWrite, %CheckBox1Button%, %IniFilePath%, Checkboxes, CheckBox1
	IniWrite, %CheckBox2Button%, %IniFilePath%, Checkboxes, CheckBox2
	IniWrite, %CheckBox3Button%, %IniFilePath%, Checkboxes, CheckBox3
return
	
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
	
	EditerGuiWidth := NamesEditWidth + LinksEditWidth + 2*Offsets + 4
	
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
	Gui, Margin, %Offsets%, %Offsets%
	
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
	ActiveMonitor 	:= GetCurrentMonitorIndex()
	EditerGuiX		:= CoordXCenterScreen(EditerGuiWidth, ActiveMonitor)
	
	Gui, Show, x%EditerGuiX% yCenter AutoSize, %EditLinksButtonName%
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

; ----------------------------------------------------------------------------------------------------------
; ------------------------------------ FIND ACTIVE WINDOW FUNCTIONS ----------------------------------------
; ----------------------------------------------------------------------------------------------------------

GetCurrentMonitorIndex()
	{
		; get current mouse coordinates
		CoordMode, Mouse, Screen
		MouseGetPos, mx, my
		
		; get number of monitors and identify currently active monitor screen
		SysGet, NumOfMonitors, 80
		Loop %NumOfMonitors%
			{
				; get monitor screen coordinates for all screens
				SysGet, MonitorCoordinate, Monitor, %A_Index%
				if (MonitorCoordinateLeft <= mx && mx <= MonitorCoordinateRight && MonitorCoordinateTop <= my && my <= MonitorCoordinateBottom)
					Return A_Index
			}
		Return 1
	}

CoordXCenterScreen(WidthOfGUI, ScreenNumber)
	{
		; get monitor screen coordinates for specified monitor
		SysGet, SpecifiedMonitor, Monitor, %ScreenNumber%
		
		; return horizontal center of specified monitor screen
		return (SpecifiedMonitorRight - SpecifiedMonitorLeft - WidthOfGUI)/2 + SpecifiedMonitorLeft
	}