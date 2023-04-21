#NoEnv, UseUnsetLocal
#Warn, UseUnsetGlobal, Off
#SingleInstance force
SendMode Event
SetWorkingDir %A_ScriptDir% 
SetBatchLines, -1
AutoTrim, On
Menu, Tray, Add, Show Quick Access Bar, TrayMenuShowQAB
Menu, Tray, Add, Dynamic Buttons, TrayMenuDynamicButtons
Menu, Tray, Add, About, TrayMenuAbout
Menu, Tray, Add
Menu, Tray, Add, Check for Updates, TrayMenuUpdate
Menu, Tray, Add, Run Tutorial, TrayMenuTutorial
Menu, Tray, Add, Exit, TrayMenuExit
Menu, Tray, NoStandard
PrerequisiteTasks()
return

#w::
GoSub, ShowQuickAccessGUI
return

; ----------------------------------------------------------------------------------------------------------
; ------------------------------------------ QUICK ACCESS BAR ----------------------------------------------
; ----------------------------------------------------------------------------------------------------------

ShowQuickAccessGUI:
	; delete and create a new shortcut in startup folder
	; this is to make sure that a new shortcut is created even after updating
	SplitPath, A_ScriptName,,,,FileName
	StartUpPath	:= "C:\Users\" . A_Username . "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" . FileName . ".lnk"
	FileDelete, %StartUpPath%
	FileCreateShortcut, %A_ScriptFullPath%, %StartUpPath%
	
	; create arrays
	GrabAll := []
	NamesArray := []
	LinksArray := []   

	; parse .ini file contents to populate arrays
	IniRead, ParsedButtonNames, %LocalIniFilePath%, Buttons
	IniRead, ParsedLinkAddresses, %LocalIniFilePath%, Links

	Loop, Parse, ParsedButtonNames, `n, `r%A_Space%%A_Tab%
		NamesArray.Push(A_LoopField)

	Loop, Parse, ParsedLinkAddresses, `n, `r%A_Space%%A_Tab%	
		LinksArray.Push(A_LoopField)
		
	; create Quick Access Bar 
	Gui, QAB:New

	GeneralButtonWidth 		= 150
	HorizontalMargin 		= 10
	VerticalMargin 			= 10
	EditLinksButtonWidth 	= 120
	EditAndAboutButtonGap 	= 5
	QABFontSize				= 8
	BoldWeight				= 600

	QABWidth := GeneralButtonWidth + 2*HorizontalMargin

	; calculate About button positioning
	AboutButtonY 	:= VerticalMargin
	AboutButtonX 	:= EditLinksButtonWidth + HorizontalMargin + EditAndAboutButtonGap
	AboutButtonWidth := GeneralButtonWidth - EditLinksButtonWidth - EditAndAboutButtonGap

	; add Edit Links and About buttons
	; defaults to 'Add' but will update to 'Edit' if dynamic buttons exist
	EditLinksButtonName := "Add Dynamic Buttons"
	Gui, Add, Button, x%HorizontalMargin% y%VerticalMargin% w%EditLinksButtonWidth% gEditLinks vEditLinksButton, %EditLinksButtonName%
	Gui, Add, Button, x%AboutButtonX% y%AboutButtonY% w%AboutButtonWidth% gPressAbout, ?

	; add Hatch Website buttons
	Gui, Font, cWhite s%QABFontSize% w%BoldWeight%
	Gui, Add, Text, x%HorizontalMargin%, Hatch
	Gui, Font 

	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressTimeSheet, Timesheet
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSharePoint, SharePoint
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressDashboard, ROAM Dashboard
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSafetyObservation, Safety Observation 

	; add Special Buttons
	Gui, Font, cWhite s%QABFontSize% w%BoldWeight%
	Gui, Add, Text, x%HorizontalMargin%, Special Buttons
	Gui, Font
	
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressUnitConverter, Unit Converter
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSnippingTool, Snipping Tool
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressCopiedLink, %CopiedAddressSpecialButtonName%
	
	; add Dynamic Buttons
	if(NamesArray.Length() != 0)
	{
		Gui, Font, cWhite s%QABFontSize% w%BoldWeight%
		Gui, Add, Text, x%HorizontalMargin%, Dynamic Buttons
		Gui, Font
		
		EditLinksButtonName := "Edit Dynamic Buttons"
		GuiControl,, EditLinksButton, %EditLinksButtonName%
	}

	Loop, 25
		Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% v%A_Index% gPressDynamicButton Hidden, % NamesArray[A_Index]

	Loop, % NamesArray.Length()
		GuiControl, Show, % A_Index

	; show Quick Access Bar
	; identify monitor screen 
	ActiveMonitor 	:= GetCurrentMonitorIndex()
	QABX 			:= CoordXCenterScreen(QABWidth, ActiveMonitor)

	; move other GUIs as needed
	WinMove, %WelcomeTitle%, , % QABX + 4*HorizontalMargin + GeneralButtonWidth
	WinMove, %UpdaterTitle%, , % QABX + 4*HorizontalMargin + GeneralButtonWidth 
	
	; set GUI preferences; E0x40000 forces taskbar to show icon
	Gui +AlwaysOnTop +ToolWindow +E0x840000
	Gui, Color, E54D31
	Gui, Show, x%QABX% yCenter Autosize, %QABTitle%
Return

GuiEscape:
	Gui, QAB:Destroy
Return

CloseQAB:
	IniRead, CurrentCheckbox1Status, %LocalIniFilePath%, Checkboxes, CheckBox1
	if (CurrentCheckbox1Status)
		Gui, QAB:Destroy
	
	UpdateLocalUsageStats()
	
	; refresh About GUI if it exists 
	if WinExist(AboutTitle)
		GoSub, PressAbout
return

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

; ----------------------------------------------------------------------------------------------------------
; ----------------------------------------- BUTTONS PROGRAMMING --------------------------------------------
; ----------------------------------------------------------------------------------------------------------

; ------------- HATCH WEBSITES ---------------

PressTimeSheet:
	IniRead, CurrentCheckbox4Status, %LocalIniFilePath%, Checkboxes, CheckBox4
	
	if (CurrentCheckbox4Status)
		run, https://idcsapcl01n01.corp.hatchglobal.com/sap/bc/webdynpro/sap/hress_a_cats_1?WDCONFIGURATIONID=HRESS_AC_CATS_1&sap-client=010&sap-language=EN#,, UseErrorLevel
	else
		run, C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SAP Front End\SAP Logon.lnk
		
	If (ErrorLevel == "ERROR")
		{
			GoSub, GenericErrorLevel
			return
		}
		
	GoSub, CloseQAB
return

PressSharePoint:
	run, https://hatchengineering.sharepoint.com/,, UseErrorLevel
	If (ErrorLevel == "ERROR")
		{
			GoSub, GenericErrorLevel
			return
		}
		
	GoSub, CloseQAB
return

PressDashboard:
	run, https://app.powerbi.com/groups/me/reports/1e3a5f1d-a1d0-4844-a959-d4f9ca39ad05/ReportSection863ec14e12657866880b?ctid=82a1cdba-ae8d-4b25-adb6-9a4173a8be58,, UseErrorLevel
	If (ErrorLevel == "ERROR")
		{
			GoSub, GenericErrorLevel
			return
		}
		
	GoSub, CloseQAB
return

PressSafetyObservation:
	run, https://ipassm/NetForms/new/ROAM-Online,, UseErrorLevel
	If (ErrorLevel == "ERROR")
		{
			GoSub, GenericErrorLevel
			return
		}
		
	GoSub, CloseQAB
return

GenericErrorLevel:
	MsgBox, 4116, ERROR, Uh oh, there was an unusual error.`n`nWould you like to send me an email? I'll sort this out for ya!
	ContactLink := "mailto:mqmotiwala@gmail.com?subject=Quick%20Access%20Bar - General Error"
	IfMsgBox, Yes
		run, %ContactLink%
return

; ------------- SPECIAL BUTTONS ---------------

PressCopiedLink:
	run, % Clipboard,, UseErrorLevel
	If (ErrorLevel == "ERROR")
	{
		; remove trailing and leading blank space
		StrippedClipboard := RegExReplace(clipboard, "[`r`n`t\s]+$")
		
		; if clipboard exists and is not a direct link, Google search clipboard contents
		if (StrippedClipboard)
			run, https://www.google.com/search?q=%clipboard%,, UseErrorLevel
		else
			{
				MsgBox, 4112, ERROR, You don't have anything copied!
				return
			}
	}
	GoSub, CloseQAB
return

PressSnippingTool:
{
	; hide QAB and activate new snip
	Gui, QAB:Hide
	run, SnippingTool.exe
	WinWaitActive, Snipping Tool
	Send, ^n
	
	; restore QAB once snip is completed
	WinWaitActive, ahk_class Microsoft-Windows-SnipperEditor,,5
	IniRead, CurrentCheckbox1Status, %LocalIniFilePath%, Checkboxes, CheckBox1
	if (!CurrentCheckbox1Status)
		Gui, QAB:Show
	else
		Gui, QAB:Destroy
	
	UpdateLocalUsageStats()
}
Return

; ------------- DYNAMIC BUTTONS ---------------

PressDynamicButton:
	run, % LinksArray[A_GuiControl],, UseErrorLevel
	If (ErrorLevel == "ERROR")
	{
		MsgBox, 4112, ERROR, % "There was an error. Ensure linked address is functional.`nLinked Address:`n`n" LinksArray[A_GuiControl]
		return
	}
	GoSub, CloseQAB
return

; ----------------------------------------------------------------------------------------------------------
; ------------------------------------------ ABOUT BUTTON GUI  ---------------------------------------------
; ----------------------------------------------------------------------------------------------------------

PressAbout:
	; destroy the QAB tutorial window at install if currently active
	Gui, Welcome:Destroy
	
	; create About GUI
	Gui, About:New
	
	; declare About GUI preferences
	FontName     	:= "Segoe UI"   
	FontSize     	:= 10       
	VerticalMargin 	:= 10            
	Margin   		:= 20       
	TextHeight		:= 17
	ButtonWidth  	:= 88      
	ButtonHeight 	:= 26 
	SS_WHITERECT 	:= 0x0006      

	ExtraRowSpacing	:= 2*TextHeight
	BottomHeight	:= ButtonHeight + 2*VerticalMargin
		
	; create a white background
	Gui, Add, Text, x0 y%VerticalMargin% %SS_WHITERECT% vWhiteBox
	
; --- ABOUT QUICK ACCESS BAR [AREA 1]
	; section heading
	Area1Y := 2*VerticalMargin
	Gui, Font, s%FontSize% w600 Underline, %FontName%
	Gui, Add, Text, x%Margin% y%Area1Y% BackgroundTrans, How To Use Quick Access Bar
	Gui, Font, Norm

	; add About information
	A1Sentence0Y := Area1Y + TextHeight
	A1Sentence1Y := A1Sentence0Y + ExtraRowSpacing
	A1Sentence2Y := A1Sentence1Y + TextHeight
	A1Sentence3Y := A1Sentence2Y + ExtraRowSpacing
	
	A1Sentence0	:= "Press any button to go to its mapped address."
	A1Sentence1 := "You can also create your own buttons! To do that, press '" . EditLinksButtonName . "' on the Quick Access Bar."
	A1Sentence2 := "Any file, folder, software or website can be linked for quick access."
	A1Sentence3	:= "For general feedback, bug reports or new feature ideas - "
    
	Gui, Add, Text, x%Margin% y%A1Sentence0Y% BackgroundTrans vA1Sentence0, %A1Sentence0%
	Gui, Add, Text, x%Margin% y%A1Sentence1Y% BackgroundTrans vA1Sentence1, %A1Sentence1%
	Gui, Add, Text, x%Margin% y%A1Sentence2Y% BackgroundTrans vA1Sentence2, %A1Sentence2%
	Gui, Add, Text, x%Margin% y%A1Sentence3Y% BackgroundTrans vA1Sentence3, %A1Sentence3%

	; add developer contact information
	GuiControlGet, A1Sentence3Dimensions, Pos, A1Sentence3
	
	LinkBoxX := Margin + A1Sentence3DimensionsW
	LinkBoxY := A1Sentence3DimensionsY

	Gui, Font, Underline
	Gui, Add, Link, x%LinkBoxX% y%LinkBoxY% BackgroundTrans vLinkBox, <a href="mailto:mqmotiwala@gmail.com?subject=Quick Access Bar - Feedback">contact Mufaddal Motiwala</a>
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
	Gui, Font, s%FontSize% w600 Underline, %FontName%
	Gui, Add, Text, x%Margin% y%Area2Y% BackgroundTrans, Special Buttons Explanation
	Gui, Font, Norm

	; add Explanations
	A2Sentence1Y 	:= Area2Y + TextHeight
	A2Sentence2Y	:= A2Sentence1Y + TextHeight + ExtraRowSpacing
	A2Sentence3Y	:= A2Sentence2Y + ExtraRowSpacing
	
	Gui, Font, italic underline
	Gui, Add, Text, x%Margin% y%A2Sentence1Y% BackgroundTrans vA2Sentence1A, %CopiedAddressSpecialButtonName%
	Gui, Add, Text, x%Margin% y%A2Sentence2Y% BackgroundTrans vA2Sentence2A, Snipping Tool
	Gui, Add, Text, x%Margin% y%A2Sentence3Y% BackgroundTrans vA2Sentence3A, Unit Converter
	Gui, Font, Norm
	
	GuiControlGet, A2Sentence1ADimensions, Pos, A2Sentence1A
	GuiControlGet, A2Sentence2ADimensions, Pos, A2Sentence2A
	GuiControlGet, A2Sentence3ADimensions, Pos, A2Sentence3A
	
	A2Sentence1B 	:= " will navigate to the currently copied content if it is a working address to a destination."
	A2Sentence1C 	:= "Otherwise, it will perform a Google search of the copied content."
	A2Sentence2B	:= " will begin a new snip with your default snip settings." 
	A2Sentence3B	:= " is a simple conversion tool for commonly used engineering units."
	A2Sentence3C	:= "It can also retrieve real-time exchange rates for popular currencies."

	A2Sentence1CY := A2Sentence1ADimensionsY + TextHeight
	A2Sentence1BX := A2Sentence1ADimensionsW + Margin
	A2Sentence2BX := A2Sentence2ADimensionsW + Margin
	A2Sentence3BX := A2Sentence3ADimensionsW + Margin
	A2Sentence3CY := A2Sentence3ADimensionsY + TextHeight
	
	Gui, Add, Text, x%A2Sentence1BX% y%A2Sentence1Y% BackgroundTrans vA2Sentence1B, %A2Sentence1B%
	Gui, Add, Text, x%Margin% y%A2Sentence1CY% BackgroundTrans vA2Sentence1C, %A2Sentence1C%
	Gui, Add, Text, x%A2Sentence2BX% y%A2Sentence2Y% BackgroundTrans vA2Sentence2B, %A2Sentence2B%
	Gui, Add, Text, x%A2Sentence3BX% y%A2Sentence3Y% BackgroundTrans vA2Sentence3B, %A2Sentence3B%
	Gui, Add, Text, x%Margin% y%A2Sentence3CY% BackgroundTrans vA2Sentence3C, %A2Sentence3C%

; --- STATISTICS [AREA 3]
	; section heading
	Area3Y := A2Sentence3CY + ExtraRowSpacing
	Gui, Font, s%FontSize% w600 Underline, %FontName%
	Gui, Add, Text, x%Margin% y%Area3Y% BackgroundTrans, Usage Statistics
	Gui, Font, Norm
	
	; add usage statistics
	AcquireLocalUsageStats()
	MinimumActivations := 60/SecondsSavedPerUse
	
	if (Activations == 0)
		Sentence6	:= "Usage statistics will be available once you start using Quick Access Bar."
	else if (DailyAverage >= 1 and DaysUsed >= 3 and Activations >= MinimumActivations)
		{
			Sentence6	:= "Total Usage:`nAverage:`nTime Saved:"
			Sentence7	:= Activations "`n" DailyAverage "x daily`n" TimeSaved
		}
	else
		{	
			if (Activations == 1)
				Sentence6	:= "You've used Quick Access Bar " . Activations . " time. Check in later for more!"
			else
				Sentence6	:= "You've used Quick Access Bar " . Activations . " times. Check in later for more!"
		}
	
	Sentence6Y	:= Area3Y + TextHeight
	Gui, Add, Text, x%Margin% y%Sentence6Y% BackgroundTrans vSentence6, %Sentence6%
	GuiControlGet, Sentence6Dim, Pos, Sentence6
	Sentence7X := Sentence6DimW + Margin + 10
	Gui, Add, Text, x%Sentence7X% y%Sentence6Y% BackgroundTrans vSentence7, %Sentence7%
	
; --- USER SETTINGS [AREA 4]
	; section heading
	Area4Y := Sentence6Y + Sentence6DimH + TextHeight
	Gui, Font, s%FontSize% w600 Underline, %FontName%
	Gui, Add, Text, x%Margin% y%Area4Y% BackgroundTrans, User Settings
	Gui, Font, Norm

	; add checkboxes
	CheckBox1Y := Area4Y + TextHeight
	CheckBox2Y := CheckBox1Y + TextHeight
	CheckBox3Y := CheckBox2Y + TextHeight
	CheckBox4Y := CheckBox3Y + TextHeight
	
	; retrieve current user settings
	GoSub, ReadCheckBoxValues	
		
	Gui, Add, CheckBox, x%Margin% y%CheckBox1Y% Checked%CurrentCheckBox1Status% vCheckBox1Button gToggleCheckBox
	Gui, Add, CheckBox, x%Margin% y%CheckBox2Y% Checked%CurrentCheckBox2Status% vCheckBox2Button gToggleCheckBox
	Gui, Add, CheckBox, x%Margin% y%CheckBox3Y% Checked%CurrentCheckBox3Status% vCheckBox3Button gToggleCheckBox
	Gui, Add, CheckBox, x%Margin% y%CheckBox4Y% Checked%CurrentCheckBox4Status% vCheckBox4Button gToggleCheckBox
	
	; set checkboxes width to match their height
	GuiControlGet, CheckBoxDim, Pos, CheckBox1Button
	GuiControl, Move, CheckBox1Button, w%CheckBoxDimH%
	GuiControl, Move, CheckBox2Button, w%CheckBoxDimH%
	GuiControl, Move, CheckBox3Button, w%CheckBoxDimH%
	GuiControl, Move, CheckBox4Button, w%CheckBoxDimH%
	
	; add checkbox description texts
	; set initial text widths sufficiently large
	CheckBoxTextX 	:= Margin + CheckBoxDimH + 10
	GuiControlGet, WidestControlDim, Pos, A1Sentence1
	Gui, Add, Text, x%CheckboxTextX% y%CheckBox1Y% w%WidestControlDimW% BackgroundTrans vCheckbox1Text
	Gui, Add, Text, x%CheckboxTextX% y%CheckBox2Y% w%WidestControlDimW% BackgroundTrans vCheckbox2Text
	Gui, Add, Text, x%CheckboxTextX% y%CheckBox3Y% w%WidestControlDimW% BackgroundTrans vCheckbox3Text
	Gui, Add, Text, x%CheckboxTextX% y%CheckBox4Y% w%WidestControlDimW% BackgroundTrans vCheckbox4Text

	; alter checkbox text values as per their current status
	CheckBox1Status1Text 	:= "Quick Access Bar will close automatically after each use."
	CheckBox1Status0Text 	:= "Quick Access Bar will remain visible after use."
	
	CheckBox2Status1Text 	:= "Ctrl + Shift + C will trigger the " . CopiedAddressSpecialButtonName " special button."
	CheckBox2Status0Text 	:= "No shortcut set for Copied Address special button."	 

	CheckBox3Status1Text 	:= "Ctrl + Shift + S will trigger Snipping Tool special button."
	CheckBox3Status0Text 	:= "No shortcut set for Snipping Tool special button."
	
	CheckBox4Status1Text 	:= "My preferred Timesheet tool is via my internet browser."
	CheckBox4Status0Text 	:= "My preferred Timesheet tool is SAP Logon."
	
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
	
	if (CurrentCheckbox4Status)
		GuiControl, Text, Checkbox4Text, %CheckBox4Status1Text%
	else
		GuiControl, Text, Checkbox4Text, %CheckBox4Status0Text%	
		
; --- OVERALL ABOUT GUI CONFIGURATION & OK BUTTON [AREA 4]
	; acquire relevant control dimensions for GUI width and height calculations
	GuiControlGet, WidestControlDim, Pos, Sentence1
	GuiControlGet, LowestControlDim, Pos, CheckBox4Button
	
	; resize white background to fit all of GUI
	WhiteBoxHeight 	:= LowestControlDimY + LowestControlDimH + VerticalMargin                                       
	GuiWidth 		:= Margin + WidestControlDimW + Margin
	
	GuiControl, Move, WhiteBox, w%GuiWidth% h%WhiteBoxHeight%
 
	; place OK button
	ButtonX := GuiWidth - Margin - ButtonWidth
	ButtonY := WhiteBoxHeight + (BottomHeight - ButtonHeight)*0.75
	
	Gui, Add, Button, x%ButtonX% y%ButtonY% w%ButtonWidth% h%ButtonHeight% gButtonOK Default, OK
	GuiControl, Focus, OK
	
	; show About GUI
	GuiHeight := WhiteBoxHeight + BottomHeight
	AboutGuiX := QABX + 2*Margin + GeneralButtonWidth
		
	Gui, +AlwaysOnTop +ToolWindow -SysMenu 
	Gui, Show, x%AboutGuiX% yCenter w%GuiWidth% h%GuiHeight%, %AboutTitle%
Return

; Pressing OK on About GUI will return to Quick Access Bar
ButtonOK:
	Gui, About:Destroy
return

ReadCheckBoxValues:
	IniRead, CurrentCheckbox1Status, %LocalIniFilePath%, Checkboxes, CheckBox1, 1
	IniRead, CurrentCheckbox2Status, %LocalIniFilePath%, Checkboxes, CheckBox2, 1
	IniRead, CurrentCheckbox3Status, %LocalIniFilePath%, Checkboxes, CheckBox3, 1
	IniRead, CurrentCheckbox4Status, %LocalIniFilePath%, Checkboxes, CheckBox4, 1
	
	; turn off special button shortcut hotkeys if user desires
	if not (CurrentCheckBox2Status)
		Hotkey, ^+c, Off
	if not (CurrentCheckBox3Status)
		Hotkey, ^+s, Off
return

ToggleCheckbox:
	; update all checkbox settings
	Gui, Submit, NoHide
			
	; dynamically update checkbox text
	if (CheckBox1Button)
		GuiControl, Text, CheckBox1Text, %CheckBox1Status1Text%
	else
		GuiControl, Text, CheckBox1Text, %CheckBox1Status0Text%
	
	if (CheckBox2Button)
		{
			GuiControl, Text, CheckBox2Text, %CheckBox2Status1Text%
			Hotkey, ^+c, On
		}
	else
		{
			GuiControl, Text, CheckBox2Text, %CheckBox2Status0Text%
			Hotkey, ^+c, Off
		}
	
	if (CheckBox3Button)
		{
			GuiControl, Text, CheckBox3Text, %CheckBox3Status1Text%
			Hotkey, ^+s, On
		}
	else
		{
			GuiControl, Text, CheckBox3Text, %CheckBox3Status0Text%
			Hotkey, ^+s, Off
		}
	
	if (CheckBox4Button)
		GuiControl, Text, CheckBox4Text, %CheckBox4Status1Text%
	else
		GuiControl, Text, CheckBox4Text, %CheckBox4Status0Text%
	
	; write Checkbox status to .ini file
	IniWrite, %CheckBox1Button%, %LocalIniFilePath%, Checkboxes, CheckBox1
	IniWrite, %CheckBox2Button%, %LocalIniFilePath%, Checkboxes, CheckBox2
	IniWrite, %CheckBox3Button%, %LocalIniFilePath%, Checkboxes, CheckBox3
	IniWrite, %CheckBox4Button%, %LocalIniFilePath%, Checkboxes, CheckBox4
return
	
; ----------------------------------------------------------------------------------------------------------
; -------------------------------------------- LINKS EDITOR  -----------------------------------------------
; ----------------------------------------------------------------------------------------------------------

EditLinks:
	; declare Edit Links GUI preferences
	Offsets				= 5
	TextsGap			= 8
	TextsHeight			= 22
	MaxButtons 			= 25
	NamesEditWidth 		= 150
	LinksEditWidth 		= 600
	SaveButtonWidth 	= 50
	SaveButtonHeight 	= 25
	EditorFontSize		= 9
	
	EditerGuiWidth := NamesEditWidth + LinksEditWidth + 2*Offsets + 4
	
	NamesEditX	:= Offsets + 1
	LinksEditX	:= Offsets + NamesEditWidth + 3
	TextsY		:= Offsets + SaveButtonHeight + TextsGap
	EditsY 		:= Offsets + SaveButtonHeight + TextsHeight
	
	; create Edit Links GUI
	Gui, Editer:New,, Add Dynamic Buttons
	Gui +AlwaysOnTop +ToolWindow -SysMenu
	Gui, Font, s%EditorFontSize%
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
	EditerGuiX := QABX - (EditerGuiWidth + 2*HorizontalMargin)	
	Gui, Show, x%EditerGuiX% yCenter AutoSize, %EditLinksButtonName%
return

PressSave:
	; save edit fields contents in .ini file
	GuiControlGet, NamesEdit
	GuiControlGet, LinksEdit
	IniDelete, %LocalIniFilePath%, Buttons
	IniDelete, %LocalIniFilePath%, Links
	IniWrite, %NamesEdit%, %LocalIniFilePath%, Buttons
	IniWrite, %LinksEdit%, %LocalIniFilePath%, Links
	
	; return to Quick Access Bar
	Gui, Editer:Destroy
	GoSub, ShowQuickAccessGUI
return

; ----------------------------------------------------------------------------------------------------------
; ---------------------------------------- USAGE STATS AND UPDATES -----------------------------------------
; ----------------------------------------------------------------------------------------------------------

AcquireLocalUsageStats()
	{
		global
		
		; read parameters
		IniRead, Activations, %LocalIniFilePath%, Usage, Activations, 0
		IniRead, InstallationDate, %LocalIniFilePath%, Usage, InstallationDate
				
		; calculate days since installation date
		DaysUsed := A_Now
		EnvSub, DaysUsed, %InstallationDate%, seconds
		DaysUsed := DaysUsed/86400

		; calculate average daily use
		DailyAverage := round(Activations/DaysUsed)
		
		; calculate time saved
		TimeSaved := round(SecondsSavedPerUse*Activations/60)
		if (TimeSaved >= 60)
			{
				; format as h:mm
				MinsSaved := round((TimeSaved/60 - floor(TimeSaved/60))*60)
				MinsSaved := Format("{1:02}", MinsSaved)
				TimeSaved := round(TimeSaved/60) . ":" . MinsSaved . " hours"
			}
		else 
			{
				if (TimeSaved == 1)
					TimeSaved := TimeSaved . " minute"
				else
					TimeSaved := TimeSaved . " minutes"
			}
	}
return

UpdateLocalUsageStats()
	{
		global
		AcquireLocalUsageStats()
		
		; update local ini file with statistics
		Activations := Activations + 1
		IniWrite, %InstallationDate%, %LocalIniFilePath%, Usage, InstallationDate
		IniWrite, %Activations%, %LocalIniFilePath%, Usage, Activations
		IniWrite, %DailyAverage%, %LocalIniFilePath%, Usage, DailyAverage
		IniWrite, %TimeSaved%, %LocalIniFilePath%, Usage, TimeSaved
		
		; update network file path with statistics at install
		if (Activations == 1 or mod(Activations, 5) == 0)
			GoSub, UpdateUsageStatsOnNetwork
	}
return

UpdateUsageStatsOnNetwork:
	; version
	IniRead, CurrentVersion, %LocalIniFilePath%, Version, VersionKey, -1
	IniWrite, %CurrentVersion%, %UserStatsNetworkFilePath%, %A_Username%, Version
	
	; installation date
	IniWrite, %InstallationDate%, %UserStatsNetworkFilePath%, %A_Username%, InstallationDate
	
	; last updated
	FormatTime, LastUpdated,,ddMMMyyyy hh:mmtt
	IniWrite, %LastUpdated%, %UserStatsNetworkFilePath%, %A_Username%, LastUpdated
	
	; stats
	IniWrite, %TimeSaved%, %UserStatsNetworkFilePath%, %A_Username%, TimeSaved
	IniWrite, %Activations%, %UserStatsNetworkFilePath%, %A_Username%, Activations
	IniWrite, %DailyAverage%, %UserStatsNetworkFilePath%, %A_Username%, DailyAverage
return

CheckForUpdates:
	; check for updates on weekdays between 10AM to 6PM only
	if (A_Hour >= 10 && A_Hour < 18 && A_WDay >= 2 && A_WDay <= 6)
		{
			; acquire latest version number and patch notes
			IniRead, LatestVersion, %UpdateInfoNetworkFilePath%, Update, LatestVersion, 0
			
			; acquire current version
			IniRead, CurrentVersion, %LocalIniFilePath%, Version, VersionKey, 0
			
			; run Updater.exe if needed
			if (LatestVersion > CurrentVersion and FileExist(LatestVersionExeFilePath))
				{
					; ensures Updater.exe is installed and is the latest associated with current version
					; wait till file is deleted 
					FileDelete, Quick Access Bar - Updater.exe
					while FileExist("Quick Access Bar - Updater.exe")
						Sleep, 50
					
					; wait till file is installed
					FileInstall, Quick Access Bar - Updater.exe, Quick Access Bar - Updater.exe, True
					while !FileExist("Quick Access Bar - Updater.exe")
						Sleep, 50 
					
					run, "Quick Access Bar - Updater.exe",, UseErrorLevel
					If (ErrorLevel == "ERROR")
						{
							; if Updater.exe fails to run, ensure Updater.exe is installed by running PrerequisiteTasks() again
							PrerequisiteTasks()
						}

					; exit current version as its no longer needed	
					ExitApp
				}
		}
return

; ----------------------------------------------------------------------------------------------------------
; ------------------------------------------- TRAY MENU CODE -----------------------------------------------
; ----------------------------------------------------------------------------------------------------------

TrayMenuShowQAB:
GoSub, ShowQuickAccessGUI
return

TrayMenuDynamicButtons:
GoSub, ShowQuickAccessGUI
GoSub, EditLinks
return

TrayMenuAbout:
GoSub, ShowQuickAccessGUI
GoSub, PressAbout
return

TrayMenuUpdate:
; ensures Updater.exe is installed and is the latest associated with current version
; wait till file is deleted 
FileDelete, Quick Access Bar - Updater.exe
while FileExist("Quick Access Bar - Updater.exe")
	Sleep, 50

; wait till file is installed
FileInstall, Quick Access Bar - Updater.exe, Quick Access Bar - Updater.exe, True
while !FileExist("Quick Access Bar - Updater.exe")
	Sleep, 50 

run, "Quick Access Bar - Updater.exe",, UseErrorLevel
	If (ErrorLevel == "ERROR")
		{
			; if Updater.exe fails to run, ensure Updater.exe is installed by running PrerequisiteTasks() again
			PrerequisiteTasks()
		}
return

TrayMenuTutorial:
Gui, QAB:Destroy
Gui, About:Destroy
Gui, Editer:Destroy
GoSub, WelcometoQAB
return

TrayMenuExit:
ExitApp
return

; ----------------------------------------------------------------------------------------------------------
; ------------------------------------------- UNIT CONVERTER -----------------------------------------------
; ----------------------------------------------------------------------------------------------------------

PressUnitConverter:
	; set GUI parameters and physical constants
	GuiW := 276
	Margin := 2
	RangeMin := -2147483648
	RangeMax := 2147483647
	UnitConverterTitle := "Quick Access Bar - Unit Converter"
	
	TopY := Margin - 0
	TopH := 150
	BottomY := TopY + TopH
	BottomH := 70
	GroupBoxX := Margin + 1
	
	DisclaimerY := BottomY + BottomH
	DisclaimerH := 55
	DisclaimerTextX := GroupBoxX + 10
	DisclaimerTextY := DisclaimerY + 20
	DisclaimerText := "Accuracy in conversions is not guaranteed.`nRemember to always double check critical values."
	
	BottomSpacing := 30
	RadioH := 20
	RadioX := (GuiW - 135)/2
	Radio1Y := BottomY + 20
	Radio2Y := Radio1Y + RadioH
	
	DecimalsTextY := Radio1Y + 2
	DecimalsEditW := 40
	DecimalsEditH := 20
	DecimalsEditX := RadioX + BottomSpacing + 55 ; 55 is estimated radio text width
	DecimalsEditY := DecimalsTextY + 15
	
	; drop down 
	DropDownVGap := 10
	DropDownHGap := 25
	DropDownH := 20
	
	DropDownW := 120
	UnitsDropDownW := 140
	UnitGroupDropDownW := 160

	UnitGroupDropDownX := (GuiW - UnitGroupDropDownW)/2
	UnitsDropDownX := (GuiW - UnitsDropDownW)/2
	FromX := GroupBoxX + 5
	ToX := FromX + DropDownW + DropDownHGap
	
	UnitGroupDropDownY := TopY + 15
	UnitsDropDownY := UnitGroupDropDownY + DropDownH + 5
	FromY := UnitsDropDownY + DropDownH + DropDownVGap
	ToY := FromY
	
	FromEditY := FromY + DropDownH + 2
	ToEditY := ToY + DropDownH + 2

	ButtonW := 70
	ButtonH := 25
	ButtonY := FromEditY + DropDownH + DropDownVGap - 5
	CopyButtonX := ToX
	SwapButtonX := FromX + DropDownW - ButtonW
	
	; currency 
	CurrencyButtonX := (GuiW - ButtonW)/2
	CurrencyButtonY := BottomY + 20
	CurrencyTextX := GroupBoxX + Margin
	CurrencyTextY := CurrencyButtonY + ButtonH + 10
	CurrencyTextW := GuiW - 4*Margin
	
	TextX := FromX + DropDownW
	ToTextY := FromY + DropDownH/4
	EqualsTextY := ToY + DropDownH + DropDownH/4
		
	; volume standard dropdown
	VolTextX := FromX
	VolTextY := Radio1Y
	
	VolDropDownX := VolTextX
	VolDropDownY := VolTextY + 15 
	VolDropDownW := DropDownW
	VolDropDownH := DropDownH
	
	; set GUI layout
	Gui, UnitConverter:New
	Gui, Margin, %Margin%, %Margin%
	Gui, Add, GroupBox, x%GroupBoxX% y%TopY% w%GuiW% h%TopH% cBlack vInputsGroupBox, Inputs
	Gui, Add, GroupBox, x%GroupBoxX% y%BottomY% w%GuiW% h%BottomH% cBlack vSettingsGroupBox, Settings
	Gui, Add, GroupBox, x%GroupBoxX% y%DisclaimerY% w%GuiW% h%DisclaimerH% cBlack vDisclaimerGroupBox, Disclaimer
	Gui, Add, Radio, x%RadioX% y%Radio1Y% h%RadioH% vRB1 gSetPrecisionToReg Checked, Regular
	Gui, Add, Radio, x%RadioX% y%Radio2Y% h%RadioH% vRB2 gSetPrecisionToSci, Scientific
	Gui, Add, Edit, x%DecimalsEditX% y%DecimalsEditY% w%DecimalsEditW% h%DecimalsEditH% gSetPrecision vSetPrecision,
	Gui, Add, UpDown, vSetPrecisionUpDown Range0-12 Wrap, 2
	Gui, Add, Button, x%CopyButtonX% y%ButtonY% w%ButtonW% h%ButtonH% vCopyResultButton gCopyResult, Copy Result
	Gui, Add, Button, x%SwapButtonX% y%ButtonY% w%ButtonW% h%ButtonH% vSwapButton gSwap, Swap Units
	Gui, Add, Text, x%DecimalsEditX% y%DecimalsTextY% vSetPrecisionText, Decimals
	Gui, Add, Text, x%DisclaimerTextX% y%DisclaimerTextY%, %DisclaimerText%
	
	; add dropdown lists
	Gui, Add, DropDownList, x%UnitGroupDropDownX% y%UnitGroupDropDownY% w%UnitGroupDropDownW% h%DropDownH% Sort r30 gSelectedUnitGroup vSelectedUnitGroup
	Gui, Add, DropDownList, x%UnitsDropDownX% y%UnitsDropDownY% w%UnitsDropDownW% h%DropDownH% Sort r30 gSelectedUnit vSelectedUnit
	Gui, Add, DropDownList, x%FromX% y%FromY% w%DropDownW% h%DropDownH% Sort r30 vFrom gCalc
	Gui, Add, DropDownList, x%ToX% y%ToY% w%DropDownW% h%DropDownH% Sort r30 vTo gCalc
	GuiControl,, SelectedUnitGroup, |All||General|Engineering|Heat Transfer|Fluids
	
	; add edit and results
	Gui, Add, Edit, x%FromX% y%FromEditY% w%DropDownW% h%DropDownH% gCalc vInputValue
	Gui, Add, UpDown, Range%RangeMin%-%RangeMax% Wrap 0x80, 1
	Gui, Add, Edit, x%ToX% y%ToEditY% w%DropDownW% h%DropDownH% vResult +ReadOnly
	
	; GUI explanation texts
	Gui, Add, Text, x%TextX% y%ToTextY% w%DropDownHGap% +Center gSwap vSwap, To:
	Gui, Add, Text, x%TextX% y%EqualsTextY% w%DropDownHGap% +Center vEqual, =
	
	; currency controls
	Gui, Add, Button, x%CurrencyButtonX% y%CurrencyButtonY% w%ButtonW% h%ButtonH% vCurrencyButton gCurrencyCalc Hidden, Convert
	Gui, Add, Text, x%CurrencyTextX% y%CurrencyTextY% w%CurrencyTextW% +Center vCurrencyText Hidden 
	
	; volume reference temperature controls
	Gui, Add, Text, x%VolTextX% y%VolTextY% vVolText Hidden, Reference Temperature
	Gui, Add, DropDownList, x%VolDropDownX% y%VolDropDownY% w%VolDropDownW% h%VolDropDownH% r30 vVolDropDown gCalc Hidden
	
	; format From and To dropdowns based on Selected Unit
	GoSub, SelectedUnitGroup
	GuiControl, Focus, InputsGroupBox
		
	UCGUIX := QABX - (GuiW + 2*HorizontalMargin)
	Gui +AlwaysOnTop +ToolWindow +E0x840000
	Gui, Color, f7f7f7
	Gui, Show, x%UCGUIX% yCenter AutoSize, %UnitConverterTitle%
	
	GoSub, CloseQAB
Return

GuiClose:
	try	pwb.Quit ; in case IE was not quit
	Gui, UnitConverter:Destroy
return

CopyResult:
	Gui, Submit, NoHide
	Clipboard := Result

	AnimationSpeed := 10
	ParseString := "Copied"
	Progress := " "
	Loop, Parse, ParseString
		{
			Progress := Progress A_LoopField
			GuiControl,, CopyResultButton, %Progress%
			Sleep, %AnimationSpeed%
		}
	Sleep, % 50*AnimationSpeed
	GuiControl,, CopyResultButton, Copy Result
	
	GuiControl, Focus, InputValue
return

Swap:
	GuiControlGet, SelectedFrom,,From
	GuiControlGet, SelectedTo,,To
		
	GuiControl, ChooseString, From, %SelectedTo%
	GuiControl, ChooseString, To, %SelectedFrom%
	
	if (SelectedUnit = "Currency")
		GuiControl,, Result, Press Convert
	else
		GoSub, Calc
return

SetPrecisionToReg:
	ShowAsScientific = 0
	GoSub, Calc
	GuiControl, Focus, InputValue
return

SetPrecisionToSci:
	ShowAsScientific = 1
	GoSub, Calc
	GuiControl, Focus, InputValue
return

SetPrecision:
	Gui, Submit, NoHide
	RegP = 0.%SetPrecision%
	SetFormat, Float, %RegP%
	GoSub, Calc
return

SelectedUnitGroup:
	Gui, Submit, NoHide
	
	if (SelectedUnitGroup == "All")
		GuiControl,, SelectedUnit, |Angle|Area|Distance||Mass|Speed|Temperature|Time|Fuel Consumption|Acceleration|Density|Energy|Force|Power|Pressure|Volume|Moment of Inertia|Coeff. of Thermal Expansion|Specific Heat Capacity|Heat Transfer Coefficient|Thermal Conductivity|Volumetric Flow Rate|Mass Flow Rate|Concentration
	
	if (SelectedUnitGroup == "General")
		GuiControl,, SelectedUnit, |Angle|Area|Distance||Mass|Speed|Temperature|Time|Fuel Consumption
	
	if (SelectedUnitGroup == "Engineering")
		GuiControl,, SelectedUnit, |Acceleration|Density|Energy|Force|Power||Pressure|Volume|Moment of Inertia
	
	if (SelectedUnitGroup == "Heat Transfer")
		GuiControl,, SelectedUnit, |Coeff. of Thermal Expansion|Energy|Specific Heat Capacity|Heat Transfer Coefficient|Temperature|Thermal Conductivity||
	
	if (SelectedUnitGroup == "Fluids")
		GuiControl,, SelectedUnit, |Volumetric Flow Rate||Mass Flow Rate|Concentration
	
	GuiControl, Focus, InputValue
	GoSub, SelectedUnit
return

SelectedUnit:
	Gui, Submit, NoHide

	; currency is currently removed from the "All" and "General" dropdown lists so this section should never run
	if (SelectedUnit == "Currency")
		{
			Guicontrol,, from, 	|USD|EUR|JPY|GBP|AUD|CAD||CHF|HKD|NZD|KRW|SGD|MXN|INR|TRY|AED|COP|SAR
			Guicontrol,, to, 	|USD||EUR|JPY|GBP|AUD|CAD|CHF|HKD|NZD|KRW|SGD|MXN|INR|TRY|AED|COP|SAR
			
			; hide volume reference controls
			GuiControl, Hide, VolDropDown
			GuiControl, Hide, VolText
			
			; hide radio buttons and precision controls
			GuiControl, Hide, RB1
			GuiControl, Hide, RB2
			GuiControl, Hide, SetPrecision
			GuiControl, Hide, SetPrecisionUpDown
			GuiControl, Hide, SetPrecisionText
			
			; show currency controls 
			GuiControl, Show, CurrencyButton
			GuiControl, Show, CurrencyText
			
			GuiControl, Text, Result, Press Convert
			GuiControl, Text, SettingsGroupBox, Currency Controls
			return
		}
	else if (SelectedUnit == "Volume")
		{
			Guicontrol,, volDropDown, |15 °C (Standard)||0 °C (Normal)|60 °F|67 °F|
			
			; show volume reference controls
			GuiControl, Show, VolDropDown
			GuiControl, Show, VolText
									
			; hide currency controls 
			GuiControl, Hide, CurrencyButton
			GuiControl, Hide, CurrencyText		
			
			; hiding before move avoids strange visuals on GUI controls
			; hide radio buttons and precision controls
			GuiControl, Hide, RB1
			GuiControl, Hide, RB2
			GuiControl, Hide, SetPrecision
			GuiControl, Hide, SetPrecisionUpDown
			GuiControl, Hide, SetPrecisionText
			
			; move radio buttons and precision controls
			Move := 75
			GuiControlGet, PrecisionUpDown, Pos, SetPrecisionUpDown
			GuiControl, Move, RB1, % "x" RadioX + Move
			GuiControl, Move, RB2, % "x" RadioX + Move
			GuiControl, Move, SetPrecision, % "x" DecimalsEditX + Move
			GuiControl, Move, SetPrecisionUpDown, % "x" PrecisionUpDownX + Move
			GuiControl, Move, SetPrecisionText, % "x" DecimalsEditX + Move
			
			; show radio buttons and precision controls
			GuiControl, Show, RB1
			GuiControl, Show, RB2
			GuiControl, Show, SetPrecision
			GuiControl, Show, SetPrecisionUpDown
			GuiControl, Show, SetPrecisionText
			
			GuiControl, Text, SettingsGroupBox, Settings
		}
	else
		{
			; hide radio buttons and precision controls
			GuiControl, Hide, RB1
			GuiControl, Hide, RB2
			GuiControl, Hide, SetPrecision
			GuiControl, Hide, SetPrecisionUpDown
			GuiControl, Hide, SetPrecisionText
			
			; move radio buttons and precision controls
			Move := 0
			GuiControl, Move, RB1, % "x" RadioX + Move
			GuiControl, Move, RB2, % "x" RadioX + Move
			GuiControl, Move, SetPrecision, % "x" DecimalsEditX + Move
			GuiControl, Move, SetPrecisionUpDown, % "x" DecimalsEditX + Move + 20
			GuiControl, Move, SetPrecisionText, % "x" DecimalsEditX + Move
			
			; show radio buttons and precision controls
			GuiControl, Show, RB1
			GuiControl, Show, RB2
			GuiControl, Show, SetPrecision
			GuiControl, Show, SetPrecisionUpDown
			GuiControl, Show, SetPrecisionText
			
			GuiControl, Text, SettingsGroupBox, Settings
			
			; hide currency controls 
			GuiControl, Hide, CurrencyButton
			GuiControl, Hide, CurrencyText
			
			; hide volume reference controls
			GuiControl, Hide, VolDropDown
			GuiControl, Hide, VolText
		}

	if (SelectedUnit == "Mass")
		{
			Guicontrol,, from, 	|kilograms||grams|ounces|pounds|stone|ton(US)|ton(UK)|slugs
			Guicontrol,, to, 	|kilograms|grams|ounces|pounds||stone|ton(US)|ton(UK)|slugs
		}

	if (SelectedUnit == "Distance")
		{
			Guicontrol,, from, 	|feet|inches|mil|meters||centimeter|kilometer|millimeter|mile|yard
			Guicontrol,, to, 	|feet|inches||mil|meters|centimeter|kilometer|millimeter|mile|yard
		}

	if (SelectedUnit == "Density")
		{
			Guicontrol,, from, 	|lb/in³|lb/ft³|g/cm³|kg/m³||
			Guicontrol,, to,	|lb/in³|lb/ft³||g/cm³|kg/m³
		}

	if (SelectedUnit == "Acceleration")
		{
			Guicontrol,, from,	|m/s²||in/s²|ft/s²|g's
			Guicontrol,, to,	|m/s²|in/s²|ft/s²||g's
		}

	if (SelectedUnit == "Force")
		{
			Guicontrol,, from,	|Newton||lbf|dyne
			Guicontrol,, to,	|Newton|lbf||dyne
		}

	if (SelectedUnit == "Pressure")
		{
			Guicontrol,, from,	|Pa|kPa||mPa|psi|psf|torr|bar|atm|mm mercury|cm water
			Guicontrol,, to,	|Pa|kPa|mPa|psi||psf|torr|bar|atm|mm mercury|cm water
		}

	if (SelectedUnit == "Energy")
		{
			Guicontrol,, from,	|J|kJ||BTU|in lbf|ft lbf|kcal|therm|eV
			Guicontrol,, to,	|J|kJ|BTU||in lbf|ft lbf|kcal|therm|eV
		}

	if (SelectedUnit == "Power")
		{
			Guicontrol,, from,	|Watt|BTU/sec|BTU/hour|HP||ft lbf/s|kW
			Guicontrol,, to,	|Watt|BTU/sec|BTU/hour|HP|ft lbf/s|kW||
		}

	if (SelectedUnit == "Thermal Conductivity")
		{
			Guicontrol,, from,	|W/m-K||kW/m-K|BTU/hr-ft-F|BTU/hr-in-F|BTU-in/hr-ft²-F|cal/s-cm-C
			Guicontrol,, to,	|W/m-K|kW/m-K|BTU/hr-ft-F||BTU/hr-in-F|BTU-in/hr-ft²-F|cal/s-cm-C
		}

	if (SelectedUnit == "Specific Heat Capacity")
		{
			Guicontrol,, from,	|J/kg-K||BTU/lb-C|BTU/lb-F|cal/g-C|kJ/kg-K
			Guicontrol,, to,	|J/kg-K|BTU/lb-C|BTU/lb-F||cal/g-C|kJ/kg-K
		}

	if (SelectedUnit == "Heat Transfer Coefficient")
		{
			Guicontrol,, from,	|Watt/m²-K||BTU/hr-ft²-F|cal/s-cm²-C|kcal/hr-ft²-C
			Guicontrol,, to,	|Watt/m²-K|BTU/hr-ft²-F||cal/s-cm²-C|kcal/hr-ft²-C
		}

	if (SelectedUnit == "Area")
		{
			Guicontrol,, from,	|m²||cm²|mm²|in²|ft²|yd²|mil²|acre|km²
			Guicontrol,, to,	|m²|cm²|mm²|in²|ft²||yd²|mil²|acre|km²
		}

	if (SelectedUnit == "Volume")
		{
			Guicontrol,, from,	|m³||cm³|mm³|in³|ft³|yd³|liter
			Guicontrol,, to,	|m³|cm³|mm³|in³|ft³||yd³|liter
		}

	if (SelectedUnit == "Angle")
		{
			Guicontrol,, from,	|radians||degrees|angular mils|minutes|seconds|gradians
			Guicontrol,, to,	|radians|degrees||angular mils|minutes|seconds|gradians
		}

	if (SelectedUnit == "Temperature")
		{
			Guicontrol,, from,	|Kelvin|Celsius||Fahrenheit
			Guicontrol,, to,	|Kelvin|Celsius|Fahrenheit||
		}

	if (SelectedUnit == "Speed")
		{
			Guicontrol,, from,	|m/s|km/h||in/s|ft/s|mph
			Guicontrol,, to,	|m/s||km/h|in/s|ft/s|mph
		}

	if (SelectedUnit == "Coeff. of Thermal Expansion")
		{
			Guicontrol,, from,	|1/°K||1/°C|1/°F
			Guicontrol,, to,	|1/°K|1/°C||1/°F
		}
		
	if (SelectedUnit == "Volumetric Flow Rate")
		{
			Guicontrol,, from,	|m³/min|m³/hr||cm³/min|mm³/min|L/min|ft³/min|gallon(US)/min|gallon(UK)/min|m³/sec|cm³/sec|mm³/sec|L/sec|ft³/sec|gallon(US)/sec|gallon(UK)/sec
			Guicontrol,, to,	|m³/min|m³/hr|cm³/min|mm³/min|L/min|ft³/min|gallon(US)/min||gallon(UK)/min|m³/sec|cm³/sec|mm³/sec|L/sec|ft³/sec|gallon(US)/sec|gallon(UK)/sec
		}
	
	if (SelectedUnit == "Mass Flow Rate")
		{
			Guicontrol,, from, 	|kilograms/min||grams/min|ounces/min|pounds/min|kilograms/sec|grams/sec|ounces/sec|pounds/sec
			Guicontrol,, to, 	|kilograms/min|grams/min|ounces/min|pounds/min||kilograms/sec|grams/sec|ounces/sec|pounds/sec
		}

	if (SelectedUnit == "Concentration")
		{
			Guicontrol,, from, 	|kg/L||g/L|ppm|lb/gallon(US)|lb/gallon(UK)
			Guicontrol,, to,	|kg/L|g/L|ppm||lb/gallon(US)|lb/gallon(UK)
		}
	
	if (SelectedUnit == "Time")
		{
			Guicontrol,, from, 	|millisecond|second|minute|hour|day|week|month||year|decade|century
			Guicontrol,, to,	|millisecond|second|minute|hour||day|week|month|year|decade|century
		}
		
	if (SelectedUnit == "Fuel Consumption")
		{
			Guicontrol,, from, 	|L/100km||L/km|km/L|miles per US gallon|miles per UK gallon
			Guicontrol,, to,	|L/100km|L/km|km/L|miles per US gallon||miles per UK gallon
		}
	
	if (SelectedUnit == "Moment of Inertia")
		{
			Guicontrol,, from, 	|kg-m²||kg-cm²|lb-ft²|lb-in²
			Guicontrol,, to,	|kg-m²|kg-cm²|lb-ft²||lb-in²
		}
	
	GuiControl, Focus, InputValue
	GoSub, Calc
return

CurrencyCalc:
	Gui, Submit, NoHide
	
	; prevents multiple IE objects 
	if (Result == "Retrieving...")
		return
	
	GuiControl,, Result, Retrieving...
	WebLink := "https://www.xe.com/currencyconverter/convert/?Amount=" . InputValue . "&From=" . From . "&To=" . To
	
	try	
		{
			pwb := ComObjCreate("InternetExplorer.Application")
			pwb.visible := True
			pwb.Navigate(WebLink)
	
			LastUpdated := "Response time depends on your network speed."
			LastUpdated := "This feature is currently under development. Sorry!"
			GuiControl,, CurrencyText, %LastUpdated%
			while pwb.busy or pwb.ReadyState != 4
				Sleep, 500
		
			val := pwb.document.getElementByID("converterResult").GetElementsByTagName("Span")[4].InnerText
			pwb.Quit
		}
	
	catch
		{
			GuiControl,, Result, Error.
			GuiControl,, CurrencyText, Error. Check network connection.
			return
		}
		
	; remove all commas from parsed text and round to 2 decimals
	val := StrReplace(val, ",")
	val := round(val, 2)
	
	if (StrLen(val))
		{
			FormatTime, UpdateTime,, Time
			LastUpdated := "Based on exchange rates as of " . UpdateTime 
		}
	else
		LastUpdated := "Error. Check network connection."

	; only update if retrieving
	; retrieving text will change if user presses 'Swap Units' mid calculation
	Gui, Submit, NoHide
	if (Result == "Retrieving...")
		GuiControl,, Result, %val%
	
	GuiControl,, CurrencyText, %LastUpdated%
return

Calc:
	Gui, Submit, NoHide
	SetFormat, Float, 0.16E
	GuiControl, Focus, InputValue
	
	if (SelectedUnit == "Currency")
		{
			GuiControl, Text, Result, Press Convert
			return
		}
	else if (SelectedUnit == "Temperature")
	   {
		   if (From == "Kelvin")
				{
					 if To = Kelvin
						val := InputValue
					 if To = Fahrenheit
						val := 9*InputValue/5 - 459.67
					 if To = Celsius
						val := InputValue - 273.15
				}
				
		   else if (From == "Fahrenheit")
				 {
					 if To = Kelvin
						val := 5*(InputValue + 459.67)/9
					 if To = Fahrenheit
						val := InputValue
					 if To = Celsius
						val := 5*(InputValue - 32)/9
				 }
				 
		   else if (From == "Celsius")
				 {
					 if To = Kelvin
						val := InputValue + 273.15
					 if To = Fahrenheit
						val := 9*InputValue/5 + 32
					 if To = Celsius
						val := InputValue
				 }
	   }
	else
		{
			; engineering units
			ConversionFactorsSection1 := {
			(Join
				"Newton":1.0, "lbf":4.4482216152605, "dyne":1.0E-5,
				"m/s²":1.0, "in/s²":0.0254, "ft/s²":0.3048, "g's":9.80665,
				"kg-m²":1.0, "kg-cm²":0.0001, "lb-ft²":0.04214011, "lb-in²":0.00029264,
				"lb/in³":1.0, "lb/ft³":0.00057870368028786, "kg/m³":0.000036127298147753, "g/cm³":0.036127298147753,
				"Watt":1.0, "BTU/hour":0.293071, "BTU/sec":1055.055, "HP":735.4987485, "kW":1000, "ft lbf/s":1.355817,
				"m³":1.0, "cm³":0.000001, "mm³":1.0E-9, "in³":0.000016387064, "ft³":0.028316846592, "yd³":0.76455485798, "liter":0.001,
				"Pa":1.0, "mPa":1000000, "kPa":1000, "psi":6894.757293168, "psf":47.88025898, "torr":133.322, "mm mercury":133.3224, "bar":1.0E5, "atm":101325, "cm water":98.0665
			)}
						
			; fluids units
			ConversionFactorsSection2 := {
			(Join
				"kg/L":1.0, "g/L":1.0E-3, "ppm":9.988590004E-7, "lb/gallon(US)":0.119826428, "lb/gallon(UK)":0.099776374,
				"kilograms/min":1.0, "grams/min":1.0E-3, "ounces/min":0.0283495, "pounds/min":0.453592, "kilograms/sec":60.0, "grams/sec":0.06, "ounces/sec":1.70097, "pounds/sec":27.215519999,
				"m³/min":1.0, "m³/hr":0.016666666666666666, "cm³/min":1.0E-6, "mm³/min":1.0E-9, "L/min":1.0E-3, "ft³/min":0.028316846, "gallon(US)/min":3.785411784E-3, "gallon(UK)/min":4.54609E-3, "m³/sec":60, "cm³/sec":0.00006, "mm³/sec":6.0E-8, "L/sec":0.06, "ft³/sec":1.69901082 , "gallon(US)/sec":0.22712470704, "gallon(UK)/sec":0.27276539999
			)}
			
			; general units
			ConversionFactorsSection3 := {
			(Join
				"m/s":1.0, "km/h":0.277777777778, "in/s":0.02539999919, "ft/s":0.3047999902464, "mph":0.44704,
				"L/100km":1.0, "L/km":100, "km/L":100, "miles per US gallon":0.004251437075, "miles per UK gallon":0.0035400619,
				"km²":1.0E+6, "m²":1.0, "cm²":0.0001, "mm²":0.000001, "in²":0.00064516, "ft²":0.09290304, "yd²":0.83612736, "mil²":6.4516E-10, "acre":4046.8564224, 
				"millisecond":0.001, "second":1.0, "minute":60.0, "hour":3600, "day":86400, "week":604800, "month":2592000, "year":31536000, "decade":315360000, "century":3153600000,
				"radians":1.0, "degrees":0.017453292519943, "minutes":0.00029088820866572, "seconds":0.0000048481368110954, "gradians":0.015707963267949, "angular mils":0.00098174770424681, 
				"kilograms":2.2046223302272, "grams":0.0022046223302272, "ounces":0.062499991732666, "pounds":1.0, "stone":13.999998148117, "ton(US)":2000, "ton(UK)":2240, "slugs":32.174048695,
				"feet":1.0, "inches":0.08333333333333, "mil":0.000083333333333, "meters":3.2808399, "kilometer":3280.8399, "centimeter":0.032808399, "millimeter":0.0032808399, "mile":5280.0, "yard":3.0
			)}
				
			; heat transfer units
			ConversionFactorsSection4 := {
			(Join
				"1/°K":1.0, "1/°C":1.0, "1/°F":1.8, "1/°R":1.8,
				"kJ/kg-K":1000, "J/kg-K":1, "BTU/lb-C":2326, "BTU/lb-F":4186.8, "cal/g-C":4186.8,
				"Watt/m²-K":1.0, "BTU/hr-ft²-F":5.678263, "cal/s-cm²-C":41868, "kcal/hr-ft²-C":12.518428,
				"W/m-K":1.0, "kW/m-K":1000, "BTU/hr-ft-F":1.729577, "BTU/hr-in-F":20.754924, "BTU-in/hr-ft²-F":0.144227888, "cal/s-cm-C":418.4,
				"J":1.0, "kJ":1000, "BTU":1.0543503E3, "eV":1.602176634E-19, "in lbf":0.112984, "ft lbf":1.3558179483314, "kcal":4186.8, "therm":105587000
			)}
			
			ConversionFactorsSections := {1:ConversionFactorsSection1, 2:ConversionFactorsSection2, 3:ConversionFactorsSection3, 4:ConversionFactorsSection4}
			For every, Section in ConversionFactorsSections
				{
					For key, value in Section
						{
							if (From == key)
								From := value
								
							if (To == key)
								To := value
						}
				}

			val := (From/To)*InputValue
	   }
		
	if (SelectedUnit == "Volume")
		{
			; based on ideal gas law: v1/v2 = t1/t2
			; t1 = 15 °C = 288.15K 
			; t2 is the various other reference temperatures
		
			if (VolDropDown == "15 °C (Standard)")
				val := val*1
			if (VolDropDown == "0 °C (Normal)")
				val := val*1.054914881933
			if (VolDropDown == "60 °F")
				val := val*0.998075701887737
			if (VolDropDown == "67 °F")
				val := val*0.984810222720109
		}
	
	if ShowAsScientific
	   SetFormat, Float, %RegP%E
	else
	   SetFormat, Float, %RegP%
	   
	val := val + 0
	GuiControl,, Result, %val%
return


; ----------------------------------------------------------------------------------------------------------
; ------------------------------------------- PREREQUISITE TASKS -------------------------------------------
; ----------------------------------------------------------------------------------------------------------

PrerequisiteTasks()
	{
		global
		
		; check to ensure user is installing QAB locally
		IsInNetworkDirectory := GetUNCPath()
		if (IsInNetworkDirectory)
			{
				GoSub, ShowInstallError
				return
			}
		
		; check for updates every 30 minutes
		; after 1 hr, update Updater.exe - only runs once
		SetTimer, CheckForUpdates, 1800000
		SetTimer, UpdateUpdater, 3600000
		
		; extract Updater.exe to working directory
		FileInstall, Quick Access Bar - Updater.exe, Quick Access Bar - Updater.exe, True
		FileInstall, Quick Access Bar - ReadMe.txt, Quick Access Bar - ReadMe.txt, True
		
		; --- DECLARE GLOBAL VARIABLES ---
		; universal variables associated with multiple subroutine
		ScratchFolderName	:= "snedir"
		UserStatsNetworkFilePath	:= "\\idcdata01\Scratch\" . ScratchFolderName . "\Users\UsageStatistics.ini"
		UpdateInfoNetworkFilePath 	:= "\\idcdata01\Scratch\" . ScratchFolderName . "\Update\LatestVersion.ini"
		LatestVersionExeFilePath	:= "\\idcdata01\Scratch\" . ScratchFolderName . "\Update\Quick Access Bar.exe"
		LocalIniFilePath := "QABInitializationFile.ini"
		
		; variables associated with Tutorial and Updater
		WelcomeTitle := "Quick Access Bar Tutorial"
		UpdaterTitle := "Quick Access Bar - Update"
		
		; variables associated with Quick Access Bar
		QABTitle := "Quick Access Bar"
		EditLinksButtonName := "Add Dynamic Buttons"
		CopiedAddressSpecialButtonName := "Copied Text"
		
		; variables associated with About window
		; adding leading and following blanks to differentiate from main QAB GUI title
		AboutTitle := " Quick Access Bar "
		SecondsSavedPerUse := 5

		; create special button shortcut hotkeys
		Hotkey, ^+s, PressSnippingTool
		Hotkey, ^+c, PressCopiedLink
		GoSub, ReadCheckBoxValues

		; --- CHECK IF VERSION KEY IS MISSING ---
		IniRead, CurrentVersion, %LocalIniFilePath%, Version, VersionKey, -1
		if (CurrentVersion == -1)
			{
				IniRead, LatestVersion, %UpdateInfoNetworkFilePath%, Update, LatestVersion, 0
				IniWrite, %LatestVersion%, %LocalIniFilePath%, Version, VersionKey
				IniWrite, %LatestVersion%, %UserStatsNetworkFilePath%, %A_Username%, Version
			}		
			
		; --- CHECK IF JUST INSTALLED ---
		IniRead, Activations, %LocalIniFilePath%, Usage, Activations, 0
		if (Activations == 0) ; i.e. at time of install
			{
				GoSub, WelcomeToQAB ; run tutorial
				IniWrite, %A_Now%, %LocalIniFilePath%, Usage, InstallationDate
				IniWrite, %A_Now%, %UserStatsNetworkFilePath%, %A_Username%, InstallationDate
			}
	}
return

UpdateUpdater:
	; this ensures Updater.exe is the latest version
	FileInstall, Quick Access Bar - Updater.exe, Quick Access Bar - Updater.exe, True
	SetTimer, UpdateUpdater, Off
Return

GetUNCPath() 
	{
		; return UNC path from drive\folder path
		; returns nothing if drive is not a network drive or the drive does not exist

		oFSO := ComObjCreate("Scripting.FileSystemObject")
		oDrive := oFSO.GetDrive(oFSO.GetDriveName(oFSO.GetAbsolutePathName(A_WorkingDir)))
		sShareName := oDrive.ShareName

		return sShareName
	}

WelcomeToQAB:
	; create Welcome GUI
	Gui, Welcome:New

	; declare Welcome GUI preferences
	FontName     	:= "Segoe UI"   
	FontSize     	:= 10       
	VerticalMargin 	:= 10            
	Margin   		:= 15     
	TextHeight		:= 17 
	SS_WHITERECT 	:= 0x0006      

	ExtraRowSpacing	:= 2*TextHeight
	
	; create a white background
	Gui, Add, Text, x0 y%VerticalMargin% %SS_WHITERECT% vWhiteBox
	
;--- ABOUT QUICK ACCESS BAR [AREA 1]
	; section heading
	Area1Y := 2*VerticalMargin
	Gui, Font, s12 w600 Underline, %FontName%
	Gui, Add, Text, x%Margin% y%Area1Y% BackgroundTrans vWelcomeTitle cRed, Quick Access Bar
	Gui, Font, Norm s%FontSize%

	; add About information
	Sentence1		:= "Thank you for installing Quick Access Bar!"
	Sentence2 		:= "Quick Access Bar acts as a hub for your frequently used files, folders, software and websites."
	Sentence3A		:= "At any time, "
	Sentence3B		:= "launch Quick Access Bar by pressing the Windows Key and W"
	Sentence3C		:= " simultaneously."
	Sentence4		:= "The software logo is a handy way to remember this shortcut. You can see it on your taskbar right now."
    Sentence5		:= "To get started, first launch Quick Access Bar then press the '?' button located at its top right."
	
	Sentence1Y		:= Area1Y + ExtraRowSpacing
	Sentence2Y		:= Sentence1Y + ExtraRowSpacing
	Sentence3Y		:= Sentence2Y + ExtraRowSpacing
	Sentence4Y		:= Sentence3Y + TextHeight
	Sentence5Y		:= Sentence4Y + ExtraRowSpacing
	
	Gui, Add, Text, x%Margin% y%Sentence1Y% BackgroundTrans vWelcomeSentence1, %Sentence1%
	Gui, Add, Text, x%Margin% y%Sentence2Y% BackgroundTrans vWelcomeSentence2, %Sentence2%
	Gui, Add, Text, x%Margin% y%Sentence3Y% BackgroundTrans vWelcomeSentence3A, %Sentence3A%
	
	; add launching instructions as underlined italics
	GuiControlGet, Sentence3ADim, Pos, WelcomeSentence3A
	Sentence3BX := Sentence3ADimX + Sentence3ADimW
	Gui, Font, Italic Underline s%FontSize%
	Gui, Add, Text, x%Sentence3BX% y%Sentence3Y% BackgroundTrans vWelcomeSentence3B, %Sentence3B%
	
	GuiControlGet, Sentence3BDim, Pos, WelcomeSentence3B
	Sentence3CX := Sentence3BDimX + Sentence3BDimW
	Gui, Font, Norm s%FontSize%
	Gui, Add, Text, x%Sentence3CX% y%Sentence3Y% BackgroundTrans vWelcomeSentence3C, %Sentence3C%
	
	Gui, Add, Text, x%Margin% y%Sentence4Y% BackgroundTrans vWelcomeSentence4, %Sentence4%
	Gui, Add, Text, x%Margin% y%Sentence5Y% BackgroundTrans vWelcomeSentence5, %Sentence5%
		
	; acquire relevant control dimensions for GUI width and height calculations
	GuiControlGet, WidestControlDim, Pos, WelcomeSentence4
	GuiControlGet, LowestControlDim, Pos, WelcomeSentence5
	GuiControlGet, WelcomeTitleDim, Pos, WelcomeTitle
	
	; resize white background to fit all of GUI                                     
	GuiWidth 		:= Margin + WidestControlDimW + Margin
	GuiHeight 		:= LowestControlDimY + ExtraRowSpacing
	WhiteBoxHeight	:= GuiHeight - 2*VerticalMargin
	GuiControl, Move, WhiteBox, w%GuiWidth% h%WhiteBoxHeight%
 	GuiControl, Move, WelcomeTitle, % "x" GuiWidth/2 - WelcomeTitleDimW/2
	
	; show Welcome GUI
	Gui, +AlwaysOnTop -SysMenu
	
	ActiveMonitor 	:= GetCurrentMonitorIndex()
	AboutGuiX		:= CoordXCenterScreen(GuiWidth, ActiveMonitor)

	Gui, Welcome:Show, x%AboutGuiX% yCenter w%GuiWidth% h%GuiHeight%, %WelcomeTitle%
return

ShowInstallError:
	; create Install Error GUI
	Gui, InstallError:New
	
	; declare About GUI preferences
	FontName     	:= "Segoe UI"   
	FontSize     	:= 10       
	VerticalMargin 	:= 10            
	Margin   		:= 20       
	TextHeight		:= 17
	ButtonWidth  	:= 88      
	ButtonHeight 	:= 26 
	SS_WHITERECT 	:= 0x0006      

	ExtraRowSpacing	:= 2*TextHeight
	BottomHeight	:= ButtonHeight + 2*VerticalMargin
		
	; create a white background
	Gui, Add, Text, x0 y%VerticalMargin% %SS_WHITERECT% vInstallErrorWhiteBox
	
; --- GIVE ERROR MESSAGE [AREA 1]
	; section heading
	Gui, Font, s%FontSize%, %FontName%
	Gui, Font, NormGui, Font, Norm

	; add information
	Sentence1Y		:= 2*VerticalMargin
	Sentence1 		:= "Sorry!`n`nQuick Access Bar works best when installed locally.`nPlease copy Quick Access Bar.exe to your preferred local directory and try again.`n`nWatch the short Installation Guide clip for step-by-step instructions."
                
	Gui, Add, Text, x%Margin% y%Sentence1Y% BackgroundTrans vInstallErrorSentence1, %Sentence1%

	; add developer contact information
	GuiControlGet, Sentence1Dimensions, Pos, InstallErrorSentence1
	
	Sentence2 		:= "Still having issues? "	
	Sentence2Y		:= Sentence1DimensionsY + Sentence1DimensionsH
	Gui, Add, Text, x%Margin% y%Sentence2Y% BackgroundTrans vInstallErrorSentence2, %Sentence2%
	
	GuiControlGet, Sentence2Dimensions, Pos, InstallErrorSentence2

	LinkBoxX := Margin + Sentence2DimensionsW
	LinkBoxY := Sentence2DimensionsY

	Gui, Font, Underline
	Gui, Add, Link, x%LinkBoxX% y%LinkBoxY% BackgroundTrans vLinkBox, <a href="mailto:mqmotiwala@gmail.com?subject=Quick Access Bar - Install Error">Send me an email</a>
	Gui, Font, s%FontSize%, %FontName%
	Gui, Font, Norm

	; acquire link dimensions to position the following period
	GuiControlGet, LinkBoxDimensions, Pos, LinkBox
	PeriodX := LinkBoxDimensionsX + LinkBoxDimensionsW
	PeriodY := LinkBoxDimensionsY
	FinalText := "."
	Gui, Add, Text, x%PeriodX% y%PeriodY% BackgroundTrans, %FinalText%
	
; --- OVERALL ABOUT GUI CONFIGURATION & OK BUTTON [AREA 2]
	; acquire relevant control dimensions for GUI width and height calculations
	GuiControlGet, WidestControlDim, Pos, InstallErrorSentence1
	GuiControlGet, LowestControlDim, Pos, InstallErrorSentence2
	
	; resize white background to fit all of GUI
	WhiteBoxHeight 	:= LowestControlDimY + LowestControlDimH + VerticalMargin                                       
	GuiWidth 		:= Margin + WidestControlDimW + Margin
	
	GuiControl, Move, InstallErrorWhiteBox, w%GuiWidth% h%WhiteBoxHeight%
 
	; place OK button
	ButtonX := GuiWidth - Margin - ButtonWidth
	ButtonY := WhiteBoxHeight + (BottomHeight - ButtonHeight)*0.75
	
	Gui, Add, Button, x%ButtonX% y%ButtonY% w%ButtonWidth% h%ButtonHeight% gInstallErrorButtonOK Default, OK
	GuiControl, Focus, OK
	
	; show Install Error GUI
	InstallErrorGuiHeight := WhiteBoxHeight + BottomHeight

	ActiveMonitor 	 := GetCurrentMonitorIndex()
	InstallErrorGuiX := CoordXCenterScreen(GuiWidth, ActiveMonitor)
	
	Gui, +AlwaysOnTop +ToolWindow -SysMenu 
	Gui, Show, x%InstallErrorGuiX% yCenter w%GuiWidth% h%InstallErrorGuiHeight%, Please Install Locally!
Return

InstallErrorButtonOK:
	ExitApp
return