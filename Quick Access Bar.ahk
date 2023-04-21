#NoEnv, UseUnsetLocal
#Warn
#SingleInstance force
SendMode Event
SetWorkingDir %A_ScriptDir% 
SetBatchLines, -1
AutoTrim, On
Menu, Tray, Add, Show Quick Access Bar, TrayMenuShowQAB
Menu, Tray, Add, Edit Dynamic Buttons, TrayMenuDynamicButtons
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
	Gui, Add, Button, x%HorizontalMargin% y%VerticalMargin% w%EditLinksButtonWidth% gEditLinks, %EditLinksButtonName%
	Gui, Add, Button, x%AboutButtonX% y%AboutButtonY% w%AboutButtonWidth% gPressAbout, ?

	; add Hatch Website buttons
	Gui, Font, cWhite s%QABFontSize% w%BoldWeight%
	Gui, Add, Text, x%HorizontalMargin%, Hatch
	Gui, Font 

	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressTimeSheet, Timesheet
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSharePoint, SharePoint
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressVPN, Hatch VPN

	; add Special Buttons
	Gui, Font, cWhite s%QABFontSize% w%BoldWeight%
	Gui, Add, Text, x%HorizontalMargin%, Special Buttons
	Gui, Font

	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressCopiedLink, %CopiedAddressSpecialButtonName%
	Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% gPressSnippingTool, Snipping Tool

	; add Dynamic Buttons
	if(NamesArray.Length() != 0)
	{
		Gui, Font, cWhite s%QABFontSize% w%BoldWeight%
		Gui, Add, Text, x%HorizontalMargin%, Dynamic Buttons
		Gui, Font
	}

	Loop, 25
		Gui, Add, Button, x%HorizontalMargin% w%GeneralButtonWidth% v%A_Index% gPressDynamicButton Hidden, % NamesArray[A_Index]

	Loop, % NamesArray.Length()
		GuiControl, Show, % A_Index

	; show Quick Access Bar
	; identify monitor screen 
	ActiveMonitor 	:= GetCurrentMonitorIndex()
	QABX 			:= CoordXCenterScreen(QABWidth, ActiveMonitor)

	; move welcome GUI if needed
	WinMove, %WelcomeTitle%, , % QABX + 4*HorizontalMargin + GeneralButtonWidth
	
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
	run, iexplore.exe https://idcsapepp.corp.hatchglobal.com/irj/portal?NavigationTarget=navurl://a4c6004e4c9a20fd30f8437568091a55,, UseErrorLevel
	If (ErrorLevel == "ERROR")
		GoSub, GenericErrorLevel
	GoSub, CloseQAB
return

PressSharePoint:
	run, https://hatchengineering.sharepoint.com/,, UseErrorLevel
	If (ErrorLevel == "ERROR")
		GoSub, GenericErrorLevel
	GoSub, CloseQAB
return

PressVPN:
	IniRead, CurrentCheckbox4Status, %LocalIniFilePath%, Checkboxes, CheckBox4
	
	if (CurrentCheckbox4Status)
		run, C:\Program Files (x86)\F5 VPN\f5fpclientW.exe,, UseErrorLevel
	else
		run, C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Palo Alto Networks\GlobalProtect\GlobalProtectz.lnk,, UseErrorLevel
	
	If (ErrorLevel == "ERROR")
		{
			GoSub, GenericErrorLevel
			return
		}
		
	GoSub, CloseQAB
return

GenericErrorLevel:
	MsgBox, 4116, ERROR, Uh oh, there was an unusual error.`n`nWould you like to send me an email? I'll sort this out for ya!
	ContactLink := "mailto:mufaddal.motiwala@hatch.com?subject=Quick%20Access%20Bar - General Error"
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
			run, https://www.google.com/#q=%clipboard%,, UseErrorLevel
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
	Sentence0Y		:= Area1Y + TextHeight
	Sentence1Y		:= Sentence0Y + ExtraRowSpacing
	Sentence2Y		:= Sentence1Y + TextHeight
	Sentence3Y		:= Sentence2Y + ExtraRowSpacing
	
	Sentence0		:= "Press any button to go to its mapped address."
	Sentence1 		:= "You can also create your own buttons! To do that, press '" . EditLinksButtonName . "' on the Quick Access Bar."
	Sentence2 		:= "Any file, folder, software or website can be linked for quick access."
	Sentence3		:= "For general feedback, bug reports or new feature ideas - "
    
	Gui, Add, Text, x%Margin% y%Sentence0Y% BackgroundTrans vSentence0, %Sentence0%
	Gui, Add, Text, x%Margin% y%Sentence1Y% BackgroundTrans vSentence1, %Sentence1%
	Gui, Add, Text, x%Margin% y%Sentence2Y% BackgroundTrans vSentence2, %Sentence2%
	Gui, Add, Text, x%Margin% y%Sentence3Y% BackgroundTrans vSentence3, %Sentence3%

	; add developer contact information
	GuiControlGet, Sentence3Dimensions, Pos, Sentence3
	
	LinkBoxX := Margin + Sentence3DimensionsW
	LinkBoxY := Sentence3DimensionsY

	Gui, Font, Underline
	Gui, Add, Link, x%LinkBoxX% y%LinkBoxY% BackgroundTrans vLinkBox, <a href="mailto:mufaddal.motiwala@hatch.com?subject=Quick Access Bar - Feedback">contact Mufaddal Motiwala</a>
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
	Sentence4Y 	:= Area2Y + TextHeight
	Sentence5Y	:= Sentence4Y + TextHeight + ExtraRowSpacing
	
	Gui, Font, italic underline
	Gui, Add, Text, x%Margin% y%Sentence4Y% BackgroundTrans vSentence4A, %CopiedAddressSpecialButtonName%
	Gui, Add, Text, x%Margin% y%Sentence5Y% BackgroundTrans vSentence5A, Snipping Tool
	Gui, Font, Norm
	
	GuiControlGet, Sentence4ADimensions, Pos, Sentence4A
	GuiControlGet, Sentence5ADimensions, Pos, Sentence5A
	
	Sentence4B 	:= " will navigate to the currently copied content if it is a working address to a destination."
	Sentence4C 	:= "Otherwise, it will perform a Google search of the copied content."
	Sentence5B	:= " will begin a new snip with your default snip settings." 

	Sentence4CY := Sentence4ADimensionsY + TextHeight
	Sentence4BX := Sentence4ADimensionsW + Margin
	Sentence5BX := Sentence5ADimensionsW + Margin
	
	Gui, Add, Text, x%Sentence4BX% y%Sentence4Y% BackgroundTrans vSentence4B, %Sentence4B%
	Gui, Add, Text, x%Margin% y%Sentence4CY% BackgroundTrans vSentence4C, %Sentence4C%
	Gui, Add, Text, x%Sentence5BX% y%Sentence5Y% BackgroundTrans vSentence5B, %Sentence5B%

; --- STATISTICS [AREA 3]
	; section heading
	Area3Y := Sentence5Y + ExtraRowSpacing
	Gui, Font, s%FontSize% w600 Underline, %FontName%
	Gui, Add, Text, x%Margin% y%Area3Y% BackgroundTrans, Usage Statistics
	Gui, Font, Norm
	
	; add usage statistics
	AcquireLocalUsageStats()
	MinimumActivations := 60/SecondsSavedPerUse
	
	if (Activations == 0)
		Sentence6	:= "Usage statistics will be available once you start using Quick Access Bar."
	else if (DailyAverage >= 1 and DaysUsed >= 3 and Activations >= MinimumActivations)
		Sentence6	:= "All-Time Total:`t" . Activations . " activations`nDaily Average:`t" . DailyAverage . " activations`nTime Savings:`t" . TimeSaved
	else
		{	
			if (Activations == 1)
				Sentence6	:= "You've used Quick Access Bar " . Activations . " time. Check in later for more!"
			else
				Sentence6	:= "You've used Quick Access Bar " . Activations . " times. Check in later for more!"
		}
	
	Sentence6Y	:= Area3Y + TextHeight
	Gui, Add, Text, x%Margin% y%Sentence6Y% BackgroundTrans vSentence6, %Sentence6%
	
; --- USER SETTINGS [AREA 4]
	; section heading
	GuiControlGet, Sentence6Dim, Pos, Sentence6
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
	GuiControlGet, WidestControlDim, Pos, Sentence1
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
	
	CheckBox4Status1Text 	:= "My preferred VPN client is the default Hatch VPN client."
	CheckBox4Status0Text 	:= "My preferred VPN client is GlobalProtect."
	
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
	IniWrite, %InstallationDate%, %UserStatsNetworkFilePath%, %A_Username%, InstallationDate
	IniWrite, %Activations%, %UserStatsNetworkFilePath%, %A_Username%, Activations
	IniWrite, %DailyAverage%, %UserStatsNetworkFilePath%, %A_Username%, DailyAverage
	IniWrite, %TimeSaved%, %UserStatsNetworkFilePath%, %A_Username%, TimeSaved
	
	IniRead, CurrentVersion, %LocalIniFilePath%, Version, VersionKey, -1
	IniWrite, %CurrentVersion%, %UserStatsNetworkFilePath%, %A_Username%, Version
	
	FormatTime, LastUpdated,,ddMMMyyyy hh:mmtt
	IniWrite, %LastUpdated%, %UserStatsNetworkFilePath%, %A_Username%, LastUpdated
return

CheckForUpdates:
	; check for updates on weekdays between 10AM to 3PM only
	if (A_Hour >= 10 && A_Hour <= 15 && A_WDay >= 2 && A_WDay <= 6)
		{
			; acquire latest version number and patch notes
			IniRead, LatestVersion, %UpdateInfoNetworkFilePath%, Update, LatestVersion, 0
			
			; acquire current version
			IniRead, CurrentVersion, %LocalIniFilePath%, Version, VersionKey, 0
			
			; run Updater.exe if needed
			if (LatestVersion > CurrentVersion and FileExist(LatestVersionExeFilePath))
				{
					; ensures Updater.exe is installed and is the latest associated with current version
					;FileDelete, Quick Access Bar - Updater.exe
					FileInstall, Quick Access Bar - Updater.exe, Quick Access Bar - Updater.exe, True
					;Sleep, 1500 
					
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
FileInstall, Quick Access Bar - Updater.exe, Quick Access Bar - Updater.exe, True
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
		SetTimer, CheckForUpdates, 1800000

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
		
		; variables associated with Tutorial
		WelcomeTitle := "Quick Access Bar Tutorial"
		
		; variables associated with Quick Access Bar
		QABTitle := "Quick Access Bar"
		EditLinksButtonName := "Edit Dynamic Buttons"
		CopiedAddressSpecialButtonName := "Search Copied Item"
		
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
	Sentence2 		:= "Quick Access Bar allows you to keep frequently used files, folders, software and websites within close reach."
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
	GuiControlGet, WidestControlDim, Pos, WelcomeSentence2
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
	Sentence2Y		:= Sentence1Y + 2*ExtraRowSpacing + 2*TextHeight
	
	Sentence1 		:= "Sorry!`n`nQuick Access Bar works best when installed locally.`nPlease copy Quick Access Bar.exe to your preferred local directory and try again.`n`nWatch the short Installation Guide clip for step-by-step instructions."
	Sentence2 		:= "Still having issues? "
                
	Gui, Add, Text, x%Margin% y%Sentence1Y% BackgroundTrans vInstallErrorSentence1, %Sentence1%
	Gui, Add, Text, x%Margin% y%Sentence2Y% BackgroundTrans vInstallErrorSentence2, %Sentence2%

	; add developer contact information
	GuiControlGet, Sentence2Dimensions, Pos, InstallErrorSentence2
	
	LinkBoxX := Margin + Sentence2DimensionsW
	LinkBoxY := Sentence2DimensionsY

	Gui, Font, Underline
	Gui, Add, Link, x%LinkBoxX% y%LinkBoxY% BackgroundTrans vLinkBox, <a href="mailto:mufaddal.motiwala@hatch.com?subject=Quick Access Bar - Install Error">Send me an email</a>
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