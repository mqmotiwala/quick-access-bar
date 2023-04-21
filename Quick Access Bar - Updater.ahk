#NoEnv, UseUnsetLocal
#Warn
#SingleInstance force
#Persistent
SendMode Event
SetWorkingDir %A_ScriptDir% 
SetBatchLines, -1
AutoTrim, On
OnExit("ExitFunction")

; ExitApp Codes
; 1 - After Successfully Updating
; 2 - No Update Required
; 3 - Generic Error Level
; 4 - Button OK pressed

PerformUpdate()
return

; ----------------------------------------------------------------------------------------------------------
; ---------------------------------------- MAIN UPDATER.EXE CODE -------------------------------------------
; ----------------------------------------------------------------------------------------------------------

PerformUpdate()
	{
		global 
			
		; define network paths
		LocalIniFilePath 	:= "QABInitializationFile.ini"
		ScratchFolderName	:= "snedir"
		LatestVersionExeFilePath	:= "\\idcdata01\Scratch\" . ScratchFolderName . "\Update\Quick Access Bar.exe"
		UpdateInfoNetworkFilePath 	:= "\\idcdata01\Scratch\" . ScratchFolderName . "\Update\LatestVersion.ini"
		UserStatsNetworkFilePath	:= "\\idcdata01\Scratch\" . ScratchFolderName . "\Users\UsageStatistics.ini"
		
		; check for versions
		IniRead, LatestVersion, %UpdateInfoNetworkFilePath%, Update, LatestVersion, 0
		IniRead, CurrentVersion, %LocalIniFilePath%, Version, VersionKey, 0
		
		; perform update steps if an update exists
		if (LatestVersion > CurrentVersion and FileExist(LatestVersionExeFilePath))
			{
				; ensures QAB.exe is closed 
				CloseStatus := CloseScript("Quick Access Bar.exe")
				if (CloseStatus == "unable to close")
					{	
						; use error level to avoid stopping the script if anything fails
						; proceed to GenericErrorLevel regardless though because CloseScript() failed
						Run, "Quick Access Bar.exe",, UseErrorLevel
						GoSub, GenericErrorLevel
					}
			
				FileDelete, Quick Access Bar.exe
		
				; copy and run updated QAB.exe file
				FileCopy, %LatestVersionExeFilePath%, Quick Access Bar.exe, True
								
				; acquire latest version number and patch notes
				IniRead, LatestVersion, %UpdateInfoNetworkFilePath%, Update, LatestVersion, 0

				; update version number locally
				IniWrite, %LatestVersion%, %LocalIniFilePath%, Version, VersionKey
				
				; update usage stats on network
				AcquireLocalUsageStats()
				IniWrite, %InstallationDate%, %UserStatsNetworkFilePath%, %A_Username%, InstallationDate
				IniWrite, %Activations%, %UserStatsNetworkFilePath%, %A_Username%, Activations
				IniWrite, %DailyAverage%, %UserStatsNetworkFilePath%, %A_Username%, DailyAverage
				IniWrite, %TimeSaved%, %UserStatsNetworkFilePath%, %A_Username%, TimeSaved
				IniWrite, %LatestVersion%, %UserStatsNetworkFilePath%, %A_Username%, Version
				
				FormatTime, LastUpdated,,ddMMMyyyy hh:mmtt
				IniWrite, %LastUpdated%, %UserStatsNetworkFilePath%, %A_Username%, LastUpdated
				
				; notify user of update as part of ExitFunction()
				; sets ExitCode to 1 - this is used to identify ExitApp was called after an update happened
				ExitApp, 1
			}
		else
			{
				MsgBox, 4160, Quick Access Bar - Updater, No update required.`n`nYou are using Version %CurrentVersion%, the latest and greatest Quick Access Bar currently has to offer!
				IfMsgBox, OK
					ExitApp, 2
			}
	}
return

AcquireLocalUsageStats()
	{
		global
		
		; read parameters
		IniRead, Activations, %LocalIniFilePath%, Usage, Activations, 0
		IniRead, InstallationDate, %LocalIniFilePath%, Usage, InstallationDate
		
		; declare seconds saved per use
		; need to match this with Quick Access Bar.exe code
		SecondsSavedPerUse := 5
		
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

GenericErrorLevel:
	FormatTime, LastUpdated,,ddMMMyyyy hh:mmtt
	IniWrite, %LastUpdated%, %UserStatsNetworkFilePath%, %A_Username%, LastUpdated	
	IniWrite, TRUE, %UserStatsNetworkFilePath%, %A_Username%, ErrorWhileUpdating	
	
	MsgBox, 4116, ERROR, There was an error while updating.`nNo action is required of you, continue enjoying Quick Access Bar.`n`nThe developer will be in touch with you soon to troubleshoot.`n`nIf you prefer, press Yes to send him an email now.
	ContactLink := "mailto:mufaddal.motiwala@hatch.com?subject=Quick%20Access%20Bar%20-%20Error%20While%20Updating"
	IfMsgBox, Yes
		run, %ContactLink%
	
	ExitApp, 3
return

; ----------------------------------------------------------------------------------------------------------
; --------------------------------------- UPDATE NOTIFICATION GUI ------------------------------------------
; ----------------------------------------------------------------------------------------------------------

ExitFunction(ExitReason, ExitCode)
	{
		global 
		
		; the first if will always run
		; QAB will show simultaneously with UpdateInfo GUI 
		; but the current drawback is that QAB will refresh after user presses OK
		if (ExitCode >= 1)
			{
				; use ExitCode != 3 to avoid recursion
				Run, "Quick Access Bar.exe",,UseErrorLevel
				If (ErrorLevel == "ERROR" and ExitCode != 3)
					GoSub, GenericErrorLevel
			}
		
		if (ExitCode == 1)
			{
				GoSub, NotifyUserOfUpdate
				
				; https://www.autohotkey.com/boards/viewtopic.php?f=76&t=74807
				; below code allows OK button to exit Updater.exe
				return true
			}
	}
return
	
NotifyUserOfUpdate:
	; create UpdateInfo GUI
	Gui, UpdateInfo:New

	; declare UpdateInfo GUI preferences
	FontName     	:= "Segoe UI"   
	FontSize     	:= 10       
	VerticalMargin 	:= 10            
	Margin   		:= 15     
	TextHeight		:= 17
	ButtonWidth  	:= 88      
	ButtonHeight 	:= 26 
	SS_WHITERECT 	:= 0x0006      

	ExtraRowSpacing	:= 2*TextHeight
	BottomHeight	:= ButtonHeight + Margin
	
	; create a white background
	Gui, Add, Text, x0 y%VerticalMargin% %SS_WHITERECT% vWhiteBox
	
;--- Update Information [AREA 1]
	; section heading
	Area1Y := 2*VerticalMargin
	Gui, Font, s11 w600 Underline, %FontName%
	Gui, Add, Text, x%Margin% y%Area1Y% BackgroundTrans vUpdateTitle, Psst! Quick Access Bar just updated to Version %LatestVersion%
	Gui, Font, Norm s%FontSize%

	; add Update information
	Sentence1	:= "No action is required from you. Continue enjoying Quick Access Bar!"
	Sentence2	:= "Here is what's new:"
	
	Sentence1Y 	:= Area1Y + ExtraRowSpacing
	Sentence2Y 	:= Sentence1Y + TextHeight
	
	Gui, Add, Text, x%Margin% y%Sentence1Y% BackgroundTrans vSentence1, %Sentence1%
	Gui, Add, Text, x%Margin% y%Sentence2Y% BackgroundTrans vSentence2, %Sentence2%
	
	IniRead, NewFeatures, %UpdateInfoNetworkFilePath%, NewFeatures
	IniRead, BugFixes, %UpdateInfoNetworkFilePath%, BugFixes
	IniRead, DeveloperComments, %UpdateInfoNetworkFilePath%, DeveloperComments
		
	Sentence3	 := "New Features"
	Sentence3Y	 := Sentence2Y + ExtraRowSpacing
	NewFeaturesY := Sentence3Y + TextHeight
	
	Gui, Font, w600
	Gui, Add, Text, x%Margin% y%Sentence3Y% BackgroundTrans vSentence3, %Sentence3%
	Gui, Font, Norm
	
	Gui, Add, Text, x%Margin% y%NewFeaturesY% BackgroundTrans vNewFeaturesText, %NewFeatures%
	GuiControlGet, NewFeaturesDim, Pos, NewFeaturesText
	
	Sentence4	:= "Bug Fixes"
	Sentence4Y	:= NewFeaturesDimY + NewFeaturesDimH + TextHeight
	BugFixesY 	:= Sentence4Y + TextHeight
	
	Gui, Font, w600
	Gui, Add, Text, x%Margin% y%Sentence4Y% BackgroundTrans vSentence4, %Sentence4%
	Gui, Font, Norm
	
	Gui, Add, Text, x%Margin% y%BugFixesY% BackgroundTrans vBugFixesText, %BugFixes%
	GuiControlGet, BugFixesDim, Pos, BugFixesText
	
	Sentence5		:= "Developer Comments"
	Sentence5Y		:= BugFixesDimY + BugFixesDimH + TextHeight
	DevCommentsY 	:= Sentence5Y + TextHeight
	
	Gui, Font, w600
	Gui, Add, Text, x%Margin% y%Sentence5Y% BackgroundTrans vSentence5, %Sentence5%
	Gui, Font, Norm
	
	Gui, Add, Text, x%Margin% y%DevCommentsY% BackgroundTrans vDevCommentsText, %DeveloperComments%
	
	GuiControlGet, DevCommentsDim, Pos, DevCommentsText
	LastSentenceY 	:= DevCommentsDimY + DevCommentsDimH + TextHeight
	LastSentence	:= "Reminder: Press Windows Key + W to launch Quick Access Bar."
	Gui, Font, Italic
	Gui, Add, Text, x%Margin% y%LastSentenceY% BackgroundTrans vLastSentence, %LastSentence%
	Gui, Font, Norm s%FontSize%
	
; --- OVERALL ABOUT GUI CONFIGURATION & OK BUTTON [AREA 4]
	; acquire relevant control dimensions for GUI width and height calculations
	GuiControlGet, Sentence1Dim, Pos, Sentence1
	GuiControlGet, LowestControlDim, Pos, LastSentence
	
	; widest control is Sentence1 if release notes are short
	WidestControlDimW := Max(Sentence1DimW, NewFeaturesDimW, BugFixesDimW, DevCommentsDimW)
	
	; resize white background to fit all of GUI
	WhiteBoxHeight 	:= LowestControlDimY + LowestControlDimH + VerticalMargin                                  
	GuiWidth 		:= WidestControlDimW + 2*Margin
		
	GuiControl, Move, WhiteBox, w%GuiWidth% h%WhiteBoxHeight%
 
	; place OK button
	ButtonX := GuiWidth - Margin - ButtonWidth
	ButtonY := WhiteBoxHeight + Margin*0.875
	
	Gui, Add, Button, x%ButtonX% y%ButtonY% w%ButtonWidth% h%ButtonHeight% gButtonOK Default, OK
	GuiControl, Focus, OK
	
	; show About GUI
	GuiHeight := WhiteBoxHeight + BottomHeight
	
	ActiveMonitor 	:= GetCurrentMonitorIndex()
	UpdateInfoX		:= CoordXCenterScreen(GuiWidth, ActiveMonitor)
	
	Gui, +AlwaysOnTop +ToolWindow -SysMenu 
	Gui, UpdateInfo:Show, x%UpdateInfoX% yCenter w%GuiWidth% h%GuiHeight%, Quick Access Bar - Update
Return

ButtonOK:
	Gui, UpdateInfo:Destroy
	ExitApp, 4
return

; ----------------------------------------------------------------------------------------------------------
; --------------------------------------- MISCELLANEOUS FUNCTIONS ------------------------------------------
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
	
CloseScript(Name)
	{
		DetectHiddenWindows On
		SetTitleMatchMode RegEx
		IfWinExist, i)%Name%.* ahk_class AutoHotkey
			{
				WinClose
				WinWaitClose, i)%Name%.* ahk_class AutoHotkey, , 2
				If ErrorLevel
					return "unable to close"
				else
					return "closed"
			}
		else
			return "not found"
	}