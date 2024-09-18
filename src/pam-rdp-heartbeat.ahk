;#SingleInstance Ignore		; don't start if already running
#SingleInstance Force		; relaunch if already running
#Persistent
#NoEnv
#InstallKeybdHook
#UseHook
#WinActivateForce

SetTitleMatchMode RegEx
DetectHiddenWindows, On
SetWinDelay,-1
SetBatchLines,-1
SetWorkingDir %A_ScriptDir%
SetKeyDelay,0,50

;-------
; Information about the script 
SplitPath, A_ScriptFullPath,,,,gScriptName, 
gVersion:= "1.0.0"

global gProgramTitle:= "PAM RDP Heartbeat"
global gSelfPid:= DllCall("GetCurrentProcessId")
global gSelfPidHex:= Format("{:04x}",gSelfPid)

;-------
; Log files and more
;
global LOG_ERROR:= 1
global LOG_WARNING:= 2
global LOG_INFO:= 3
global LOG_DEBUG:= 4
global LOG_TRACE:= 5

FormatTime, timestamp ,, yyyyMMdd
;gLogFile= %gScriptPath%\%gScriptName%-%timestamp%.log
;global gLogFile:= "c:\windows\temp\pam-rdp-" timestamp ".log"
;global gLogFile:= A_Temp "\" gScriptName "-" timestamp ".log"
global gLogFile:= A_Temp "\" gScriptName ".log"
global gLogLevel:= LOG_DEBUG
global ErrorMessage:= ""

; roll log files before we begin
LogRoll(gLogFile,5,5)

LogInfo(A_Linenumber, "Start ----- gVersion= " gVersion)

;-----------------------------------------------------------------
; Default variables / load from registry
;-----------------------------------------------------------------
gosub SetVariableDefaults
gosub LoadVariablesFromRegistry
gosub TestScreenSaver

; global variables
global SessionList
global WindowList:= []
global clickActive= true

;-----------------------------------------------------------------
; Menus and GUI
;-----------------------------------------------------------------
Menu, Tray, NoStandard
;Menu, Tray, MainWindow
Menu, Tray, Add, Show, ProgramUnhide
Menu, Tray, Add
Menu, Tray, Add, Exit, GuiClose
Menu, Tray, Default, Show
Menu, Tray, Tip, %gProgramTitle%
Menu, Tray, Click, 1

Menu, ListMenu, Add, Show, menuDo

; --- 
; Application menu

Menu, FileMenu, Add, %mnuFileExitTxt%, mnuFileExitEvent

Menu, SettingsMenu, Add, %mnuSettingsStayOnTopTxt%, mnuSettingsStayOnTopEvent
Menu, SettingsMenu, Add, %mnuSettingsStartMinimizedTxt%, mnuSettingsStartMinimizedEvent
Menu, SettingsMenu, Add
Menu, SettingsMenu, Add, %mnuSettingsPreferencesTxt%, mnuSettingsPreferencesEvent

Menu, ActionMenu, Add, %mnuActionRefreshTxt%`tF5, mnuRefreshAllEvent
Menu, ActionMenu, Add, %mnuActionShowAllTxt%, mnuShowAllEvent
Menu, ActionMenu, Add, %mnuActionMinimizeAllTxt%, mnuMinimizeAllEvent

Menu, HelpMenu, Add, %mnuHelpQuickGuideTxt%, mnuHelpQuickGuideEvent
Menu, HelpMenu, Add
Menu, HelpMenu, Add, %mnuHelpAboutTxt%, mnuHelpAboutEvent

; Attach the sub-menus that were created above.
Menu, MyMenuBar, Add, %mnuFileTxt%, :FileMenu
Menu, MyMenuBar, Add, %mnuActionTxt%, :ActionMenu
Menu, MyMenuBar, Add, %mnuSettingsTxt%, :SettingsMenu
Menu, MyMenuBar, Add, %mnuHelpTxt%, :HelpMenu

if (StayOnTop)
	Menu, SettingsMenu, Check, %mnuSettingsStayOnTopTxt%
	
if (StartMinimized)
	Menu, SettingsMenu, Check, %mnuSettingsStartMinimizedTxt%

GuiOptions:= "+SysMenu +Owner -ToolWindow +MinimizeBox -MaximizeBox +Hwndmain_hWnd"
if (StayOnTop)
	GuiOptions:= "+AlwaysOnTop " GuiOptions
Gui, 1:New, %GuiOptions%
Gui, 1:Margin, %marginX%,%marginY%
Gui, 1:Add, listview, xm r%listRows% HWNDhSessionList vSessionList AltSubmit gSessionListEvent w%listWidth% -Multi -ReadOnly, %listHeader%
Gui, 1:Add, Button, section w%btnWidth% h%btnHeight% gbtnExitEvent vbtnExit, %btnExitTxt%
Gui, 1:Add, Button, ys w%btnWidth% h%btnHeight% Default gbtnCloseEvent vbtnClose, %btnCloseTxt%
Gui, 1:Menu, MyMenuBar

gosub, RefreshProcess
LV_Modifycol(listColTitleIdx,listColTitleWdt)
LV_Modifycol(listColPIDIdx,listColPIDWdt)
LV_Modifycol(listColStartIdx,listColStartWdt)
LV_Modifycol(listColDurationIdx,listColDurationWdt)
LV_Modifycol(listColStateIdx,listColStateWdt)
LV_Modifycol(listColHwndIdx,listColHwndWdt)

wmin:= 2*(btnWidth+2*marginX)
w:= listWidth+2*marginX
if (w < wmin)
	w:= wmin
Gui, 1:Show, w%w% Hide, %gProgramTitle%

; Validate GUI on screen
hwnd:= WinExist(gProgramTitle)
WinGetNormalPos(hwnd, x, y, w, h)
SysGet, MonitorWorkArea, MonitorWorkArea
if (StartX<0)
	StartX:= 0
if (StartX+w>MonitorWorkAreaRight)
	StartX:= MonitorWorkAreaRight-w
if (StartY<0)
	StartY:= 0
if (StartY+h>MonitorWorkAreaBottom)
	StartY:= MonitorWorkAreaBottom-h
WinMove, ahk_id %hwnd%,,%StartX%,%StartY%
guiHidden= 1
if (StartMinimized == 0) {
	WinShow, ahk_id %hwnd%
	WinActivate, ahk_id %hwnd%
	GuiControl, 1:Focus, btnClose
	guiHidden= 0
}

OnExit( "ExitFunc" )
; Set hook for minimizestart
SetEventHook(true)

; Set hook for WinActivate
SHELL_MSG := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK", "UInt")
OnMessage(SHELL_MSG, Func("ShellCallback"))
SetShellHook(true)

HeartbeatTime:= HeartbeatFrequency*1000
SetTimer, HeartbeatProcess,% HeartbeatTime
RefreshTime:= RefreshFrequency*1000
SetTimer, RefreshProcess,% RefreshTime

return

;-------------------------
GuiSize:
	if (A_EventInfo = 1) {
		GoSub btnCloseEvent
	}
	else {
		GuiControl, 1:Move, btnClose, % "x" (A_GuiWidth-btnWidth-marginX)
	}
	return

;-------------------------
; Show GUI - any window, including MSTSC
; alt-ctrl-scrolllock
!^ScrollLock::
	if (guiHidden) {
		WinShow, %gProgramTitle%
		WinActivate, %gProgramTitle%
		guiHidden:= 0
	}
	else {
		ifWinNotActive, %gProgramTitle% 
		{
			WinShow, %gProgramTitle%
			WinActivate, %gProgramTitle%
		}
		else {
			WinHide, %gProgramTitle%
			guiHidden:= 1
		}
	}
	return
	
#if WinActive(gProgramTitle)
; Only when Heartbeat GUI is active

; Refresh list
mnuRefreshAllEvent:
F5::
	Gosub RefreshProcess
	return

; Minimize all
; alt-ctrl-Down
mnuMinimizeAllEvent:
^!Down::
	; all mstsc.exe windows
	WinGet, allMstsc, List, ahk_class TscShellContainerClass
	Loop, %allMstsc% {
		hwnd:= allMstsc%A_Index%
		WinMinimize, ahk_id %hwnd%
		Sleep, 50
	}
	
	; all SymantecPAM applets
	WinGet, allSymantecPAM, List, ahk_exe CAPAMClient\.exe ahk_class SunAwtFrame,,(CA PAM LDAP Browser|CA Privileged Access Manager|Symantec PAM LDAP Browser|Symantec Privileged Access Manager)
	Loop, %allSymantecPAM% {
		hwnd:= allSymantecPAM%A_Index%
		WinMinimize, ahk_id %hwnd%
		Sleep, 50
	}
	
	WinActivate, ahk_id %main_hWnd%
	GuiControl, Focus, btnClose
	return

; Show all
; alt-ctrl-Up
mnuShowAllEvent:
^!Up::
	SetTimer, HeartbeatProcess, Off
	ShowAllMstscWindow()
	WinActivate, ahk_id %main_hWnd%
	GuiControl, Focus, btnClose
	SetTimer, HeartbeatProcess, On
	return

#if
	
;-------------------------
mnuFileExitEvent:
btnExitEvent:
GuiClose:
	SetTimer, RefreshProcess,Delete
	SetTimer, HeartbeatProcess,Delete
	GoSub SaveVariablesToRegistry
	ExitApp


;-------------------------
mnuHelpQuickGuideEvent:
	Gui, +OwnDialogs
	MsgBox, 0, % StrReplace(gProgramTitle " - " mnuHelpQuickGuideTxt, "&"), 
	(
PAM RDP Heartbeat
-----------------------
%QuickGuideTxt%
)
	return

;-------------------------
mnuHelpAboutEvent:
	Gui, +OwnDialogs
	MsgBox, 0x40, % StrReplace(gProgramTitle " - " mnuHelpAboutTxt, "&"), 
	(
PAM RDP Heartbeat
-----------------------
%AboutTxt%

Version %gVersion%
Copyright ©2024 PAM-Exchange
	)
	return

;-------------------------
ProgramUnhide:
	WinShow, %gProgramTitle%
	WinActivate, %gProgramTitle%
	return

;-------------------------
GuiEscape:
btnCloseEvent:
	clickActive:= false
	WinHide, %grogramTitle%
	nextHwnd:= NextWindow()
	WinActivate, ahk_id %nextHwnd%
	clickActive:= false
	return

;-------------------------
mnuSettingsStayOnTopEvent:
	StayOnTop:= !StayOnTop
	if (StayOnTop) {
		WinSet, AlwaysOnTop, On
		Menu, SettingsMenu, Check, %mnuSettingsStayOnTopTxt%
	}
	else {
		WinSet, AlwaysOnTop, Off
		Menu, SettingsMenu, Uncheck, %mnuSettingsStayOnTopTxt%
	}
	return
	
;-------------------------
mnuSettingsStartMinimizedEvent:
	StartMinimized:= !StartMinimized
	if (StartMinimized)
		Menu, SettingsMenu, Check, %mnuSettingsStartMinimizedTxt%
	else
		Menu, SettingsMenu, Uncheck, %mnuSettingsStartMinimizedTxt%
	return
	
;-------------------------
mnuSettingsPreferencesEvent:	
	SetTimer, RefreshProcess, Off
	SetTimer, HeartbeatProcess, Off

	w1:= GetTextSize(HeartbeatFrequencyTxt)
	w2:= GetTextSize(RefreshFrequencyTxt)
	edtX:= (w1<w2) ? w2 : w1
	edtX:= edtX+20
	
	WinGetPos,posX,posY,,,%gProgramTitle%
	posX:= posX+4*marginX
	posY:= posY+4*marginY
	
	Gui, 1: +Disabled
	GuiOptions:= "+SysMenu +Owner +AlwaysOnTop -ToolWindow -MinimizeBox -MaximizeBox +Hwndsettings_hWnd"
	Gui, 2:New, %GuiOptions%
	Gui, 2:Margin, %marginX%,%marginY%

	h:= 18
	w:= 30
	d1:= 5
	d2:= 4
	Gui, 2:Add, Text, xm ym+%d1% h%h%, %HeartbeatFrequencyTxt%
	Gui, 2:Add, Edit, xm+%edtX% ym+%d2% w%w% h%h% +Right +0x2000 vedtHeartbeatFreqency, %HeartbeatFrequency%

	delta:= round(2.5*marginY)
	d1:= d1+delta
	d2:= d2+delta
	Gui, 2:Add, Text, xm ym+%d1% h%h%, %RefreshFrequencyTxt%
	Gui, 2:Add, Edit, xm+%edtX%  ym+%d2% w%w% h%h% +Right +0x2000 vedtRefreshFreqency, %RefreshFrequency%

	Gui, 2:Add, CheckBox, xm vStayOnTop checked%StayOnTop%, %StayOnTopTxt%
	Gui, 2:Add, CheckBox, xm vStartMinimized checked%StartMinimized%, %StartMinimizedTxt%
	Gui, 2:Add, Text, xm h5, 
	
	Gui, 2:Add, Button, section xs w%btnWidth% h%btnHeight% gbtnCancelEvent vbtnCancel, %btnCancelTxt%
	Gui, 2:Add, Button, ys w%btnWidth% h%btnHeight% Default gbtnUpdateEvent vbtnUpdate, %btnUpdateTxt%
	
	Gui, 2:Show, x%posX% y%posY% Autosize, %mnuSettingsPreferencesTxt%
	Gui, 2:+LastFound
	WinWaitClose
	Gui, 1: -Disabled
	return

2GuiSize:
	GuiControl, 2:Move, btnUpdate, % "x" (A_GuiWidth-btnWidth-marginX)
	return

btnCancelEvent:
2GuiEscape:
2GuiClose:
	Gui, 2: Destroy
	Gosub RefreshProcess
	SetTimer, RefreshProcess, On
	SetTimer, HeartbeatProcess, On
	WinActivate, ahk_id %main_Hwnd%
	return
	
btnUpdateEvent:
	Gui, 2: Submit, NoHide
	
	; Stay On Top
	if (StayOnTop) {
		WinSet, AlwaysOnTop, On
		Menu, SettingsMenu, Check, %mnuSettingsStayOnTopTxt%
	}
	else {
		WinSet, AlwaysOnTop, Off
		Menu, SettingsMenu, Uncheck, %mnuSettingsStayOnTopTxt%
	}

	; Start Minimized
	if (StartMinimized)
		Menu, SettingsMenu, Check, %mnuSettingsStartMinimizedTxt%
	else
		Menu, SettingsMenu, Uncheck, %mnuSettingsStartMinimizedTxt%

	; RefreshFrequency
	edtRefreshFreqency:= edtRefreshFreqency+0
	If (edtRefreshFreqency < RefreshMinimum or strlen(edtRefreshFreqency)==0) {
		edtRefreshFreqency:= RefreshMinimum
		GuiControl,,edtRefreshFreqency, % edtRefreshFreqency
	}
	if (RefreshFrequency <> edtRefreshFreqency) {
		RefreshFrequency:= edtRefreshFreqency
		RefreshTime:= RefreshFrequency*1000
		SetTimer, RefreshProcess,% RefreshTime
	}
	
	; HeartbeatFrequency
	edtHeartbeatFreqency:= edtHeartbeatFreqency+0
	If (edtHeartbeatFreqency < HeartbeatMinimum or strlen(edtHeartbeatFreqency)==0) {
		edtHeartbeatFreqency:= HeartbeatMinimum
		GuiControl,,edtHeartbeatFreqency, % edtHeartbeatFreqency
	}
	if (HeartbeatFrequency <> edtHeartbeatFreqency) {
		HeartbeatFrequency:= edtHeartbeatFreqency
		HeartbeatTime:= HeartbeatFrequency*1000
		SetTimer, HeartbeatProcess,% HeartbeatTime
	}
	goto 2GuiClose

;-------------------------
UseTransparentEvent:
	GuiControlGet, UseTransparent,, UseTransparent
	return
	
;-------------------------
SessionListEvent:
	Gui,ListView,%A_GuiControl%
	switch A_GuiEvent {
	case "RightClick": 
		MouseGetPos, musX, musY
		Menu, ListMenu, Show, %musX%,%musY%
		return
	case "DoubleClick":
		gosub ShowMstscWindow
		return
	}
	return

;-------------------------
ShowMstscWindow:
	Gui,ListView,%A_GuiControl%
	RN:=LV_GetNext("Checked")
	LV_GetText(hwnd,RN,listColHwndIdx)
	ActivateWindow(hwnd)
	return
	
;-------------
menuDo:
	If (A_ThisMenuItem = "Show")
	   gosub,ShowMstscWindow
	return

;--------------------------------------------------------------------------------------------------------
RefreshProcess:
	SetTimer, RefreshProcess, Off
	Gui,Default
	Gui,ListView,SessionList
	GuiControl, -Redraw, SessionList

	LV_Delete()

	WinGet, allMstsc, List, ahk_class TscShellContainerClass
	Loop, %allMstsc% {
		hwnd:= allMstsc%A_Index%
		WinGet, pid, PID, ahk_id %hwnd%
		WinGetTitle, title, ahk_id %hwnd%

		ct:= ProcessCreationTime( pid )
		now:= A_Now
		delta:= now
		delta-= %ct%, seconds
		duration:= FormatSeconds(delta)
		createTime:= SubStr(ct,9,2) ":" SubStr(ct,11,2) ":" SubStr(ct,13,2)

		if (strlen(WindowList[hwnd]) == 0)
			WindowList[hwnd]= 0
		
		state:= WindowList[hwnd]
		lv_add("", title, pid, createTime, duration, state, hwnd)
		logTrace(A_Linenumber, "RefreshProcess: title= '" title "', pid= " pid ", start= " createTime ", duration= " duration ", state= " state ", hwnd= " hwnd)
	}

	; Symantec PAM applet
	WinGet, allSymantecPAM, List, ahk_exe CAPAMClient\.exe ahk_class SunAwtFrame,,(CA PAM LDAP Browser|CA Privileged Access Manager|Symantec PAM LDAP Browser|Symantec Privileged Access Manager)
	Loop, %allSymantecPAM% {
		hwnd:= allSymantecPAM%A_Index%
		WinGet, pid, PID, ahk_id %hwnd%
		WinGetTitle, title, ahk_id %hwnd%

		ct:= ProcessCreationTime( pid )
		now:= A_Now
		delta:= now
		delta-= %ct%, seconds
		duration:= FormatSeconds(delta)
		createTime:= SubStr(ct,9,2) ":" SubStr(ct,11,2) ":" SubStr(ct,13,2)

		if (strlen(WindowList[hwnd]) == 0)
			WindowList[hwnd]= 0
		
		state:= WindowList[hwnd]
		lv_add("", title, pid, createTime, duration, state, hwnd)
		logTrace(A_Linenumber, "RefreshProcess: title= " title ", pid= " pid ", start= " createTime ", duration= " duration ", state= " state ", hwnd= " hwnd)
	}
	
	LV_ModifyCol(1, "Sort")
	GuiControl, +Redraw, SessionList
	
	SetTimer, RefreshProcess, On
	return

;--------------------------------------------------------------------------------------------------------
HeartbeatProcess:
	SetTimer, HeartbeatProcess, Off		; stop timer

	; mstsc.exe programs
	WinGet, allMstsc, List, ahk_class TscShellContainerClass
	Loop, %allMstsc% {
		hwnd:= allMstsc%A_Index%
		
		; just in case the EVENT_SYSTEM_MINIMIZESTART was lost ...
		WinGet MMX, MinMax, ahk_id %hwnd%
		if (MMX == -1) {
			HideWindow(hwnd)
		}
		
		; Send heartbeat signal to mstsc windows (not minimized)
		SendMessage, 0x006, 1, 0,, ahk_id %hwnd%
		;WM_ACTIVATE(0x006)  WA_ACTIVE(1)
		logDebug(A_Linenumber, "HeartbeatProcess: (Mstsc) send heartbeat SendMessage to hwnd= " hwnd)
	}
	
	; Symantec PAM applet
	SetControlDelay -1
	WinGet, allSymantecPAM, List, ahk_exe CAPAMClient\.exe ahk_class SunAwtFrame,, (CA PAM LDAP Browser|CA Privileged Access Manager|Symantec PAM LDAP Browser|Symantec Privileged Access Manager)
	Loop, %allSymantecPAM% {
		hwnd:= allSymantecPAM%A_Index%
		if (WinActive("ahk_id " hwnd)) {
			logDebug(A_Linenumber, "HeartbeatProcess: (SymantecPAM Applet) send heartbeat ControlSend to hwnd= " hwnd)
			if (WinActive("ahk_id " hwnd) == 0) {
				logDebug(A_Linenumber, "HeartbeatProcess: (SymantecPAM Applet) ControlFocus hwnd= " hwnd)
				ControlFocus,,ahk_id %hwnd%
			}
			loop, 1 {
				ControlSend,,{NumLock},ahk_id %hwnd%
				ControlSend,,{ScrollLock},ahk_id %hwnd%
				;Sleep, 50
			}
		}
		/*
		else {
			logDebug(A_Linenumber, "HeartbeatProcess: (SymantecPAM Applet) send heartbeat ControlClick to hwnd= " hwnd)
			loop, 2 {
				;ControlFocus,,ahk_id %hwnd%
				;ControlClick, x30 y30, ahk_id %hwnd%,,,,NA 
				ControlClick, x200 y200, ahk_id %hwnd%,,,,NA POS
				Sleep, 10
			}
		}
		*/
	}
	
	SetTimer, HeartbeatProcess, On		; restart timer
	return

;--------------------------------------------------------------------------------------------------------
HideWindow(hwnd) 
{
	SetTimer, HeartbeatProcess, Off
	global WindowList
	global clickActive
	global UseTransparent

	Sleep, 100
	WinGet MMX, MinMax, ahk_id %hwnd%
	logDebug(A_LineNumber, "HideWindow: hwnd= " hwnd ", MMX= " MMX)
	if (MMX == -1) {
		; minimized
		clickActive:= false
		logDebug(A_LineNumber, "HideWindow: WinSet Transparent=0, hwnd= " hwnd)
		WinSet, Transparent, 0, ahk_id %hwnd%
		logDebug(A_LineNumber, "HideWindow: WinRestore, hwnd= " hwnd)
		WinRestore, ahk_id %hwnd%
		if (IsWindowFullScreen( hwnd )) {
			PostMessage, 0x112, 0xF120,,, ahk_id %hwnd%   ; 0x112 = WM_SYSCOMMAND, 0xF120 = SC_RESTORE
			WindowList[hwnd]:= 1	; 00001 - fullscreen
		}
		else 
			WindowList[hwnd]:= 2	; 00010 - window

		logDebug(A_LineNumber, "HideWindow: WinSet Bottom, hwnd= " hwnd)
		WinSet, Bottom,, ahk_id %hwnd%
	
		nextHwnd:= NextWindow()
		WinGetTitle, title, ahk_id %nextHwnd%
		logDebug(A_LineNumber, "HideWindow: nextHwnd= " nextHwnd ", title= " title)
		WinSet, Transparent, 255, ahk_id %nextHwnd%
		WinActivate, ahk_id %nextHwnd%
		Sleep, 50
		clickActive:= true
	}
	SetTimer, HeartbeatProcess, On		; restart timer
}
	
;--------------------------------------------------------------------------------------------------------
ActivateWindow(hwnd)
{
	clickActive:= false
	
	logDebug(A_LineNumber, "ActivateWindow: WinSet Transparent=255, hwnd= " hwnd)
	WinSet, Transparent, 255, ahk_id %hwnd%
	logDebug(A_LineNumber, "ActivateWindow: WinActivate, hwnd= " hwnd)
	WinActivate, ahk_id %hwnd%
	if (WindowList[hwnd] & 1) {
		Sleep,10
		PostMessage, 0x112, 0xF030,,, ahk_id %hwnd%   ; 0x112 = WM_SYSCOMMAND, 0xF030 = SC_MAXIMIZE
	}
	WindowList[hwnd]:= 0
	/*
	if (!UseTransparent) {
		;Sleep, 50
		;WinSet, Top,, ahk_id %hwnd%
		;HWND_TOPMOST:= -1
		;DllCall("SetWindowPos","UInt",hwnd,"UInt",HWND_TOPMOST,"Int",0,"Int",0,"Int",0,"Int",0,"UInt",0)
	}
	*/
	clickActive:= true
}

;--------------------------------------------------------------------------------------------------------
FormatSeconds(NumberOfSeconds)  ; Convert the specified number of seconds to hh:mm:ss format.
{
	t:= NumberOfSeconds+0
	hrs:= floor(t/3600)
	t:= t-hrs*3600
	min:= floor(t/60)
	t:= t-min*60
	sec:= t
	str:= Format("{:02}:{:02}:{:02}", hrs,min,sec)
	return str
}	

;--------------------------------------------------------------------------------------------------------
ProcessCreationTime( PID )  ; Requires AutoHotkey v1.0.46.03+
{

	VarSetCapacity(PrCT,16)
	VarSetCapacity(Dummy,16)
	VarSetCapacity(SysT,16)
	VarSetCapacity(SysT2,16)

	AccessRights := 1040       ; PROCESS_QUERY_INFORMATION = 1024,  PROCESS_VM_READ = 16
	hPr:=DllCall( "OpenProcess", Int,AccessRights, Int,0, Int,PID )
	DllCall( "GetProcessTimes" , Int,hPr, Int,&PrCT, Int,&Dummy, Int,&Dummy, Int,&Dummy)
	DllCall("CloseHandle"      , Int,hPr)

	DllCall( "FileTimeToLocalFileTime" , Int,&PrCT, Int,&PrCT )  ; PrCT is Creation time
	DllCall( "FileTimeToSystemTime"    , Int,&PrCt, Int,&SysT )  ; SysT is System Time

	Loop 16   {       ; Extracting and concatenating 8 words from a SYSTEMTIME structure
		Word := Mod(A_Index-1,2) ? "" :  *( &SysT +A_Index-1 ) + ( *(&SysT +A_Index) << 8 )
		Time .= StrLen(Word) = 1 ? ( "0" . Word ) : Word  ; Prefixing "0" for single digits
	} 

	Return SubStr(Time,1,6) . SubStr(Time,9,8) ; YYYYMMDD24MISS
}

;--------------------------------------------------------------------------------------------------------
WinGetNormalPos(hwnd, ByRef x, ByRef y, ByRef w="", ByRef h="")
{
    VarSetCapacity(wp, 44), NumPut(44, wp)
    DllCall("GetWindowPlacement", "uint", hwnd, "uint", &wp)
    x := NumGet(wp, 28, "int")
    y := NumGet(wp, 32, "int")
    w := NumGet(wp, 36, "int") - x
    h := NumGet(wp, 40, "int") - y
}

;--------------------------------------------------------------------------------------------------------
IsWindowFullScreen( hwnd ) {
	;checks if the specified window is full screen
	;winID := WinExist( winTitle )
	If ( !hwnd )
		Return false
	WinGet style, Style, ahk_id %hwnd%
	;WinGetPos ,,,winW,winH, %winTitle%
	WinGetPos ,,,winW,winH, ahk_id %hwnd%
	; 0x800000 is WS_BORDER.
	; 0x20000000 is WS_MINIMIZE.
	; no border and not minimized
	Return ((style & 0x20800000) or winH < A_ScreenHeight or winW < A_ScreenWidth) ? false : true
}

;--------------------------------------------------------------------------------------------------------
GetTextSize(pStr, pSize=8, pFont="", pHeight=false) {
   Gui 9:Font, %pSize%, %pFont%
   Gui 9:Add, Text, R1, %pStr%
   GuiControlGet T, 9:Pos, Static1
   Gui 9:Destroy
   Return pHeight ? TW "," TH : TW
}

;--------------------------------------------------------------------------------------------------------
ExitFunc(ExitReason, ExitCode) {
	SetTimer, RefreshProcess,Delete
	SetTimer, HeartbeatProcess,Delete
	SetShellHook(false)
	SetEventHook(false)
	ShowAllMstscWindow()
}

;--------------------------------------------------------------------------------------------------------
ShowAllMstscWindow() {
	clickActive:= false
	WinGet, allMstsc, List, ahk_class TscShellContainerClass
	Loop, %allMstsc% {
		hwnd:= allMstsc%A_Index%
		WinSet, Transparent, 255, ahk_id %hwnd%
		WinActivate, ahk_id %hwnd%
		Sleep, 50
	}
	
	; all SymantecPAM applets
	WinGet, allSymantecPAM, List, ahk_exe CAPAMClient\.exe ahk_class SunAwtFrame,,(CA PAM LDAP Browser|CA Privileged Access Manager|Symantec PAM LDAP Browser|Symantec Privileged Access Manager)
	Loop, %allSymantecPAM% {
		hwnd:= allSymantecPAM%A_Index%
		WinSet, Transparent, 255, ahk_id %hwnd%
		WinActivate, ahk_id %hwnd%
		Sleep, 50
	}

	clickActive:= false
}

;--------------------------------------------------------------------------------------------------------
; Sets whether the shell hook is registered
SetShellHook(state) {
    global main_hWnd
    static shellHookInstalled := false
    if (!shellHookInstalled and state) {
        if (!DllCall("RegisterShellHookWindow", "Ptr", main_hWnd)) {
            return false
        }
        shellHookInstalled := true
    }
    else if (shellHookInstalled and !state) {
        if (!DllCall("DeregisterShellHookWindow", "Ptr", main_hWnd)) {
            return false
        }
        shellHookInstalled := false
    }
    return true
}

;--------------------------------------------------------------------------------------------------------
; Shell messages callback
ShellCallback(wParam, lParam) {

	global clickActive

	if (wParam == 0x4 or wParam == 0x8004) {
		; HSHELL_WINDOWACTIVATED = 4, HSHELL_RUDEAPPACTIVATED = 0x8004
		WinGetClass, class, ahk_id %lParam%
		if (class == "TscShellContainerClass") {
			; lParam = hWnd of activated window
			;MsgBox, ShellCallback, class= %class%`nclickActive= %clickActive%`nlParam= %lParam%
			if (clickActive) {
				logDebug(A_LineNumber, "ShellCallback: lParam= " lParam)
				ActivateWindow(lParam)
				WinSet, Transparent, 255, ahk_id %lParam%
			}
			return false
		} 
	} 
}

;--------------------------------------------------------------------------------------------------------
EventHookProc( hWinEventHook, Event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime )
{
	Event += 0

	; 0x16 EVENT_SYSTEM_MINIMIZESTART
	if (Event == 0x16) {
		WinGetClass, class, ahk_id %hwnd%
		logDebug(A_LineNumber, "EventHookProc: event= 0x16, class= " class ", hwnd= " hwnd)
		if (class == "TscShellContainerClass") {
			HideWindow(hwnd) 
			return false
		}
	}
}

;--------------------------------------------------------------------------------------------------------
SetWinEventHook(eventMin, eventMax, hmodWinEventProc, lpfnWinEventProc, idProcess, idThread, dwFlags) 
{
	DllCall("CoInitialize", Uint, 0)
	return DllCall("SetWinEventHook"
	, Uint,eventMin	
	, Uint,eventMax	
	, Uint,hmodWinEventProc
	, Uint,lpfnWinEventProc
	, Uint,idProcess
	, Uint,idThread
	, Uint,dwFlags)	
}

;--------------------------------------------------------------------------------------------------------
SetEventHook(state)
{
    static eventHookInstalled := false
	static hWinEventHook
	static EventHookProcAdr
	
    if (!eventHookInstalled and state) {
		ExcludeScriptMessages = 1 ; 0 to include
		EventHookProcAdr := RegisterCallback( "EventHookProc", "F" )
		dwFlags := ( ExcludeScriptMessages = 1 ? 0x1 : 0x0 )
		hWinEventHook := SetWinEventHook( 0x16, 0x16, 0, EventHookProcAdr, 0, 0, 0 )
	}
    else if (eventHookInstalled and !state) {
		DllCall( "UnhookWinEvent", Uint,hWinEventHook )
		DllCall( "GlobalFree", UInt,&EventHookProcAdr ) ; free up allocated memory for RegisterCallback
        eventHookInstalled := false
    }
    return true
}

;--------------------------------------------------------------------------------------------------------
ReadRegistry(keyName,valueName,defaultValue=0)
{
	tmp:= 0
	RegRead, tmp, %keyName%, %valueName%
	
	if (strlen(tmp)>0) {
		return %tmp%
	}
	return %defaultValue%
}

;--------------------------------------------------------------------------------------------------------
WriteRegistry(type,keyName,valueName,value)
{
	RegWrite, %type%, %keyName%, %valueName%, %value%
}

;--------------------------------------------------------------------------------------------------------
SetVariableDefaults:
	marginX:= 10
	marginY:= 10
	btnWidth:= 0
	btnHeight:= 26
	StartX:= 10
	StartY:= 10
	language:= "en-US"

	HeartbeatFrequency:= 60
	HeartbeatFrequencyTxt:= "&Heartbeat frequency (sec)"
	HeartbeatMinimum:= 10
	RefreshFrequency:= 10
	RefreshFrequencyTxt:= "&Refresh frequency (sec)"
	RefreshMinimum:= 5
	StartMinimized:= 0
	StartMinimizedTxt:= "Start &Minimized"
	StayOnTop:= 1
	StayOnTopTxt:= "Stay on &Top"
	UseTransparent:= 1
	UseTransparentTxt:= "Use Transparent"
	ScreenSaverRequired:= 1
	ScreenSaverIdleMaximum:= 1800	; 30 minutes

	btnExitTxt:= "E&xit"
	btnCloseTxt:= "&Close"
	btnUpdateTxt:= "&Update"
	btnCancelTxt:= "&Cancel"

	listRows:= 10
	listColTitleTxt:= "Title"
	listColTitleWdt:= 170
	listColPIDTxt:= "PID"
	listColPIDWdt:= 50
	listColStartTxt:= "Start"
	listColStartWdt:= 55
	listColDurationTxt:= "Duration"
	listColDurationWdt:= 55
	listColStateTxt:= "State"
	listColStateWdt:= 0
	listColHwndTxt:= "HWND"
	listColHwndWdt:= 0

	listColTitleIdx:= 1
	listColPIDIdx:= 2
	listColStartIdx:= 3
	listColDurationIdx:= 4
	listColStateIdx:= 5
	listColHwndIdx:= 6
	
	mnuFileTxt:= "&File"
	mnuFileExitTxt:= "E&xit"

	mnuSettingsTxt:= "&Settings"
	mnuSettingsStayOnTopTxt:= "Stay on &Top"
	mnuSettingsStartMinimizedTxt:= "Start &Minimized"
	mnuSettingsPreferencesTxt:= "&Preferences"

	mnuActionTxt:= "&Action"
	mnuActionRefreshTxt:= "&Refresh"
	mnuActionShowAllTxt:= "&Show All"
	mnuActionMinimizeAllTxt:= "&Minimize All"

	mnuHelpTxt:= "&Help"
	mnuHelpQuickGuideTxt:= "Quick &guide"
	mnuHelpAboutTxt:= "&About"
	
	AboutTxt:= "The program sends heartbeat signals to RDP sessions. When RDP sessions are minimized, the RDP GUI is hidden and heartbeat messages are still being accepted by the session."
	QuickGuideTxt:= "In the menu item 'Settings > Preferences' timers for background actions can be defined. The heartbeat timer will control how often the heartbeat is send to both active and hidden RDP sessions. The value should depend of the timeout values defined on the servers connected to from the Desktop. Typically, sending a heartbeat signal every 30'th second is sufficient. The refresh timer will control how often the list of RDP sessions is updated. Typical values are every 10 to 30 seconds. It is also possible to refress the list manually by pressiong F5. The flags 'Stay on Top' and 'Start Minimized' are controlling the PAM RDP Heartbeat window behaviour.\n\nAlt-Ctrl-ScrollLock\nWill show/activate or hide the GUI window. This will work even when showing an RDP fullscreen session.\n\nWhen the PAM RDP Heartbeat GUI is active, additional functions are available.\n\nAlt-Ctrl-Down\nMinimize/hide all RDP windows.\n\nAlt-Ctrl-Up\nRestore all hidden RDP windows.\n\nF5\nManually refresh the list of RDP sessions."
	
	AboutTxt:= StrReplace(AboutTxt,"\n","`n")
	QuickGuideTxt:= StrReplace(QuickGuideTxt,"\n","`n")
	return

;--------------------------------------------------------------------------------------------------------
LoadVariablesFromRegistry:
	key:= "HKCU\Software\PAM-Exchange\PAM-RDP-Heartbeat"
	gLogLevel:= 					ReadRegistry(key, "LogLevel", gLogLevel)
	logInfo(A_Linenumber, "LoadVariablesFromRegistry: LogLevel= " gLogLevel)
	
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: RegistryKey= " key)
	HeartbeatFrequency:= 			ReadRegistry(key, "HeartbeatFrequency", HeartbeatFrequency)
	HeartbeatMinimum:= 	 			ReadRegistry(key, "HeartbeatMinimum", HeartbeatMinimum)
	listColDurationWdt:= 			ReadRegistry(key, "listColDurationWdt", listColDurationWdt)
	listColHwndWdt:= 				ReadRegistry(key, "listColHwndWdt", listColHwndWdt)
	listColPIDWdt:= 				ReadRegistry(key, "listColPIDWdt", listColPIDWdt)
	listColStartWdt:= 				ReadRegistry(key, "listColStartWdt", listColStartWdt)
	listColStateWdt:= 				ReadRegistry(key, "listColStateWdt", listColStateWdt)
	listColTitleWdt:= 				ReadRegistry(key, "listColTitleWdt", listColTitleWdt)
	listRows:= 			 			ReadRegistry(key, "listRows", listRows)
	RefreshFrequency:= 	 			ReadRegistry(key, "RefreshFrequency", RefreshFrequency)
	RefreshMinimum:= 				ReadRegistry(key, "RefreshMinimum", RefreshMinimum)
	StartMinimized:= 	 			ReadRegistry(key, "StartMinimized", StartMinimized)
	StartX:= 			 			ReadRegistry(key, "StartX", StartX)
	StartY:= 			 			ReadRegistry(key, "StartY", StartY)
	StayOnTop:= 		 			ReadRegistry(key, "StayOnTop", StayOnTop)
	UseTransparent:=	 			ReadRegistry(key, "UseTransparent", UseTransparent)
	language:= 						ReadRegistry(key, "Language", language)

	ScreenSaverRequired:=			ReadRegistry(key, "ScreenSaverRequired", ScreenSaverRequired)
	ScreenSaverIdleMaximum:=		ReadRegistry(key, "ScreenSaverIdleMaximum", ScreenSaverIdleMaximum)

	key:= key . "\" . language
	btnCancelTxt:= 		 			ReadRegistry(key, "btnCancelTxt", btnCancelTxt)
	btnCloseTxt:= 		 			ReadRegistry(key, "btnCloseTxt", btnCloseTxt)
	btnExitTxt:= 		 			ReadRegistry(key, "btnExitTxt", btnExitTxt)
	btnUpdateTxt:= 		 			ReadRegistry(key, "btnUpdateTxt", btnUpdateTxt)
	HeartbeatFrequencyTxt:=			ReadRegistry(key, "HeartbeatFrequencyTxt", HeartbeatFrequencyTxt)
	listColDurationTxt:= 			ReadRegistry(key, "listColDurationTxt", listColDurationTxt)
	listColHwndTxt:= 				ReadRegistry(key, "listColHwndTxt", listColHwndTxt)
	listColPIDTxt:= 				ReadRegistry(key, "listColPIDTxt", listColPIDTxt)
	listColStartTxt:= 				ReadRegistry(key, "listColStartTxt", listColStartTxt)
	listColStateTxt:= 				ReadRegistry(key, "listColStateTxt", listColStateTxt)
	listColTitleTxt:= 				ReadRegistry(key, "listColTitleTxt", listColTitleTxt)
	mnuActionMinimizeAllTxt:= 		ReadRegistry(key, "mnuActionMinimizeAllTxt", mnuActionMinimizeAllTxt)
	mnuActionRefreshTxt:= 			ReadRegistry(key, "mnuActionRefreshTxt", mnuActionRefreshTxt)
	mnuActionShowAllTxt:= 			ReadRegistry(key, "mnuActionShowAllTxt", mnuActionShowAllTxt)
	mnuActionTxt:= 					ReadRegistry(key, "mnuActionTxt", mnuActionTxt)
	mnuFileExitTxt:= 				ReadRegistry(key, "mnuFileExitTxt", mnuFileExitTxt)
	mnuFileTxt:= 					ReadRegistry(key, "mnuFileTxt", mnuFileTxt)
	mnuHelpAboutTxt:= 				ReadRegistry(key, "mnuHelpAboutTxt", mnuHelpAboutTxt)
	mnuHelpQuickGuideTxt:= 			ReadRegistry(key, "mnuHelpQuickGuideTxt", mnuHelpQuickGuideTxt)
	mnuHelpTxt:= 					ReadRegistry(key, "mnuHelpTxt", mnuHelpTxt)
	mnuSettingsPreferencesTxt:= 	ReadRegistry(key, "mnuSettingsPreferencesTxt", mnuSettingsPreferencesTxt)
	mnuSettingsStartMinimizedTxt:= 	ReadRegistry(key, "mnuSettingsStartMinimizedTxt", mnuSettingsStartMinimizedTxt)
	mnuSettingsStayOnTopTxt:= 		ReadRegistry(key, "mnuSettingsStayOnTopTxt", mnuSettingsStayOnTopTxt)
	mnuSettingsTxt:= 				ReadRegistry(key, "mnuSettingsTxt", mnuSettingsTxt)

	AboutTxt:= 						ReadRegistry(key, "AboutTxt", AboutTxt)
	QuickGuideTxt:= 				ReadRegistry(key, "QuickGuideTxt", QuickGuideTxt)

	RefreshFrequencyTxt:=			ReadRegistry(key, "RefreshFrequencyTxt", RefreshFrequencyTxt)
	StartMinimizedTxt:=  			ReadRegistry(key, "StartMinimizedTxt", StartMinimizedTxt)
	StayOnTopTxt:= 		 			ReadRegistry(key, "StayOnTopTxt", StayOnTopTxt)
	UseTransparentTxt:=	 			ReadRegistry(key, "UseTransparentTxt", UseTransparentTxt)


	; ------------------------------
	; Sanitize variables read from registry

	AboutTxt:= StrReplace(AboutTxt,"\n","`n")
	QuickGuideTxt:= StrReplace(QuickGuideTxt,"\n","`n")

	if (StayOnTop>1)
		StayOnTop:= 1
		
	if (StartMinimized>1)
		StartMinimized:= 1

	if (UseTransparent>1)
		UseTransparent:= 1

	if (StartX>0x0000FFFF)
		StartX:= 0	; negative
		
	if (StartY>0x0000FFFF)
		StartY:= 0	; negative

	if (HeartbeatFrequency < HeartbeatMinimum)
		HeartbeatFrequency:= HeartbeatMinimum

	if (RefreshFrequency < RefreshMinimum)
		RefreshFrequency:= RefreshMinimum

	w:= GetTextSize(btnExitTxt)
	if (w>btnWidth)
		btnWidth:= w
	w:= GetTextSize(btnCloseTxt)
	if (w>btnWidth)
		btnWidth:= w
	w:= GetTextSize(btnCancelTxt)
	if (w>btnWidth)
		btnWidth:= w
	w:= GetTextSize(btnUpdateTxt)
	if (w>btnWidth)
		btnWidth:= w
	btnWidth:= btnWidth+round(1.5*marginX)

	listWidth:= 7+listColTitleWdt+listColPIDWdt+listColStartWdt+listColDurationWdt+listColStateWdt+listColHwndWdt
	listHeader:= listColTitleTxt "|" listColPIDTxt "|" listColStartTxt "|" listColDurationTxt "|" listColStateTxt "|" listColHwndTxt

	logDebug(A_Linenumber, "LoadVariablesFromRegistry: HeartbeatFrequency= " HeartbeatFrequency)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: HeartbeatMinimum= " HeartbeatMinimum)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: RefreshFrequency= " RefreshFrequency)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: RefreshMinimum= " RefreshMinimum)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: StartMinimized= " StartMinimized)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: StartX= " StartX)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: StartY= " StartY)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: StayOnTop= " StayOnTop)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: UseTransparent= " UseTransparent)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: btnWidth= " btnWidth)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: listColDurationWdt= " listColDurationWdt)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: listColHwndWdt= " listColHwndWdt)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: listColPIDWdt= " listColPIDWdt)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: listColStartWdt= " listColStartWdt)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: listColStateWdt= " listColStateWdt)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: listColTitleWdt= " listColTitleWdt)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: listRows= " listRows)
	logDebug(A_Linenumber, "LoadVariablesFromRegistry: listWidth= " listWidth)

	;MsgBox, listWidth= %listWidth%`nlistHeader= %listHeader%`nlistColTitleWdt= %listColTitleWdt%, listColTitleTxt= %listColTitleTxt%`nlistColPIDWdt= %listColPIDWdt%, listColPIDTxt= %listColPIDTxt%`nlistColStartWdt= %listColStartWdt%, listColStartTxt= %listColStartTxt%`nlistColDurationWdt= %listColDurationWdt%, listColDurationTxt= %listColDurationTxt%`nlistColStateWdt= %listColStateWdt%, listColStateTxt= %listColStateTxt%
	return
	
;--------------------------------------------------------------------------------------------------------
SaveVariablesToRegistry:
	WinGetPos,posX,posY,,,%gProgramTitle%
	key:= "HKCU\Software\PAM-Exchange\PAM-RDP-Heartbeat"

	logDebug(A_Linenumber, "SaveVariablesToRegistry: RegistryKey= " key)

	WriteRegistry("REG_DWORD", key, "LogLevel", gLogLevel)
	WriteRegistry("REG_DWORD", key, "HeartbeatFrequency", HeartbeatFrequency)
	WriteRegistry("REG_DWORD", key, "HeartbeatMinimum", HeartbeatMinimum)
	WriteRegistry("REG_DWORD", key, "listColDurationWdt", listColDurationWdt)
;	WriteRegistry("REG_DWORD", key, "listColHwndWdt", listColHwndWdt)
	WriteRegistry("REG_DWORD", key, "listColPIDWdt", listColPIDWdt)
	WriteRegistry("REG_DWORD", key, "listColStartWdt", listColStartWdt)
	WriteRegistry("REG_DWORD", key, "listColStateWdt", listColStateWdt)
	WriteRegistry("REG_DWORD", key, "listColTitleWdt", listColTitleWdt)
	WriteRegistry("REG_DWORD", key, "RefreshFrequency", RefreshFrequency)
	WriteRegistry("REG_DWORD", key, "RefreshMinimum", RefreshMinimum)
	WriteRegistry("REG_DWORD", key, "StartMinimized", StartMinimized)
	WriteRegistry("REG_DWORD", key, "StartX", posX)
	WriteRegistry("REG_DWORD", key, "StartY", posY)
	WriteRegistry("REG_DWORD", key, "StayOnTop", StayOnTop)
;	WriteRegistry("REG_DWORD", key, "UseTransparent", UseTransparent)
	WriteRegistry("REG_SZ",    key, "Language", language)

	key:= key . "\" . language
	logDebug(A_Linenumber, "SaveVariablesToRegistry: RegistryKey= " key)

	WriteRegistry("REG_SZ", key, "btnCancelTxt", btnCancelTxt)
	WriteRegistry("REG_SZ", key, "btnCloseTxt", btnCloseTxt)
	WriteRegistry("REG_SZ", key, "btnExitTxt", btnExitTxt)
	WriteRegistry("REG_SZ", key, "btnUpdateTxt", btnUpdateTxt)
	WriteRegistry("REG_SZ", key, "HeartbeatFrequencyTxt", HeartbeatFrequencyTxt)
	WriteRegistry("REG_SZ", key, "listColDurationTxt", listColDurationTxt)
;	WriteRegistry("REG_SZ", key, "listColHwndTxt", listColHwndTxt)
	WriteRegistry("REG_SZ", key, "listColPIDTxt", listColPIDTxt)
	WriteRegistry("REG_SZ", key, "listColStartTxt", listColStartTxt)
	WriteRegistry("REG_SZ", key, "listColStateTxt", listColStateTxt)
	WriteRegistry("REG_SZ", key, "listColTitleTxt", listColTitleTxt)
	WriteRegistry("REG_SZ", key, "mnuActionMinimizeAllTxt", mnuActionMinimizeAllTxt)
	WriteRegistry("REG_SZ", key, "mnuActionRefreshTxt", mnuActionRefreshTxt)
	WriteRegistry("REG_SZ", key, "mnuActionShowAllTxt", mnuActionShowAllTxt)
	WriteRegistry("REG_SZ", key, "mnuActionTxt", mnuActionTxt)
	WriteRegistry("REG_SZ", key, "mnuFileExitTxt", mnuFileExitTxt)
	WriteRegistry("REG_SZ", key, "mnuFileTxt", mnuFileTxt)
	WriteRegistry("REG_SZ", key, "mnuHelpAboutTxt", mnuHelpAboutTxt)
	WriteRegistry("REG_SZ", key, "mnuHelpQuickGuideTxt", mnuHelpQuickGuideTxt)
	WriteRegistry("REG_SZ", key, "mnuHelpTxt", mnuHelpTxt)
	WriteRegistry("REG_SZ", key, "mnuSettingsPreferencesTxt", mnuSettingsPreferencesTxt)
	WriteRegistry("REG_SZ", key, "mnuSettingsStartMinimizedTxt", mnuSettingsStartMinimizedTxt)
	WriteRegistry("REG_SZ", key, "mnuSettingsStayOnTopTxt", mnuSettingsStayOnTopTxt)
	WriteRegistry("REG_SZ", key, "mnuSettingsTxt", mnuSettingsTxt)
	WriteRegistry("REG_SZ", key, "RefreshFrequencyTxt", RefreshFrequencyTxt)
	WriteRegistry("REG_SZ", key, "StartMinimizedTxt", StartMinimizedTxt)
	WriteRegistry("REG_SZ", key, "StayOnTopTxt", StayOnTopTxt)
;	WriteRegistry("REG_SZ", key, "UseTransparentTxt", UseTransparentTxt)

	QuickGuideTxt:= StrReplace(QuickGuideTxt,"`n","\n")
	WriteRegistry("REG_SZ", key, "QuickGuideTxt", QuickGuideTxt)

	AboutTxt:= StrReplace(AboutTxt,"`n","\n")
	WriteRegistry("REG_SZ", key, "AboutTxt", AboutTxt)

	return

;--------------------------------------------------------------------------------------------------------
TestScreenSaver:
	; Test if a screensaver is enabled and secure
	if (!ScreenSaverRequired) {
		logInfo(A_LineNumber, "TestScreenSaver: ScreenSaver on user's desktop is not required")
		return
	}

	; Machine inactivity timeout
	ssMachineOK:= false
	ssInactivityTimeout:= ReadRegistry("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "InactivityTimeoutSecs", 0)
	if (ssInactivityTimeout != 0)
		ssMachineOK:= true

	; local ScreenSaver
	ssLocalOK:= false
	ssLocalActive:=		ReadRegistry("HKCU\Control Panel\Desktop", "ScreenSaveActive", 0)
	ssLocalIsSecure:=	ReadRegistry("HKCU\Control Panel\Desktop", "ScreenSaverIsSecure", 0)
	ssLocalTimeOut:=	ReadRegistry("HKCU\Control Panel\Desktop", "ScreenSaveTimeOut", 0)
	if (ssLocalActive == 1 and ssLocalIsSecure == 1 and ssLocalTimeOut <= ScreenSaverIdleMaximum)
		ssLocalOK:= true

	; GPO ScreenSaver
	ssGpoOK:= false
	ssGpoActive:=	ReadRegistry("HKCU\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "ScreenSaveActive", 0)
	ssGpoIsSecure:=	ReadRegistry("HKCU\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "ScreenSaverIsSecure", 0)
	ssGpoTimeOut:=	ReadRegistry("HKCU\SOFTWARE\Policies\Microsoft\Windows\Control Panel\Desktop", "ScreenSaveTimeOut", 0)
	if (ssGpoActive == 1 and ssGpoIsSecure == 1 and ssGpoTimeOut <= ScreenSaverIdleMaximum)
		ssGpoOK:= true

	logInfo(A_LineNumber, "TestScreenSaver: ssInactivityTimeout= " ssInactivityTimeout ", ssLocalActive= " ssLocalActive ", ssLocalIsSecure= " ssLocalIsSecure ", ssLocalTimeOut= " ssLocalTimeOut ", ssGpoOK= " ssGpoOK ", ssGpoActive= " ssGpoActive ", ssGpoIsSecure= " ssGpoIsSecure ", ssGpoTimeOut= " ssGpoTimeOut)
	if (!ssLocalOK and !ssGpoOK and !ssMachineOK) {
		LogError(A_LineNumber, "TestScreenSaver: Active screensaver not found . PAM-RDP-Heartbeat is not started")
		MsgBox, 0x1010, %gProgramTitle%, An active password protected ScreenSaver is not identified on the desktop. If defined, the timout may be set beyond the maximum of 30 minutes. This is a mandatory requirement when using the PAM RDP Heartbeat program.`n`nPAM RDP Heartbeat is not started.
		ExitApp
	}
	return

;--------------------------------------------------------------------------------------------------------
NextWindow() {
	HiddenWindows:= A_DetectHiddenWindows

	WS_EX_APPWINDOW = 0x40000 ; provides a taskbar button
	WS_EX_TOOLWINDOW = 0x80 ; removes the window from the alt-tab list
	GW_OWNER = 4

	AltTabListID:= []
	AltTabTotalNum := 0 ; the number of windows found
	DetectHiddenWindows, Off ; makes DllCall("IsWindowVisible") unnecessary
	WinGet, winList, List ; gather a list of running programs
	Loop, %winList%
	{
		ownerID := windowID := winList%A_Index%
		Loop {
			ownerID := Decimal_to_Hex( DllCall("GetWindow", "UInt", ownerID, "UInt", GW_OWNER))
		} Until !Decimal_to_Hex( DllCall("GetWindow", "UInt", ownerID, "UInt", GW_OWNER))
		ownerID := ownerID ? ownerID : windowID
		If (Decimal_to_Hex(DllCall("GetLastActivePopup", "UInt", ownerID)) = windowID)
		{
			WinGet, es, ExStyle, ahk_id %windowID%
			If (!((es & WS_EX_TOOLWINDOW) && !(es & WS_EX_APPWINDOW)) && !IsInvisibleWin10BackgroundAppWindow(windowID))
			{
				AltTabTotalNum ++
				AltTabListID[AltTabTotalNum] := windowID			
			}
		}
	}
	DetectHiddenWindows, %HiddenWindows%
	
	; now find next window not minimized by heartbeat
	for index, element in AltTabListID {
		hwnd:= AltTabListID[A_Index]
		if (WindowList[hwnd]!=1 and WindowList[hwnd]!=2) {
			break
		}
	}
	if (WindowList[hwnd]==1 or WindowList[hwnd]==2) {
		; If nothing else is found, use desktop
		hwnd:= WinExist("Program Manager ahk_class Progman")
	}
	
	/*
	WinGet, pid, PID, ahk_id %hwnd%
	WinGetTitle, title, ahk_id %hwnd%
	wl:= WindowList[hwnd]
	MsgBox, % "pid= " pid . "`nhwnd= " . hwnd . "`ntitle= " . title "`nwl= " wl
	*/
	return hwnd
}

Decimal_to_Hex(var) {
	SetFormat, IntegerFast, H
	var += 0 
	var .= "" 
	SetFormat, Integer, D
	return var
}

IsInvisibleWin10BackgroundAppWindow(hWindow) {
	result := 0
	VarSetCapacity(cloakedVal, A_PtrSize) ; DWMWA_CLOAKED := 14
	hr := DllCall("DwmApi\DwmGetWindowAttribute", "Ptr", hWindow, "UInt", 14, "Ptr", &cloakedVal, "UInt", A_PtrSize)
	if !hr ; returns S_OK (which is zero) on success. Otherwise, it returns an HRESULT error code
	result := NumGet(cloakedVal) ; omitting the "&" performs better
	return result ? true : false
}

/*
DWMWA_CLOAKED: If the window is cloaked, the following values explain why:
1  The window was cloaked by its owner application (DWM_CLOAKED_APP)
2  The window was cloaked by the Shell (DWM_CLOAKED_SHELL)
4  The cloak value was inherited from its owner window (DWM_CLOAKED_INHERITED)
*/

;---------------------------------------------------------------------------------
LogRoll(filename, maxSize:= 5, keep:=5) {

	; Roll files if the filename is larger than maxSize. 
	; Copy/move filename.4 to filename.5, filename.3 to filename.4, etc.
	; finally copy/move filename to filename.1
	; filename with highest index (keep value) is replaced with previous index file.
	
	if (maxSize < 1) 
		maxSize:= 1
		
	FileGetSize, size, %filename%, M
	if (size>=maxSize) {
		logDebug(A_LineNumber, "LogRoll: roll files - filename= '" filename "', size= " size " MB, maxSize= " maxSize " MB")
		while (keep>1) {
			fn1:= filename "." keep
			fn2:= filename "." (keep-1)
			FileMove, %fn2%, %fn1%, true	; move and overwrite if exist
			if ErrorLevel {
				; using explicit filename, thus errorlevel is only set if the file exist and 
				; the move is unsuccessful. 
				ErrorMessage:= "Cannot copy/move '" fn2 "' to '" fn1 "', lastError= " A_LastError
				logError(A_LineNumber, "LogRoll: " ErrorMessage)
			}
			keep:= keep-1
		}
		FileMove, %filename%, %fn2%, true
		if ErrorLevel {
			; using explicit filename, thus errorlevel is only set if the file exist and 
			; the move is unsuccessful. 
			ErrorMessage:= "Cannot copy/move '" filename "' to '" fn2 "', lastError= " A_LastError
			logError(A_LineNumber, "LogRoll: " ErrorMessage)
		}
	}
}

;---------------------------------------------------------------------------------
log(level, line, msg) {
	if (level <= gLogLevel) { 
		if (level == LOG_ERROR)
			levelTxt:= "ERR"
		if (level == LOG_WARNING)
			levelTxt:= "WRN"
		if (level == LOG_INFO)
			levelTxt:= "INF"
		if (level == LOG_DEBUG)
			levelTxt:= "DBG"
		if (level == LOG_TRACE)
			levelTxt:= "TRC"
		
		FormatTime, TimeString, , yy/MM/dd HH:mm:ss
		txt:= TimeString " [" gSelfPidHex "] " levelTxt " " msg " [" line "]`n"
		h:= FileOpen(gLogFile,"a")
		h.write(txt)
		h.close()
	}
}

logTrace(line, msg) {
	log(LOG_TRACE,line,msg)
}

logDebug(line, msg) {
	log(LOG_DEBUG,line,msg)
}

logInfo(line, msg) {
	log(LOG_INFO,line,msg)
}

logWarning(line, msg) {
	log(LOG_WARNING,line,msg)
}

logError(line, msg) {
	log(LOG_ERROR,line,msg)
}

;--- end of script ---

;--- end of script ---
