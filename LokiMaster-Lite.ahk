; ///////////////////////////////////////////
; An enhencement for Logi Master 3 like mouse
; The lite version without mods, osd & config
; AutoHotkey 1.1.33.10 / Win11 21H2 22000.348
; ///////////////////////////////////////////

; ///////
; Compile
; ///////
;@Ahk2Exe-SetMainIcon %A_ScriptDir%/res/1/1.ico

; /////////
; Privilege
; /////////
full_command_line := DllCall("GetCommandLine", "str") 
if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)")) 
{ 
    try ; leads to having the script re-launching itself as administrator 
    { 
     if A_IsCompiled 
      Run *RunAs "%A_ScriptFullPath%" /restart 
     else 
      Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%" 
    } 
    ExitApp 
} 

; ////////////
; Environments
; ////////////
#NoEnv
;#Warn
#Persistent
#SingleInstance force
#UseHook
SendMode Input
SetWorkingDir %A_ScriptDir%

; ///////
; Globals
; ///////
global app_Name = "Loki Master Lite" ; This is Used by Autorun & Msgbox
global app_Version = "0.21.1228"
global desktop_Count = 1 ; Will be Updated by desktop_Update()
global desktop_Current = 1 ; Current desktop

; /////////
; Tray Menu
; /////////
Menu, Tray, NoStandard
Menu, Tray, Tip, %app_Name% %app_Version%
Menu, Tray, Add , Autorun, menuAutorun
Menu, Tray, Add , Suspend, menuSuspend
Menu, Tray, Add , Reload, menuReload
Menu, Tray, Add , Exit, menuExit
Menu, Tray, Default, Exit

; ////
; Main
; ////
SetKeyDelay, 75
autorun_Check()
SetTimer, mode_WatchDog, 600

; /////
; Mouse
; /////
XButton1:: ; Cycle Desktops
{
    desktop_SwitchByCycle() 
}
XButton2::Send #d
WheelLeft::Volume_Up
WheelRight::Volume_Down
; OutputDebug, [MODE] User: %mode_User%

; ////////
; Keyboard
; ////////
; Desktop Control
+#1:: desktop_SwitchByNumber(1)
+#2:: desktop_SwitchByNumber(2)
+#3:: desktop_SwitchByNumber(3)
+#4:: desktop_SwitchByNumber(4)
+#5:: desktop_SwitchByNumber(5)
+#6:: desktop_SwitchByNumber(6)
+#7:: desktop_SwitchByNumber(7)
+#8:: desktop_SwitchByNumber(8)
+#9:: desktop_SwitchByNumber(9)
+#NumpadAdd:: desktop_Create()
+#NumpadSub:: desktop_Delete()
; Media Control
AppsKey & Up::Volume_Up
AppsKey & Down::Volume_Down
AppsKey & Left::Media_Prev
AppsKey & Right::Media_Next
AppsKey & Enter::Media_Play_Pause
; Swtich Appskey & RAlt
AppsKey::Ralt 
RAlt::AppsKey

; /////////
; Menu Subs
; /////////
; Autorun
menuAutorun:
if autorun_Check() {
    autorun_Disable() ; Check/Uncheck this Menu in autorun_Check()
}
Else {
    autorun_Enable()
}
Return
; Suspend
menuSuspend:
if A_IsSuspended {
    Suspend OFF
    Menu, Tray, UnCheck, Suspend
}
Else {
    Suspend On
    Menu, Tray, Check, Suspend
}
Return
; Reload
menuReload:
Reload
Return
; Exit
menuExit:
MsgBox, 36, %app_Name%, Are you sure to Exit?
IfMsgBox Yes
ExitApp
Return

; /////////////
; Sub Functions
; /////////////
; Watchdog for Desktop
mode_WatchDog() {
    static lastDesktop = 0
    ; Virtual Desktop
    desktop_Update()
    if (lastDesktop <> desktop_Current) {
        lastDesktop := desktop_Current
        RegRead, themeID, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize, SystemUsesLightTheme
        if ErrorLevel {
            themeID = 0 ; Default to dark
        }
        Menu, Tray, Icon, %A_ScriptDir%\res\%themeID%\%desktop_Current%.ico
        ; OutputDebug, [DESKTOP] OnTheFly: %mode_OnTheFly%
    }
    Return
}
; Update Virtual Desktops from Registry
desktop_Update() {
    ; Get the UUID of Current Desktop
    IdLength = 32
    RegRead, desktop_CurrentId, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops, CurrentVirtualDesktop
    if ErrorLevel {
        OutputDebug, [DESKTOP] CurrentVirtualDesktop Failed
        Return
    }
    if (desktop_CurrentId) {
        IdLength := StrLen(desktop_CurrentId)
    }
    RegRead, DesktopList, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops, VirtualDesktopIDs
    if ErrorLevel {
        OutputDebug, [DESKTOP] VirtualDesktopIDs Failed
        Return
    }
    ; Count Desktops
    if DesktopList {
        DesktopListLength := StrLen(DesktopList)
        desktop_Count := Floor(DesktopListLength / IdLength) 
    }
    else {
        desktop_Count = 1
    }
    ; Find the Current Desktop
    while (desktop_CurrentId and A_Index <= desktop_Count) {
        StartPos := ((A_Index-1) * IdLength) + 1
        DesktopIter := SubStr(DesktopList, StartPos, IdLength)
        if (DesktopIter = desktop_CurrentId) {
            desktop_Current := A_Index
            Break
        }
    }
}
; Cycle Desktops
desktop_SwitchByCycle() {
    ; Check Update before Switch
    desktop_Update()
    if (desktop_Current < desktop_Count)
    {
        Send ^#{Right}
        desktop_Current++
    }
    else 
    {
        While(desktop_Current>1)
        {
            Send ^#{Left}
            desktop_Current--
        } 
    }
    ; OutputDebug, [DESKTOP] Current: %desktop_Current%
}
; Swtich Desktops by Number
desktop_SwitchByNumber(desktop_Target)
{
    desktop_Update()
    ; Check Invalid Numbers
    if (desktop_Count < desktop_Target) {
        desktop_Target := desktop_Count
    }
    ; Swith to Target Desktop by Sending Keys
    while (desktop_Current < desktop_Target) {
        Send ^#{Right}
        desktop_Current++
    }
    while (desktop_Current > desktop_Target) {
        Send ^#{Left}
        desktop_Current--
    }
    ; OutputDebug, [DESKTOP] Current: %desktop_Current%
}
; Create & Switch to a New Virtual Desktop
desktop_Create()
{
    Send, #^d
    desktop_Count++
    desktop_Current = %desktop_Count%
    ; OutputDebug, [DESKTOP] Created, Count: %desktop_Count%
}
; Delete the Current Virtual Desktop
desktop_Delete()
{
    Send, #^{F4}
    desktop_Count--
    desktop_Current--
    ; OutputDebug, [DESKTOP] Deleted, Count: %desktop_Count%
}
; Enable Autorun
autorun_Enable() {
    RegWrite, REG_SZ, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Run, %app_Name%, %A_ScriptFullPath%
    if autorun_Check() {
        Return True
    }
    Return False
}
; Disable Autorun
autorun_Disable() {
    RegDelete, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run, %app_Name%
    if autorun_Check() {
        Return False
    }
    Return True
}
; Check Autorun Status
autorun_Check() {
    RegRead, testPath, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Run, %app_Name%
    if (testPath = A_ScriptFullPath) {
        Menu, Tray, Check, Autorun
        Return True
    }
    Menu, Tray, UnCheck, Autorun
    Return False
}
