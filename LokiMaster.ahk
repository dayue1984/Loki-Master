; ///////////////////////////////////////////
; An enhencement for Logi Master 3 like mouse
; Works with Logi Master 3 / Filco Minila air
; AutoHotkey 1.1.33.10 / Win11 21H2 22000.348
; ///////////////////////////////////////////

; ///////
; Compile
; ///////
;@Ahk2Exe-SetMainIcon %A_ScriptDir%/res/1/1.ico

; ////////////
; Environments
; ////////////
#NoEnv
;#Warn
#Persistent
#SingleInstance force
SendMode Input
SetWorkingDir %A_ScriptDir%

; ///////
; Globals
; ///////
global app_Name = "Loki Master" ; This is Used by Autorun & Msgbox
global app_Version = "0.21.1208"
global app_ConfigFile = "Config.ini"
global app_KeyDelay = 50
global app_TimerPeriod = 500
global desktop_Count = 1 ; Will be Updated by desktop_Update()
global desktop_Current = 1 ; Current desktop
global mode_User = 1 ; User Selected Mode
global mode_OnTheFly = 0 ; The Current Actived Mode by mode_WatchDog()
global mode_NameList := ["SMART","DESKTOP","CODE","OFFICE","BROWSER","MEDIA"] ; Mode Name List for OSD/Config
global mode_AppList := [] ; App Titles for Smart Mode
global mode_OSDList := [] ; Turn ON/OFF OSD with Specified Modes

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
config_Read()
SetTimer, mode_WatchDog, %app_TimerPeriod%
SetKeyDelay, app_KeyDelay
autorun_Check()

; /////
; Mouse
; /////
; Gestrue Key: Loop the Modes
XButton1:: ; Cycle Desktops
{
    desktop_SwitchByCycle() 
}
XButton2::Send {LWin Down}d{LWin Up} ; Show Desktop
#Tab:: ; Switch Mode
If (mode_User < mode_NameList.Length()) {
    mode_User++
}
Else {
    mode_User = 1
}
mode_WatchDog()
config_Save()
; OutputDebug, [MODE] User: %mode_User%
Return

; /////
; Modes
; /////
; Desktop
#If mode_OnTheFly = 2
WheelLeft::Volume_Up
WheelRight::Volume_Down
#If
; CODE
#If mode_OnTheFly = 3
; XButton1::Home
; XButton2::End
WheelLeft::Send {LCtrl Down}{NumpadAdd}{NumpadAdd Up}{LCtrl Up} ; Zoom In
WheelRight::Send {LCtrl Down}{NumpadSub}{NumpadSub Up}{LCtrl Up} ; Zoom Out
#If
; Office
#If mode_OnTheFly = 4
; XButton1::Home
; XButton2::End
WheelLeft::Send {LCtrl Down}{WheelUp}{LCtrl Up} ; Zoom In
WheelRight::Send {LCtrl Down}{WheelDown}{LCtrl Up} ; Zoom Out
#If
; Browser
#If mode_OnTheFly = 5
; XButton1::Browser_Forward
; XButton2::Browser_Back
; MButton::Browser_Refresh
WheelLeft::Send {LCtrl Down}{NumpadAdd}{NumpadAdd Up}{LCtrl Up} ; Zoom In
WheelRight::Send {LCtrl Down}{NumpadSub}{NumpadSub Up}{LCtrl Up} ; Zoom Out
#If
; Media
#If mode_OnTheFly = 6
LButton::Media_Prev
RButton::Media_Next
MButton::Media_Play_Pause
WheelLeft::Volume_Up
WheelRight::Volume_Down
#If

; ////////
; Keyboard
; ////////
; Capslock OSD
CapsLock:: invert_CapsLock()
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
; Watchdog for Desktop & Mode
mode_WatchDog() {
    static lastUserMode := 0, lastDesktop = 0
    ; Virtual Desktop
    desktop_Update()
    if (lastDesktop <> desktop_Current) {
        lastDesktop := desktop_Current
        themeID := get_Theme()
        Menu, Tray, Icon, %A_ScriptDir%\res\%themeID%\%desktop_Current%.ico
        ; OutputDebug, [DESKTOP] OnTheFly: %mode_OnTheFly%
    }
    ; Switch Mode
    if (lastUserMode <> mode_User) {
        lastUserMode := mode_User
        mode_OnTheFly := mode_User
        OSD(mode_NameList[(mode_User)])
    }
    ; Smart Mode
    if (mode_User = 1) {
        dogCatched := False
        winTitle = ""
        WinGetActiveTitle, winTitle
        if StrLen(winTitle) {
            Loop % mode_NameList.Length() {
                if winTitle Contains % mode_AppList[(A_Index)]
                {
                    ; New Mode Catched?
                    if (mode_OnTheFly <> A_Index) {
                        mode_OnTheFly := A_Index
                        ; Check [OSD] Section in Config
                        if (mode_OSDList[(mode_OnTheFly)]) {
                            OSD(mode_NameList[(mode_OnTheFly)], True)
                        }
                        ; OutputDebug, [MODE] OnTheFly: %mode_OnTheFly% 
                    }
                    dogCatched := True
                    Break
                }
            }
        }
        ; Fallback to Desktop Mode
        if !dogCatched {
            mode_OnTheFly := 2
        }
    }
    Return
}
; Get System Theme from Registry: Light = 1 / Dark = 0
get_Theme() {
    RegRead, useLightTheme, HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize, SystemUsesLightTheme
    if ErrorLevel {
        Return 0 ; Default to dark
    }
    Return useLightTheme
}
; Invert the CapsLock
invert_CapsLock() {
    static flag_CpasLock := False
    if flag_CpasLock {
        SetCapsLockState, off
        flag_CpasLock := False
        OSD("Caps OFF")
    }
    else {
        SetCapsLockState, on
        flag_CpasLock := True
        OSD("Caps ON")
    }
    Return
}
; On Screen Display
OSD(TXT, Grayed:=False)
{
    backColor = EEAA99
    Gui, OSD: +AlwaysOnTop +LastFound +Owner 
    Gui, OSD: Color, %backColor% 
    Gui, OSD: Font, s20, Verdana
    if Grayed {
        Gui, OSD: Add, Text, cSilver, %TXT%
    }
    Else {
        Gui, OSD: Add, Text, cLime, %TXT%
    }
    WinSet, TransColor, %backColor% 200 ; Transparent
    Gui, OSD: -Caption
    Gui, OSD: Show, center center
    Sleep, 500
    Gui, OSD: destroy
    return
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
; Save Configs
config_Save() {
    IniWrite, %mode_User%, %app_ConfigFile%, APP, MODE_USER
    Return
}
; Read Configs
config_Read() {
    IniRead, mode_User, %app_ConfigFile%, APP, MODE_USER, 1
    IniRead, app_KeyDelay, %app_ConfigFile%, APP, KEY_DELAY, 50
    IniRead, app_TimerPeriod, %app_ConfigFile%, APP, TIMER_PERIOD, 500
    Loop % mode_NameList.Length() {
        strApp := ""
        strOSD := False
        strKey := mode_NameList[(A_Index)]
        IniRead, strApp, %app_ConfigFile%, MODE, %strKey%, ""
        IniRead, strOSD, %app_ConfigFile%, SMART_OSD, %strKey%, False
        mode_AppList[(A_Index)] := strApp
        mode_OSDList[(A_Index)] := strOSD
    }
    Return
}
