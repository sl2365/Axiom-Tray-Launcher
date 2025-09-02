; EnvVarsListView.au3
; Displays environment variables in a ListView with descriptions, allows copying variable name
; When a variable is copied, shows a green "Copied!" label for 1 second.

#include-once
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <Array.au3>
#include <Clipboard.au3>
#include <WinAPI.au3>
#include <Timers.au3>

Global $g_EnvVarsListView, $g_EnvVarsLabelCopied, $g_EnvVarsLabelTimer = -1
Global $g_EnvVarsFooterMsg
Global $g_EnvVarsLabelCopied
Global $g_EnvVarsLabelTimer = 0
Global $g_EnvVarsLabelTimeout = 3000 ; milliseconds

Global $g_EnvVars[43][2]
$g_EnvVars[0][0] = "%ALLUSERSPROFILE%"
$g_EnvVars[0][1] = "C:\ProgramData"
$g_EnvVars[1][0] = "%APPDATA%"
$g_EnvVars[1][1] = "C:\Users\{username}\AppData\Roaming"
$g_EnvVars[2][0] = "%COMMONPROGRAMFILES%"
$g_EnvVars[2][1] = "C:\Program Files\Common Files"
$g_EnvVars[3][0] = "%COMMONPROGRAMFILES(x86)%"
$g_EnvVars[3][1] = "C:\Program Files (x86)\Common Files"
$g_EnvVars[4][0] = "%CommonProgramW6432%"
$g_EnvVars[4][1] = "C:\Program Files\Common Files"
$g_EnvVars[5][0] = "%COMSPEC%"
$g_EnvVars[5][1] = "C:\Windows\System32\cmd.exe"
$g_EnvVars[6][0] = "%HOMEDRIVE%"
$g_EnvVars[6][1] = "C:\"
$g_EnvVars[7][0] = "%HOMEPATH%"
$g_EnvVars[7][1] = "C:\Users\{username}"
$g_EnvVars[8][0] = "%LOCALAPPDATA%"
$g_EnvVars[8][1] = "C:\Users\{username}\AppData\Local"
$g_EnvVars[9][0] = "%LOGONSERVER%"
$g_EnvVars[9][1] = "\{domain_logon_server}"
$g_EnvVars[10][0] = "%PATH%"
$g_EnvVars[10][1] = "C:\Windows\system32 ; C:\Windows ; C:\Windows\System32\Wbem"
$g_EnvVars[11][0] = "%PathExt%"
$g_EnvVars[11][1] = ".com;.exe;.bat;.cmd;.vbs;.vbe;.js;.jse;.wsf;.wsh;.msc"
$g_EnvVars[12][0] = "%PROGRAMDATA%"
$g_EnvVars[12][1] = "C:\ProgramData"
$g_EnvVars[13][0] = "%PROGRAMFILES%"
$g_EnvVars[13][1] = "C:\Program Files"
$g_EnvVars[14][0] = "%ProgramW6432%"
$g_EnvVars[14][1] = "C:\Program Files"
$g_EnvVars[15][0] = "%PROGRAMFILES(X86)%"
$g_EnvVars[15][1] = "C:\Program Files (x86)"
$g_EnvVars[16][0] = "%PROMPT%"
$g_EnvVars[16][1] = "$P$G"
$g_EnvVars[17][0] = "%SystemDrive%"
$g_EnvVars[17][1] = "C:"
$g_EnvVars[18][0] = "%SystemRoot%"
$g_EnvVars[18][1] = "C:\Windows"
$g_EnvVars[19][0] = "%TEMP%"
$g_EnvVars[19][1] = "C:\Users\{username}\AppData\Local\Temp"
$g_EnvVars[20][0] = "%TMP%"
$g_EnvVars[20][1] = "C:\Users\{username}\AppData\Local\Temp"
$g_EnvVars[21][0] = "%USERDOMAIN%"
$g_EnvVars[21][1] = "Userdomain associated with current user."
$g_EnvVars[22][0] = "%USERDOMAIN_ROAMINGPROFILE%"
$g_EnvVars[22][1] = "Userdomain associated with roaming profile."
$g_EnvVars[23][0] = "%USERNAME%"
$g_EnvVars[23][1] = "{username}"
$g_EnvVars[24][0] = "%USERPROFILE%"
$g_EnvVars[24][1] = "C:\Users\{username}"
$g_EnvVars[25][0] = "%WINDIR%"
$g_EnvVars[25][1] = "C:\Windows"
$g_EnvVars[26][0] = "%PUBLIC%"
$g_EnvVars[26][1] = "C:\Users\Public"
$g_EnvVars[27][0] = "%PSModulePath%"
$g_EnvVars[27][1] = "%SystemRoot%\system32\WindowsPowerShell\v1.0\Modules\"
$g_EnvVars[28][0] = "%OneDrive%"
$g_EnvVars[28][1] = "C:\Users\{username}\OneDrive"
$g_EnvVars[29][0] = "%DriverData%"
$g_EnvVars[29][1] = "C:\Windows\System32\Drivers\DriverData"
$g_EnvVars[30][0] = "%CD%"
$g_EnvVars[30][1] = "Outputs current directory path."
$g_EnvVars[31][0] = "%CMDCMDLINE%"
$g_EnvVars[31][1] = "Outputs command line used to launch current Command Prompt session."
$g_EnvVars[32][0] = "%CMDEXTVERSION%"
$g_EnvVars[32][1] = "Outputs the number of current command processor extensions."
$g_EnvVars[33][0] = "%COMPUTERNAME%"
$g_EnvVars[33][1] = "Outputs the system name."
$g_EnvVars[34][0] = "%DATE%"
$g_EnvVars[34][1] = "Outputs current date."
$g_EnvVars[35][0] = "%TIME%"
$g_EnvVars[35][1] = "Outputs time."
$g_EnvVars[36][0] = "%ERRORLEVEL%"
$g_EnvVars[36][1] = "Outputs the exit status of previous command."
$g_EnvVars[37][0] = "%PROCESSOR_IDENTIFIER%"
$g_EnvVars[37][1] = "Outputs processor identifier."
$g_EnvVars[38][0] = "%PROCESSOR_LEVEL%"
$g_EnvVars[38][1] = "Outputs processor level."
$g_EnvVars[39][0] = "%PROCESSOR_REVISION%"
$g_EnvVars[39][1] = "Outputs processor revision."
$g_EnvVars[40][0] = "%NUMBER_OF_PROCESSORS%"
$g_EnvVars[40][1] = "Outputs the number of physical and virtual cores."
$g_EnvVars[41][0] = "%RANDOM%"
$g_EnvVars[41][1] = "Outputs random number from 0 through 32767."
$g_EnvVars[42][0] = "%OS%"
$g_EnvVars[42][1] = "Windows_NT"

Func Tab_EnvVars_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("Env Vars")
    Local $rows = UBound($g_EnvVars)
    Local $w = $guiW -20, $h = 550 ; match tab size
    ; ListView
    $g_EnvVarsListView = GUICtrlCreateListView("Variable|Description", $listviewX, $listviewY, $w-20, $guiH-155, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL)
    GUICtrlSetFont($g_EnvVarsListView, 10, 400, 0, "Consolas")
    GUICtrlSetBkColor($g_EnvVarsListView, 0xF8F8F8)
    GUICtrlSetColor($g_EnvVarsListView, 0x222222)
    GUICtrlSetResizing($g_EnvVarsListView, $GUI_DOCKALL)
    ; Add all environment variable rows
    For $i = 0 To $rows - 1
        GUICtrlCreateListViewItem($g_EnvVars[$i][0] & "|" & $g_EnvVars[$i][1], $g_EnvVarsListView)
    Next

    _GUICtrlListView_SetColumnWidth($g_EnvVarsListView, 0, 200)
    _GUICtrlListView_SetColumnWidth($g_EnvVarsListView, 1, 500)
    ; Footer message
    $g_EnvVarsFooterMsg = GUICtrlCreateLabel("Double-click to copy variable to clipboard.", 20, $guiH-100, 350, 30)
    GUICtrlSetFont($g_EnvVarsFooterMsg, 10, 400, 0, "Segoe UI")
    GUICtrlSetColor($g_EnvVarsFooterMsg, 0x666666)
    GUICtrlSetBkColor($g_EnvVarsFooterMsg, $GUI_BKCOLOR_TRANSPARENT)
    GUICtrlSetResizing($g_EnvVarsFooterMsg, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM)
    ; "Copied!" label
    $g_EnvVarsLabelCopied = GUICtrlCreateLabel("", 20, $h - $footer_gap + 25, 350, 25)
    GUICtrlSetFont($g_EnvVarsLabelCopied, 10, 700, 0, "Segoe UI")
    GUICtrlSetColor($g_EnvVarsLabelCopied, 0x008000) ; Green
    GUICtrlSetBkColor($g_EnvVarsLabelCopied, $GUI_BKCOLOR_TRANSPARENT)
    GUICtrlSetState($g_EnvVarsLabelCopied, $GUI_HIDE)
    ; Register double-click notification for listview
    GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY_ENVVARS_TAB")

    Return 1
EndFunc

Func WM_NOTIFY_ENVVARS_TAB($hWnd, $Msg, $wParam, $lParam)
    Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
    Local $code = DllStructGetData($tNMHDR, "Code")
    If $wParam = $g_EnvVarsListView And $code = -3 Then
        Local $tNMITEMACTIVATE = DllStructCreate($tagNMITEMACTIVATE, $lParam)
        Local $iItem = DllStructGetData($tNMITEMACTIVATE, "Index")
        If $iItem <> -1 And $iItem >= 0 And $iItem < UBound($g_EnvVars) Then
            ClipPut($g_EnvVars[$iItem][0])
            GUICtrlSetData($g_EnvVarsLabelCopied, "Copied: " & $g_EnvVars[$iItem][0])
            GUICtrlSetState($g_EnvVarsLabelCopied, $GUI_SHOW)
            $g_EnvVarsLabelTimer = TimerInit()
        EndIf
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc
