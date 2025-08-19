#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <Array.au3>
#include <Clipboard.au3>
#include <WinAPI.au3>

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
$g_EnvVars[9][1] = "\\{domain_logon_server}"
$g_EnvVars[10][0] = "%PATH%"
$g_EnvVars[10][1] = "C:\Windows\system32;C:\Windows;C:\Windows\System32\Wbem"
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

Global $listview

Func Main()
    Local $w = 600, $h = 200
    Local $mainGUI = GUICreate("My Tools Menu", $w, $h)
    Local $menuBar = GUICtrlCreateMenu("Tools")
    Local $envMenu = GUICtrlCreateMenuItem("Env Vars", $menuBar)
    Local $closeBtn = GUICtrlCreateButton("Exit", $w - 100, $h - 50, 80, 30)
    GUISetState(@SW_SHOW, $mainGUI)
    While 1
        Local $msg = GUIGetMsg()
        If $msg = $GUI_EVENT_CLOSE Or $msg = $closeBtn Then
            GUIDelete($mainGUI)
            Exit
        ElseIf $msg = $envMenu Then
            ShowEnvVarListView()
        EndIf
    WEnd
EndFunc

Func ShowEnvVarListView()
    Local $rows = UBound($g_EnvVars)
    Local $w = 640, $h = 500
    Local $footer_gap = 80
    Local $button_top = $h - $footer_gap + 15
    Local $gui = GUICreate("Environment Variables Reference", $w, $h, -1, -1, $WS_EX_TOPMOST + $WS_EX_WINDOWEDGE)
    $listview = GUICtrlCreateListView("Variable|Description", 10, 10, $w-20, $h-$footer_gap, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL)
    GUICtrlSetFont($listview, 10, 400, 0, "Consolas")
    GUICtrlSetBkColor($listview, 0xF8F8F8)
    GUICtrlSetColor($listview, 0x222222)
    GUICtrlSetResizing($listview, $GUI_DOCKALL)

    _GUICtrlListView_DeleteAllItems(GUICtrlGetHandle($listview))
    For $i = 0 To $rows-1
        GUICtrlCreateListViewItem($g_EnvVars[$i][0] & "|" & $g_EnvVars[$i][1], $listview)
    Next

    _GUICtrlListView_SetColumnWidth($listview, 0, 200)
    _GUICtrlListView_SetColumnWidth($listview, 1, 400)

    Local $closeBtn = GUICtrlCreateButton("Close", $w-140, $button_top, 120, 30)
    Local $footerMsg = GUICtrlCreateLabel("Double click to copy variable.", 10, $h - $footer_gap + 15, 300, 30)
    GUICtrlSetFont($footerMsg, 10, 400, 0, "Segoe UI")
    GUICtrlSetColor($footerMsg, 0x666666)
    GUICtrlSetBkColor($footerMsg, $GUI_BKCOLOR_TRANSPARENT)
    GUICtrlSetResizing($footerMsg, $GUI_DOCKLEFT + $GUI_DOCKBOTTOM)

    GUISetState(@SW_SHOW, $gui)
    WinSetOnTop($gui, "", 1)
    GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

    While 1
        Local $msg = GUIGetMsg()
        If $msg = $GUI_EVENT_CLOSE Or $msg = $closeBtn Then
            GUIRegisterMsg($WM_NOTIFY, "")
            GUIDelete($gui)
            ExitLoop
        EndIf
    WEnd
EndFunc

Func WM_NOTIFY($hWnd, $Msg, $wParam, $lParam)
    Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
    Local $code = DllStructGetData($tNMHDR, "Code")
    If $wParam = $listview And $code = -3 Then ; NMITEMACTIVATE
        Local $tNMITEMACTIVATE = DllStructCreate($tagNMITEMACTIVATE, $lParam)
        Local $iItem = DllStructGetData($tNMITEMACTIVATE, "Index")
        If $iItem <> -1 And $iItem >= 0 And $iItem < UBound($g_EnvVars) Then
            ClipPut($g_EnvVars[$iItem][0])
            TrayTip("Copied", "Copied: " & $g_EnvVars[$iItem][0], 1000)
        EndIf
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc
