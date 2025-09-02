; SettingsGUI.au3

#include-once
#include <GUIConstantsEx.au3>
#include <TabConstants.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <Array.au3>
#include "Tab_Global.au3"
#include "Tab_Apps.au3"
#include "Tab_ScanFolders.au3"
#include "Tab_Ignore.au3"
#include "Tab_EnvVars.au3"
#include "Tab_About.au3"
#include "Updates.au3"

; Global handles for window and controls (for main loop access)
Global $hSettingsGUI = 0
Global $hTab
Global $g_TabPages[5]
Global $g_AboutLinkCtrl, $g_CheckUpdatesCtrl, $g_OpenSettingsFolderCtrl
; Global handles for OK and Cancel buttons
Global $g_SettingsOKBtn, $g_SettingsCancelBtn
Global $g_OKBtn, $g_CancelBtn
Global $g_ScanFoldersBtn
Global $guiW = 700
Global $guiH = 570
Global $listviewX = 20
Global $listviewY = 50
Global $btnW = 60
Global $btnH = 25
Global $footer_gap = 80
Global $g_DarkThemeCheckbox

Func ShowSettingsGUI()
    ; If already open, bring to front
    If WinExists("Axiom Settings") Then
        WinActivate("Axiom Settings")
        Return
    EndIf

    $hSettingsGUI = GUICreate("Axiom Settings", $guiW, $guiH)
    $hTab = GUICtrlCreateTab(10, 10, $guiW-20, $guiH-50)

    ; Create tab pages using modular functions
    $g_TabPages[0] = Tab_Global_Create($hSettingsGUI, $hTab)
    $g_TabPages[1] = Tab_ScanFolders_Create($hSettingsGUI, $hTab)
    $g_TabPages[2] = Tab_Apps_Create($hSettingsGUI, $hTab)
    GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
    $g_TabPages[4] = Tab_Ignore_Create($hSettingsGUI, $hTab)
    $g_TabPages[3] = Tab_EnvVars_Create($hSettingsGUI, $hTab)
    $g_TabPages[4] = Tab_About_Create($hSettingsGUI, $hTab)
	GUICtrlCreateTabItem("")
	GUICtrlSendMsg($hTab, $TCM_SETCURSEL, 0, 0)
    ; Add OK and Cancel buttons at the bottom right
    $g_OKBtn = GUICtrlCreateButton("OK", $guiW -180, 535, 70, $btnH)
    $g_CancelBtn = GUICtrlCreateButton("Cancel", $guiW -100, 535, 70, $btnH)
	$g_ScanFoldersBtn = GUICtrlCreateButton("Scan Folders", 30, 535, 90, $btnH)
    GUISetState(@SW_SHOW, $hSettingsGUI)
EndFunc

Func SettingsGUI_HandleEvents()
    Local $msg = GUIGetMsg()
	If $msg = $g_ViewSwitch Then
        If GUICtrlRead($g_ViewSwitch) = $GUI_CHECKED Then
            $g_TreeViewMode = "folder"
        Else
            $g_TreeViewMode = "tray"
        EndIf
        _AppTab_LoadIniFilesAndApps()
    EndIf
    ; Pass messages to tab event handlers if needed
	Tab_Global_HandleEvents($msg)
	Vars_HandleEvents($msg)
	Tab_App_HandleEvents($msg)
	SymLinks_HandleEvents($msg)
    ScanFolders_HandleEvents($msg)
	Ignore_HandleEvents($msg)
	
    ; Handle OK/Cancel buttons
    If $msg = $g_OKBtn Then
        Tab_ScanFolders_Save()
        _TrayMenu_RefreshFull()
        GUIDelete($hSettingsGUI)
        $hSettingsGUI = 0
    ElseIf $msg = $g_CancelBtn Then
        _TrayMenu_RefreshFull()
        GUIDelete($hSettingsGUI)
        $hSettingsGUI = 0
    ElseIf $msg = $GUI_EVENT_CLOSE Then
        _TrayMenu_RefreshFull()
        GUIDelete($hSettingsGUI)
        $hSettingsGUI = 0
    ElseIf $msg = $g_ScanFoldersBtn Then
        _ScanFolders_AndShowResults($g_ScanItems)
    ; Handle About tab buttons
    ElseIf $msg = $g_AboutLinkCtrl Then
        ShellExecute("https://github.com/sl2365/Axiom-Tray-Launcher")
    ElseIf $msg = $g_CheckUpdatesCtrl Then
        Updates_Check(True)
    ElseIf $msg = $g_OpenSettingsFolderCtrl Then
        ShellExecute(@ScriptDir & "\App")
    EndIf
    If $g_EnvVarsLabelTimer <> 0 Then
        If TimerDiff($g_EnvVarsLabelTimer) >= $g_EnvVarsLabelTimeout Then
            GUICtrlSetData($g_EnvVarsLabelCopied, "")
            GUICtrlSetState($g_EnvVarsLabelCopied, $GUI_HIDE)
            $g_EnvVarsLabelTimer = 0
        EndIf
    EndIf
EndFunc

Func ApplyTheme($themeValue)
    If $themeValue = "1" Then
        GUISetBkColor(0x232323, $hSettingsGUI) ; Main window dark
        ; Optionally: update label/input/button colors here
    Else
        GUISetBkColor(0xFFFFFF, $hSettingsGUI) ; Main window light
    EndIf
EndFunc
