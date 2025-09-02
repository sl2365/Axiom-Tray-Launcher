; Tab_Apps - ListView.au3
; This code is now redundant as it was replaced with a treeview version in Tab_Apps.au3

#include-once

Global $g_CategoryIniDir = @ScriptDir & "\App"
Global $g_AppListView, $g_AppFields, $g_IniFiles, $g_AppSections, $g_SelectedIni, $g_SelectedSection
Global $g_SaveBtn, $g_DeleteBtn
Global $g_AppLastSelected = ""

Func Tab_Apps_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("Apps")
    $g_AppListView = GUICtrlCreateListView("#|App", $listviewX, $listviewY, 390, $guiH-155)
    GUICtrlSendMsg($g_AppListView, $LVM_SETCOLUMNWIDTH, 0, 50)
    GUICtrlSendMsg($g_AppListView, $LVM_SETCOLUMNWIDTH, 1, 300)

    ; Right side fields container
    Local $x = 440
    $g_AppFields = ObjCreate("Scripting.Dictionary")
    Local $fields = ["ButtonText", "RunFile", "RunAsAdmin", "WorkDir", "Arguments", "SingleInstance", "Sandboxie", "SandboxName", "Category", "SymLinkCreate", "SymLink1", "Fave", "Hide"]
    For $i = 0 To UBound($fields) - 1
        GUICtrlCreateLabel($fields[$i] & ":", $x, 60 + $i * 27, 70, 18)
        ; Checkbox for 1/0 fields, input otherwise
        If StringInStr("RunAsAdmin SingleInstance Sandboxie SymLinkCreate Fave Hide", $fields[$i]) Then
            $g_AppFields($fields[$i]) = GUICtrlCreateCheckbox("", $x + 90, 56 + $i * 27, 18, 18)
        Else
            $g_AppFields($fields[$i]) = GUICtrlCreateInput("", $x + 90, 56 + $i * 27, 220, 18)
        EndIf
    Next

    $g_DeleteBtn = GUICtrlCreateButton("Delete", $x + 180, 400, 60, 24)
    $g_SaveBtn = GUICtrlCreateButton("Save", $x + 250, 400, 60, 24)

    _AppTab_LoadIniFilesAndApps()
EndFunc

Func _AppTab_LoadIniFilesAndApps()
    Global $g_AppSections
    $g_AppSections = ObjCreate("Scripting.Dictionary")
    $g_IniFiles = _FileListToArray($g_CategoryIniDir, "*.ini", 1)
    If @error Or Not IsArray($g_IniFiles) Then Return
    Local $idx = 1
    For $f = 1 To $g_IniFiles[0]
        Local $iniFile = $g_CategoryIniDir & "\" & $g_IniFiles[$f]
        Local $sections = IniReadSectionNames($iniFile)
        If Not IsArray($sections) Then ContinueLoop
        For $i = 1 To $sections[0]
            GUICtrlCreateListViewItem($idx & "|" & $sections[$i], $g_AppListView)
            $g_AppSections.Add($sections[$i], $iniFile)
            $idx += 1
        Next
    Next
EndFunc

Global $g_AppLastSelectedIdx = -1

Func Tab_App_HandleEvents($msg)
    ; Handle buttons
    Switch $msg
        Case $g_SaveBtn
            If $g_SelectedIni <> "" And $g_SelectedSection <> "" Then
                _AppTab_SaveFields($g_SelectedIni, $g_SelectedSection)
            EndIf
        Case $g_DeleteBtn
            _AppTab_DeleteSelected()
    EndSwitch

    ; Poll ListView for selection change (works for single-click)
    Local $itemCount = _GUICtrlListView_GetItemCount($g_AppListView)
    Local $selIdx = -1
    For $i = 0 To $itemCount - 1
        If _GUICtrlListView_GetItemSelected($g_AppListView, $i) Then
            $selIdx = $i
            ExitLoop
        EndIf
    Next

    If $selIdx <> $g_AppLastSelectedIdx Then
        $g_AppLastSelectedIdx = $selIdx

        If $selIdx <> -1 Then
            Local $appName = _GUICtrlListView_GetItemText($g_AppListView, $selIdx, 1)
            $g_SelectedSection = $appName
            If $g_AppSections.Exists($appName) Then
                $g_SelectedIni = $g_AppSections.Item($appName)
                _AppTab_PopulateFields($g_SelectedIni, $g_SelectedSection)
                _AppTab_EnableFields()
            EndIf
        Else
            _AppTab_ClearFields()
            $g_SelectedIni = ""
            $g_SelectedSection = ""
        EndIf
    EndIf
EndFunc

Func _AppTab_PopulateFields($iniFile, $section)
    For $key In $g_AppFields.Keys()
        Local $val = IniRead($iniFile, $section, $key, "")
        If StringInStr("RunAsAdmin SingleInstance Sandboxie SymLinkCreate Fave Hide", $key) Then
            GUICtrlSetState($g_AppFields($key), ($val = "1") ? $GUI_CHECKED : $GUI_UNCHECKED)
        Else
            GUICtrlSetData($g_AppFields($key), $val)
        EndIf
        ; Enable fields for editing
        GUICtrlSetState($g_AppFields($key), $GUI_ENABLE)
    Next
EndFunc

Func _AppTab_SaveFields($iniFile, $section)
    For $key In $g_AppFields.Keys()
        If StringInStr("RunAsAdmin SingleInstance Sandboxie SymLinkCreate Fave Hide", $key) Then
            Local $v = (GUICtrlRead($g_AppFields($key)) = $GUI_CHECKED) ? "1" : "0"
            IniWrite($iniFile, $section, $key, $v)
        Else
            IniWrite($iniFile, $section, $key, GUICtrlRead($g_AppFields($key)))
        EndIf
        ; Disable fields after Save
        GUICtrlSetState($g_AppFields($key), $GUI_DISABLE)
    Next
EndFunc

Func _AppTab_ClearFields()
    For $key In $g_AppFields.Keys()
        If StringInStr("RunAsAdmin SingleInstance Sandboxie SymLinkCreate Fave Hide", $key) Then
            GUICtrlSetState($g_AppFields($key), $GUI_UNCHECKED)
        Else
            GUICtrlSetData($g_AppFields($key), "")
        EndIf
        GUICtrlSetState($g_AppFields($key), $GUI_ENABLE)
    Next
EndFunc

Func _AppTab_EnableFields()
    For $key In $g_AppFields.Keys()
        GUICtrlSetState($g_AppFields($key), $GUI_ENABLE)
    Next
EndFunc

Func _AppTab_DeleteSelected()
    Local $itemCount = _GUICtrlListView_GetItemCount($g_AppListView)
    Local $selIdx = -1
    For $i = 0 To $itemCount - 1
        If _GUICtrlListView_GetItemSelected($g_AppListView, $i) Then
            $selIdx = $i
            ExitLoop
        EndIf
    Next
    If $selIdx = -1 Then Return

    Local $appName = _GUICtrlListView_GetItemText($g_AppListView, $selIdx, 1)
    ; Confirmation dialog
    Local $answer = MsgBox(1 + 32, "Delete App", "Are you sure you want to delete '" & $appName & "'?" & @CRLF & "This cannot be undone.", 0)
    If $answer = 1 Then ; OK pressed
        ; Remove section from INI file
        If $g_AppSections.Exists($appName) Then
            Local $iniFile = $g_AppSections.Item($appName)
            IniDelete($iniFile, $appName)
            ; Remove from ListView
            _GUICtrlListView_DeleteItem($g_AppListView, $selIdx)
            ; Remove from dictionary
            $g_AppSections.Remove($appName)
        EndIf
        _AppTab_ClearFields()
        $g_SelectedIni = ""
        $g_SelectedSection = ""
        $g_AppLastSelectedIdx = -1
    EndIf
EndFunc

Func _AppTab_RemoveListViewItem($appName)
    ; Find and remove item from ListView
    Local $itemCount = _GUICtrlListView_GetItemCount($g_AppListView)
    For $i = 0 To $itemCount - 1
        Local $text = _GUICtrlListView_GetItemText($g_AppListView, $i, 1)
        If $text = $appName Then
            _GUICtrlListView_DeleteItem($g_AppListView, $i)
            ExitLoop
        EndIf
    Next
EndFunc
