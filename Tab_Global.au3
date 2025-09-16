; Tab_Global.au3

#include-once

Global $g_SettingsIni = @ScriptDir & "\App\Settings.ini"
Global $g_SandboxiePathEdit, $g_SandboxiePathBrowseBtn
Global $guiW, $guiH, $listviewX, $listviewY, $btnW, $btnH, $footer_gap

;----------------- Variables Section Globals -----------------
Global $g_VarsListView, $g_VarsAddBtn, $g_VarsSaveBtn
Global $g_VarsItems[0][2], $g_VarsSelectedIdx = -1, $g_VarsLastSelectedIdx = -1
Global $g_VarsListViewMenu, $g_VarsDeleteMenuItem

;----------------- SymLinks Section Globals -----------------
Global $g_SymLinksListView, $g_SymLinksAddBtn, $g_SymLinksSaveBtn
Global $g_SymLinksItems[0][2], $g_SymLinksSelectedIdx = -1, $g_SymLinksLastSelectedIdx = -1
Global $g_SymLinksListViewMenu, $g_SymLinksDeleteMenuItem
Global $g_SymLinksAddCheckbox, $g_SymLinksRemoveCheckbox, $g_CreateSymlinksBtn, $g_RemoveSymlinksBtn

;----------------- Unified Edit Field -----------------
Global $g_CombinedEdit

Func _Vars_Buttons_DisableAll()
    GUICtrlSetState($g_VarsAddBtn, $GUI_DISABLE)
    GUICtrlSetState($g_VarsSaveBtn, $GUI_DISABLE)
EndFunc

Func _Vars_Buttons_EnableAdd()
    GUICtrlSetState($g_VarsAddBtn, $GUI_ENABLE)
    GUICtrlSetState($g_VarsSaveBtn, $GUI_ENABLE)
EndFunc

Func _Vars_Buttons_EnableSaveOnly()
    GUICtrlSetState($g_VarsAddBtn, $GUI_DISABLE)
    GUICtrlSetState($g_VarsSaveBtn, $GUI_ENABLE)
EndFunc

Func _SymLinks_Buttons_DisableAll()
    GUICtrlSetState($g_SymLinksAddBtn, $GUI_DISABLE)
    GUICtrlSetState($g_SymLinksSaveBtn, $GUI_DISABLE)
EndFunc

Func _SymLinks_Buttons_EnableAdd()
    GUICtrlSetState($g_SymLinksAddBtn, $GUI_ENABLE)
    GUICtrlSetState($g_SymLinksSaveBtn, $GUI_ENABLE)
EndFunc

Func _SymLinks_Buttons_EnableSaveOnly()
    GUICtrlSetState($g_SymLinksAddBtn, $GUI_DISABLE)
    GUICtrlSetState($g_SymLinksSaveBtn, $GUI_ENABLE)
EndFunc

Func Tab_Global_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("Global")
	
    ; ---- Global Options ----
    $UPDATE_ON_START_KEY = GUICtrlCreateCheckbox("Check for Updates on Startup", 20, 435, 200, 18)
    Local $val = IniRead(@ScriptDir & "\App\Settings.ini", "GLOBAL", "UpdateOnStart", "1")
    GUICtrlSetState($UPDATE_ON_START_KEY, ($val = "1") ? $GUI_CHECKED : $GUI_UNCHECKED)

    GUICtrlCreateLabel("Sandboxie Path:", 20, 465, 230, 18)
    $g_SandboxiePathBrowseBtn = GUICtrlCreateButton("📂 ...", 260, 461, 40, $btnH)
    GUICtrlSetTip($g_SandboxiePathBrowseBtn, "Set Path to SandMan.exe")
    $g_SandboxiePathEdit      = GUICtrlCreateInput("", 20, 492, 280, 22)
    GUICtrlSetTip($g_SandboxiePathEdit, "Use '?' for portable drive letter.")
    Local $sandboxiePath = IniRead(@ScriptDir & "\App\Settings.ini", "GLOBAL", "SandboxiePath", "")
    GUICtrlSetData($g_SandboxiePathEdit, $sandboxiePath)

    ; ---- Variables ListView (Left) ----
    $g_VarsListView = GUICtrlCreateListView("#|Set User Variables", $listviewX, $listviewY, 250, 320, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL)
    GUICtrlSetFont($g_VarsListView, 10, 500, 0, "Consolas")
    GUICtrlSendMsg($g_VarsListView, $LVM_SETCOLUMNWIDTH, 0, 30)
    GUICtrlSendMsg($g_VarsListView, $LVM_SETCOLUMNWIDTH, 1, 200)
    $g_VarsListViewMenu = GUICtrlCreateContextMenu($g_VarsListView)
    $g_VarsDeleteMenuItem = GUICtrlCreateMenuItem("Delete", $g_VarsListViewMenu)

    _Vars_ReadItems()
    _Vars_ListViewPopulate()

    ; ---- Unified Edit Field ----
    $g_CombinedEdit = GUICtrlCreateInput("", 80, 382, 490, 20)
    GUICtrlSetTip($g_CombinedEdit, "Add System/User EnvVars or SymLinks" & @CRLF & "(Type, press 'Enter' Click '➕')")

    $g_VarsAddBtn    = GUICtrlCreateButton("➕", 20, 380, $btnH, $btnH)
    GUICtrlSetTip($g_VarsAddBtn, "Add to Variables")
    $g_VarsSaveBtn   = GUICtrlCreateButton("💾", 50, 380, $btnH, $btnH)
    GUICtrlSetTip($g_VarsSaveBtn, "Save Variable")
    $g_SymLinksAddBtn    = GUICtrlCreateButton("➕", 605, 380, $btnH, $btnH)
    GUICtrlSetTip($g_SymLinksAddBtn, "Add to SymLinks")
    $g_SymLinksSaveBtn   = GUICtrlCreateButton("💾", 575, 380, $btnH, $btnH)
    GUICtrlSetTip($g_SymLinksSaveBtn, "Save SymLink")

    ; ---- SymLinks ListView (Right) ----
    $g_SymLinksListView = GUICtrlCreateListView("#|SymLinks", 280, 50, 350, 320, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL)
	GUICtrlSetTip($g_SymLinksListView, "'SymLink Location~Target Folder'" & @CRLF & "Location/target separator: ~")
    GUICtrlSetFont($g_SymLinksListView, 10, 500, 0, "Consolas")
    GUICtrlSendMsg($g_SymLinksListView, $LVM_SETCOLUMNWIDTH, 0, 30)
    GUICtrlSendMsg($g_SymLinksListView, $LVM_SETCOLUMNWIDTH, 1, 300)
    $g_SymLinksListViewMenu = GUICtrlCreateContextMenu($g_SymLinksListView)
    $g_SymLinksDeleteMenuItem = GUICtrlCreateMenuItem("Delete", $g_SymLinksListViewMenu)

    _SymLinks_ReadItems()
    _SymLinks_ListViewPopulate()

    $g_SymLinksAddCheckbox = GUICtrlCreateCheckbox("Create Symlinks on Startup", 370, 433, 220, 22)
    $g_SymLinksRemoveCheckbox = GUICtrlCreateCheckbox("Remove Symlinks on Shutdown", 370, 461, 220, 22)
    $g_CreateSymlinksBtn = GUICtrlCreateButton("➕ Create Symlinks", 350, 490, 100, $btnH)
    $g_RemoveSymlinksBtn = GUICtrlCreateButton("➖ Remove Symlinks", 465, 490, 100, $btnH)
	
    If IniRead($g_SettingsIni, "SymLinks", "SymLinksAdd", "0") = "1" Then
        GUICtrlSetState($g_SymLinksAddCheckbox, $GUI_CHECKED)
    Else
        GUICtrlSetState($g_SymLinksAddCheckbox, $GUI_UNCHECKED)
    EndIf
    If IniRead($g_SettingsIni, "SymLinks", "SymLinksRemove", "0") = "1" Then
        GUICtrlSetState($g_SymLinksRemoveCheckbox, $GUI_CHECKED)
    Else
        GUICtrlSetState($g_SymLinksRemoveCheckbox, $GUI_UNCHECKED)
    EndIf
	
    _Vars_Buttons_DisableAll()
    _SymLinks_Buttons_DisableAll()
    _SymLinks_EnableDisableSymlinkControls()
EndFunc

Func Tab_Global_HandleEvents($msg)
    If $msg = $UPDATE_ON_START_KEY Then
    EndIf

    If $msg = $g_CreateSymlinksBtn Then
        _SymLink_ManualCreateGlobalSymlinks($globalIni)
    EndIf

    If $msg = $g_RemoveSymlinksBtn Then
        _SymLink_ManualRemoveGlobalSymlinks($globalIni)
    EndIf

    If $msg = $g_SandboxiePathBrowseBtn Then
        Local $path = FileOpenDialog("Select Sandboxie Folder", "", "Folders (*.exe;*.dll;*)", 2)
        If Not @error And $path <> "" Then
            Local $scriptDrive = StringLeft(@ScriptDir, 2)
            Local $selectedDrive = StringLeft($path, 2)
            If $selectedDrive = $scriptDrive Then
                $path = "?:" & StringMid($path, 3)
            EndIf
            GUICtrlSetData($g_SandboxiePathEdit, $path)
        EndIf
    EndIf

    If $msg = $g_OKBtn Then
        Local $v = (GUICtrlRead($UPDATE_ON_START_KEY) = $GUI_CHECKED) ? "1" : "0"
        Local $sandboxiePath = GUICtrlRead($g_SandboxiePathEdit)
        IniWrite(@ScriptDir & "\App\Settings.ini", "GLOBAL", "UpdateOnStart", $v)
        IniWrite(@ScriptDir & "\App\Settings.ini", "GLOBAL", "SandboxiePath", $sandboxiePath)
    EndIf
EndFunc

Func ClearListViewToWhitespace($listview, ByRef $selectedIdx, ByRef $lastSelectedIdx)
    ; Clear selection and focus for all items
    For $i = 0 To _GUICtrlListView_GetItemCount($listview) - 1
        _GUICtrlListView_SetItemSelected($listview, $i, False)
        _GUICtrlListView_SetItemFocused($listview, $i, False)
    Next
    ; Set focused item to -1 (no focus)
    _GUICtrlListView_SetItemFocused($listview, -1)
    ; Clear selection index variables
    $selectedIdx = -1
    $lastSelectedIdx = -1
EndFunc

;----------------- Variables Section Logic -----------------
Func _Vars_ReadItems()
    Local $arr[0][2], $idx = 1
    While True
        Local $key = "SetVar" & $idx
        Local $val = IniRead($g_SettingsIni, "Variables", $key, "")
        If $val = "" Then ExitLoop
        ReDim $arr[UBound($arr)+1][2]
        $arr[UBound($arr)-1][0] = $key
        $arr[UBound($arr)-1][1] = $val
        $idx += 1
    WEnd
    $g_VarsItems = $arr
EndFunc

Func _Vars_ListViewPopulate()
    _GUICtrlListView_DeleteAllItems($g_VarsListView)
    For $i = 0 To UBound($g_VarsItems) - 1
        GUICtrlCreateListViewItem(StringFormat("%d|%s", $i+1, $g_VarsItems[$i][1]), $g_VarsListView)
    Next
EndFunc

Func Vars_HandleEvents($msg)
    If $msg = $g_VarsAddBtn Then
        _Vars_AddItem()
        Return
    ElseIf $msg = $g_VarsSaveBtn Then
        _Vars_SaveEdits()
        Return
    EndIf

    ; Context menu delete
    If $msg = $g_VarsDeleteMenuItem Then
        Local $selIdx = -1
        For $i = 0 To UBound($g_VarsItems) - 1
            If _GUICtrlListView_GetItemSelected($g_VarsListView, $i) Then
                $selIdx = $i
                ExitLoop
            EndIf
        Next
        If $selIdx = -1 Then Return
        Local $result = MsgBox(33, "Confirm Variable Delete", "Are you sure you want to DELETE this environment variable?" & @CRLF & _
            "Click OK to proceed or Cancel to abort.")
        If $result = 1 Then
            _Vars_DeleteItem()
        EndIf
        Return
    EndIf

    ; ListView selection logic
    Local $selIdx = -1
    For $i = 0 To UBound($g_VarsItems) - 1
        If _GUICtrlListView_GetItemSelected($g_VarsListView, $i) Then
            $selIdx = $i
            ExitLoop
        EndIf
    Next

    If $selIdx <> $g_VarsLastSelectedIdx Then
        $g_VarsSelectedIdx = $selIdx
        $g_VarsLastSelectedIdx = $selIdx
        If $selIdx <> -1 Then
			ClearListViewToWhitespace($g_SymLinksListView, $g_SymLinksSelectedIdx, $g_SymLinksLastSelectedIdx)
            GUICtrlSetData($g_CombinedEdit, $g_VarsItems[$selIdx][1])
            _Vars_Buttons_EnableSaveOnly()
			GUICtrlSetState($g_SymLinksSaveBtn, $GUI_DISABLE)
        Else
			ClearListViewToWhitespace($g_SymLinksListView, $g_SymLinksSelectedIdx, $g_SymLinksLastSelectedIdx)
            GUICtrlSetData($g_CombinedEdit, "")
            _Vars_Buttons_DisableAll()
			GUICtrlSetState($g_SymLinksSaveBtn, $GUI_DISABLE)
        EndIf
    EndIf

    ; Enable Add button only when clicking into the field and text exists
    If $msg = $g_CombinedEdit Then
        If StringLen(GUICtrlRead($g_CombinedEdit)) > 0 Then
            _Vars_Buttons_EnableAdd()
        Else
            _Vars_Buttons_DisableAll()
        EndIf
    EndIf

    ; Disable Add button when field loses focus (any other control is clicked)
    If $msg <> $g_CombinedEdit _
        And $msg <> $g_VarsAddBtn _
        And $msg <> $g_VarsSaveBtn _
        And $msg <> $g_VarsListView _
        And $msg > 0 Then
        If BitAND(GUICtrlGetState($g_VarsAddBtn), $GUI_ENABLE) Then
        _Vars_Buttons_DisableAll()
        EndIf
    EndIf
EndFunc

Func _Vars_SaveEdits()
    If $g_VarsSelectedIdx >= 0 And $g_VarsSelectedIdx < UBound($g_VarsItems) Then
        Local $newVal = GUICtrlRead($g_CombinedEdit)
        $g_VarsItems[$g_VarsSelectedIdx][1] = $newVal
        _GUICtrlListView_SetItemText($g_VarsListView, $g_VarsSelectedIdx, $newVal, 1)
        _Vars_Buttons_DisableAll()
    EndIf
EndFunc

Func _Vars_AddItem()
    If UBound($g_VarsItems) >= 100 Then
        MsgBox(64, "Limit reached", "You can only add up to 100 variables.")
        Return
    EndIf
    Local $newVal = GUICtrlRead($g_CombinedEdit)
    If $newVal = "" Then Return
    Local $key = "SetVar" & (UBound($g_VarsItems)+1)
    ReDim $g_VarsItems[UBound($g_VarsItems)+1][2]
    $g_VarsItems[UBound($g_VarsItems)-1][0] = $key
    $g_VarsItems[UBound($g_VarsItems)-1][1] = $newVal
    _Vars_ListViewPopulate()
    GUICtrlSetData($g_CombinedEdit, "")
    $g_VarsSelectedIdx = -1
    $g_VarsLastSelectedIdx = -1
    Vars_Save()
EndFunc

Func _Vars_DeleteItem()
    Local $selIdx = -1
    For $i = 0 To UBound($g_VarsItems) - 1
        If _GUICtrlListView_GetItemSelected($g_VarsListView, $i) Then
            $selIdx = $i
            ExitLoop
        EndIf
    Next
    If $selIdx >= 0 And $selIdx < UBound($g_VarsItems) Then
        For $i = $selIdx To UBound($g_VarsItems) - 2
            $g_VarsItems[$i][0] = $g_VarsItems[$i+1][0]
            $g_VarsItems[$i][1] = $g_VarsItems[$i+1][1]
        Next
        ReDim $g_VarsItems[UBound($g_VarsItems)-1][2]
        _Vars_ListViewPopulate()
        $g_VarsSelectedIdx = -1
        $g_VarsLastSelectedIdx = -1
        GUICtrlSetData($g_CombinedEdit, "")
        Vars_Save()
    EndIf
EndFunc

Func Vars_Save()
    For $i = 1 To 100
        IniDelete($g_SettingsIni, "Variables", "SetVar" & $i)
    Next
    For $i = 0 To UBound($g_VarsItems) - 1
        IniWrite($g_SettingsIni, "Variables", $g_VarsItems[$i][0], $g_VarsItems[$i][1])
    Next
EndFunc

;----------------- SymLinks Section Logic -----------------
Func _SymLinks_ReadItems()
    Local $arr[0][2], $idx = 1
    While True
        Local $key = "SymLink" & $idx
        Local $val = IniRead($g_SettingsIni, "SymLinks", $key, "")
        If $val = "" Then ExitLoop
        ReDim $arr[UBound($arr)+1][2]
        $arr[UBound($arr)-1][0] = $key
        $arr[UBound($arr)-1][1] = $val
        $idx += 1
    WEnd
    $g_SymLinksItems = $arr
EndFunc

Func _SymLinks_ListViewPopulate()
    _GUICtrlListView_DeleteAllItems($g_SymLinksListView)
    If Not IsArray($g_SymLinksItems) Or UBound($g_SymLinksItems) = 0 Then
        _SymLinks_EnableDisableSymlinkControls()
        Return
    EndIf
    For $i = 0 To UBound($g_SymLinksItems) - 1
        GUICtrlCreateListViewItem(StringFormat("%d|%s", $i+1, $g_SymLinksItems[$i][1]), $g_SymLinksListView)
    Next
    _SymLinks_EnableDisableSymlinkControls()
EndFunc

Func SymLinks_HandleEvents($msg)
    If $msg = $g_SymLinksAddBtn Then
        _SymLinks_AddItem()
        Return
    ElseIf $msg = $g_SymLinksSaveBtn Then
        _SymLinks_SaveEdits()
        Return
    EndIf

    ; Context menu delete
    If $msg = $g_SymLinksDeleteMenuItem Then
        Local $selIdx = -1
        For $i = 0 To UBound($g_SymLinksItems) - 1
            If _GUICtrlListView_GetItemSelected($g_SymLinksListView, $i) Then
                $selIdx = $i
                ExitLoop
            EndIf
        Next
        If $selIdx = -1 Then Return
        Local $result = MsgBox(33, "Confirm SymLink Delete", "Are you sure you want to DELETE this symlink item?" & @CRLF & _
            "Click OK to proceed or Cancel to abort.")
        If $result = 1 Then
            _SymLinks_DeleteItem()
        EndIf
        Return
    EndIf

    ; ListView selection logic
    Local $selIdx = -1
    For $i = 0 To UBound($g_SymLinksItems) - 1
        If _GUICtrlListView_GetItemSelected($g_SymLinksListView, $i) Then
            $selIdx = $i
            ExitLoop
        EndIf
    Next

    If $selIdx <> $g_SymLinksLastSelectedIdx Then
        $g_SymLinksSelectedIdx = $selIdx
        $g_SymLinksLastSelectedIdx = $selIdx
        If $selIdx <> -1 Then
			ClearListViewToWhitespace($g_VarsListView, $g_VarsSelectedIdx, $g_VarsLastSelectedIdx)
            GUICtrlSetData($g_CombinedEdit, $g_SymLinksItems[$selIdx][1])
            _SymLinks_Buttons_EnableSaveOnly()
			GUICtrlSetState($g_VarsSaveBtn, $GUI_DISABLE)
        Else
			ClearListViewToWhitespace($g_VarsListView, $g_VarsSelectedIdx, $g_VarsLastSelectedIdx)
            GUICtrlSetData($g_CombinedEdit, "")
            _SymLinks_Buttons_DisableAll()
			GUICtrlSetState($g_VarsSaveBtn, $GUI_DISABLE)
        EndIf
    EndIf

    If $msg = $g_CombinedEdit Then
        If StringLen(GUICtrlRead($g_CombinedEdit)) > 0 Then
            _SymLinks_Buttons_EnableAdd()
        Else
            _SymLinks_Buttons_DisableAll()
        EndIf
    EndIf

    If $msg <> $g_CombinedEdit _
        And $msg <> $g_SymLinksAddBtn _
        And $msg <> $g_SymLinksSaveBtn _
        And $msg <> $g_SymLinksListView _
        And $msg > 0 Then
        If BitAND(GUICtrlGetState($g_SymLinksAddBtn), $GUI_ENABLE) Then
            _SymLinks_Buttons_DisableAll()
        EndIf
    EndIf

    If $msg = $g_SymLinksAddCheckbox Then
        IniWrite($g_SettingsIni, "SymLinks", "SymLinksAdd", GUICtrlRead($g_SymLinksAddCheckbox) = $GUI_CHECKED ? "1" : "0")
    EndIf
    If $msg = $g_SymLinksRemoveCheckbox Then
        IniWrite($g_SettingsIni, "SymLinks", "SymLinksRemove", GUICtrlRead($g_SymLinksRemoveCheckbox) = $GUI_CHECKED ? "1" : "0")
    EndIf
EndFunc

Func _SymLinks_SaveEdits()
    If $g_SymLinksSelectedIdx >= 0 And $g_SymLinksSelectedIdx < UBound($g_SymLinksItems) Then
        Local $newVal = GUICtrlRead($g_CombinedEdit)
        $g_SymLinksItems[$g_SymLinksSelectedIdx][1] = $newVal
        _GUICtrlListView_SetItemText($g_SymLinksListView, $g_SymLinksSelectedIdx, $newVal, 1)
        SymLinks_Save()
    EndIf
EndFunc

Func _SymLinks_AddItem()
    If UBound($g_SymLinksItems) >= 100 Then
        MsgBox(64, "Limit reached", "You can only add up to 100 symlinks.")
        Return
    EndIf
    Local $newVal = GUICtrlRead($g_CombinedEdit)
    If $newVal = "" Then Return
    Local $key = "SymLink" & (UBound($g_SymLinksItems)+1)
    ReDim $g_SymLinksItems[UBound($g_SymLinksItems)+1][2]
    $g_SymLinksItems[UBound($g_SymLinksItems)-1][0] = $key
    $g_SymLinksItems[UBound($g_SymLinksItems)-1][1] = $newVal
    _SymLinks_ListViewPopulate()
    GUICtrlSetData($g_CombinedEdit, "")
    $g_SymLinksSelectedIdx = -1
    $g_SymLinksLastSelectedIdx = -1
    SymLinks_Save()
    _SymLinks_EnableDisableSymlinkControls()
EndFunc

Func _SymLinks_DeleteItem()
    Local $selIdx = -1
    For $i = 0 To UBound($g_SymLinksItems) - 1
        If _GUICtrlListView_GetItemSelected($g_SymLinksListView, $i) Then
            $selIdx = $i
            ExitLoop
        EndIf
    Next
    If $selIdx >= 0 And $selIdx < UBound($g_SymLinksItems) Then
        For $i = $selIdx To UBound($g_SymLinksItems) - 2
            $g_SymLinksItems[$i][0] = $g_SymLinksItems[$i+1][0]
            $g_SymLinksItems[$i][1] = $g_SymLinksItems[$i+1][1]
        Next
        ReDim $g_SymLinksItems[UBound($g_SymLinksItems)-1][2]
        _SymLinks_ListViewPopulate()
        $g_SymLinksSelectedIdx = -1
        $g_SymLinksLastSelectedIdx = -1
        GUICtrlSetData($g_CombinedEdit, "")
        SymLinks_Save()
    EndIf
    _SymLinks_EnableDisableSymlinkControls()
EndFunc

Func SymLinks_Save()
    For $i = 1 To 100
        IniDelete($g_SettingsIni, "SymLinks", "SymLink" & $i)
    Next
    For $i = 0 To UBound($g_SymLinksItems) - 1
        IniWrite($g_SettingsIni, "SymLinks", $g_SymLinksItems[$i][0], $g_SymLinksItems[$i][1])
    Next
EndFunc

Func _SymLinks_EnableDisableSymlinkControls()
    Local $hasSymlinks = IsArray($g_SymLinksItems) And UBound($g_SymLinksItems) > 0
    GUICtrlSetState($g_SymLinksAddCheckbox, $hasSymlinks ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($g_SymLinksRemoveCheckbox, $hasSymlinks ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($g_CreateSymlinksBtn, $hasSymlinks ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($g_RemoveSymlinksBtn, $hasSymlinks ? $GUI_ENABLE : $GUI_DISABLE)
EndFunc
