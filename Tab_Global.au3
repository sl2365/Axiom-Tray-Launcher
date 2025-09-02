; Tab_Global.au3

#include-once
;~ #include <GUIConstantsEx.au3>
;~ #include <WindowsConstants.au3>
;~ #include <GuiListView.au3>

Global $g_SettingsIni = @ScriptDir & "\App\Settings.ini"
Global $g_SandboxiePathEdit, $g_SandboxiePathBrowseBtn

;----------------- Variables Section Globals -----------------
Global $g_VarsListView, $g_VarsEdit, $g_VarsAddBtn, $g_VarsDeleteBtn, $g_VarsSaveBtn
Global $g_VarsItems[0][2], $g_VarsSelectedIdx = -1, $g_VarsLastSelectedIdx = -1

;----------------- SymLinks Section Globals -----------------
Global $g_SymLinksListView, $g_SymLinksEdit, $g_SymLinksAddBtn, $g_SymLinksDeleteBtn, $g_SymLinksSaveBtn
Global $g_SymLinksItems[0][2], $g_SymLinksSelectedIdx = -1, $g_SymLinksLastSelectedIdx = -1
Global $g_SymLinksAddCheckbox, $g_SymLinksRemoveCheckbox

Func Tab_Global_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("Global")
	
    ; ---- Global Options ----
	$UPDATE_ON_START_KEY = GUICtrlCreateCheckbox("Check for Updates on Startup", 20, 435, 200, 18)
    Local $val = IniRead(@ScriptDir & "\App\Settings.ini", "GLOBAL", "UpdateOnStart", "1")
    GUICtrlSetState($UPDATE_ON_START_KEY, ($val = "1") ? $GUI_CHECKED : $GUI_UNCHECKED)

	GUICtrlCreateLabel("Sandboxie Path:", 20, 465, 230, 18)
    $g_SandboxiePathBrowseBtn = GUICtrlCreateButton("...", 270, 461, 30, 22)
	GUICtrlSetTip($g_SandboxiePathBrowseBtn, "Set Path to SandMan.exe")
    $g_SandboxiePathEdit      = GUICtrlCreateInput("", 20, 490, 280, 22)
	GUICtrlSetTip($g_SandboxiePathEdit, "Use '?' for portable drive letter.")
    Local $sandboxiePath = IniRead(@ScriptDir & "\App\Settings.ini", "GLOBAL", "SandboxiePath", "")
    GUICtrlSetData($g_SandboxiePathEdit, $sandboxiePath)

    ; ---- Variables ListView (Left) ----
    $g_VarsListView = GUICtrlCreateListView("#|Set User Variables", $listviewX, $listviewY, 250, 300, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL)
    GUICtrlSetFont($g_VarsListView, 10, 500, 0, "Consolas")
    GUICtrlSendMsg($g_VarsListView, $LVM_SETCOLUMNWIDTH, 0, 30)
    GUICtrlSendMsg($g_VarsListView, $LVM_SETCOLUMNWIDTH, 1, 200)

    _Vars_ReadItems()
    _Vars_ListViewPopulate()

    $g_VarsEdit      = GUICtrlCreateInput("", 20, 360, 250, 20)
	GUICtrlSetTip($g_VarsEdit, "Add System or User EnvVars")
    $g_VarsAddBtn    = GUICtrlCreateButton("Add", 45, 390, $btnW, $btnH)
    $g_VarsDeleteBtn = GUICtrlCreateButton("Delete", 115, 390, $btnW, $btnH)
    $g_VarsSaveBtn  = GUICtrlCreateButton("Save", 185, 390, $btnW, $btnH)

    ; ---- SymLinks ListView (Right) ----
    $g_SymLinksListView = GUICtrlCreateListView("#|SymLinks", 280, 50, 350, 300, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL)
    GUICtrlSetFont($g_SymLinksListView, 10, 500, 0, "Consolas")
    GUICtrlSendMsg($g_SymLinksListView, $LVM_SETCOLUMNWIDTH, 0, 30)
    GUICtrlSendMsg($g_SymLinksListView, $LVM_SETCOLUMNWIDTH, 1, 300)

    _SymLinks_ReadItems()
    _SymLinks_ListViewPopulate()

    $g_SymLinksEdit      = GUICtrlCreateInput("", 280, 360, 350, 20)
	GUICtrlSetTip($g_SymLinksEdit, "'SymLink Location~Target Folder'" & @CRLF & "Use all EnvVars here.")
    $g_SymLinksAddBtn    = GUICtrlCreateButton("Add", 355, 390, $btnW, $btnH)
    $g_SymLinksDeleteBtn = GUICtrlCreateButton("Delete", 425, 390, $btnW, $btnH)
    $g_SymLinksSaveBtn  = GUICtrlCreateButton("Save", 495, 390, $btnW, $btnH)
	Global $g_SymLinksAddCheckbox = GUICtrlCreateCheckbox("Create Symlinks on Startup", 370, 433, 220, 22)
	Global $g_SymLinksRemoveCheckbox = GUICtrlCreateCheckbox("Remove Symlinks on Shutdown", 370, 461, 220, 22)
	Global $g_CreateSymlinksBtn = GUICtrlCreateButton("Create Symlinks", 350, 490, 100, $btnH)
	Global $g_RemoveSymlinksBtn = GUICtrlCreateButton("Remove Symlinks", 465, 490, 100, $btnH)
	
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
EndFunc

Func Tab_Global_HandleEvents($msg)
    ; If the event is for the UpdateOnStart checkbox or a button (like Save/OK)
    If $msg = $UPDATE_ON_START_KEY Then
    EndIf

	If $msg = $g_CreateSymlinksBtn Then
		MsgBox(64, "Debug", "Create Symlinks Clicked")
        _SymLink_ManualCreateGlobalSymlinks($globalIni)
    EndIf

    If $msg = $g_RemoveSymlinksBtn Then
		MsgBox(64, "Debug", "Remove Symlinks Clicked")
        _SymLink_ManualRemoveGlobalSymlinks($globalIni)
    EndIf

	If $msg = $g_SandboxiePathBrowseBtn Then
        Local $path = FileOpenDialog("Select Sandboxie Folder", "", "Folders (*.exe;*.dll;*)", 2)
        If Not @error And $path <> "" Then
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
    ElseIf $msg = $g_VarsDeleteBtn Then
		Local $selIdx = -1
        For $i = 0 To UBound($g_VarsItems) - 1
            If _GUICtrlListView_GetItemSelected($g_VarsListView, $i) Then
                $selIdx = $i
                ExitLoop
            EndIf
        Next
        If $selIdx = -1 Then Return ; Nothing selected
        Local $result = MsgBox(33, "Confirm Variable Delete", "Are you sure you want to DELETE this environment variable?" & @CRLF & _
            "Click OK to proceed or Cancel to abort.")
        If $result = 1 Then ; OK
            _Vars_DeleteItem()
        EndIf
        Return
    ElseIf $msg = $g_VarsSaveBtn Then
        _Vars_SaveEdits()
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
            GUICtrlSetData($g_VarsEdit, $g_VarsItems[$selIdx][1])
        Else
            GUICtrlSetData($g_VarsEdit, "")
        EndIf
    EndIf
EndFunc

Func _Vars_SaveEdits()
    If $g_VarsSelectedIdx >= 0 And $g_VarsSelectedIdx < UBound($g_VarsItems) Then
        Local $newVal = GUICtrlRead($g_VarsEdit)
        $g_VarsItems[$g_VarsSelectedIdx][1] = $newVal
        _GUICtrlListView_SetItemText($g_VarsListView, $g_VarsSelectedIdx, $newVal, 1)
        Vars_Save()
    EndIf
EndFunc

Func _Vars_AddItem()
    If UBound($g_VarsItems) >= 100 Then
        MsgBox(64, "Limit reached", "You can only add up to 100 variables.")
        Return
    EndIf
    Local $newVal = GUICtrlRead($g_VarsEdit)
    If $newVal = "" Then Return
    Local $key = "SetVar" & (UBound($g_VarsItems)+1)
    ReDim $g_VarsItems[UBound($g_VarsItems)+1][2]
    $g_VarsItems[UBound($g_VarsItems)-1][0] = $key
    $g_VarsItems[UBound($g_VarsItems)-1][1] = $newVal
    _Vars_ListViewPopulate()
    GUICtrlSetData($g_VarsEdit, "")
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
        GUICtrlSetData($g_VarsEdit, "")
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
    For $i = 0 To UBound($g_SymLinksItems) - 1
;~         ConsoleWrite("Array[" & $i & "]: " & $g_SymLinksItems[$i][1] & @CRLF)
    Next
EndFunc

Func _SymLinks_ListViewPopulate()
    _GUICtrlListView_DeleteAllItems($g_SymLinksListView)
    If Not IsArray($g_SymLinksItems) Or UBound($g_SymLinksItems) = 0 Then
        ConsoleWrite("SymLinks ListView: No items to show." & @CRLF)
        Return
    EndIf
    For $i = 0 To UBound($g_SymLinksItems) - 1
;~         ConsoleWrite(StringFormat("SymLinks ListView: Adding item #%d: %s", $i+1, $g_SymLinksItems[$i][1]) & @CRLF)
        GUICtrlCreateListViewItem(StringFormat("%d|%s", $i+1, $g_SymLinksItems[$i][1]), $g_SymLinksListView)
    Next
EndFunc

Func SymLinks_HandleEvents($msg)
    If $msg = $g_SymLinksAddBtn Then
        _SymLinks_AddItem()
        Return
    ElseIf $msg = $g_SymLinksDeleteBtn Then
		Local $selIdx = -1
        For $i = 0 To UBound($g_SymLinksItems) - 1
            If _GUICtrlListView_GetItemSelected($g_SymLinksListView, $i) Then
                $selIdx = $i
                ExitLoop
            EndIf
        Next
        If $selIdx = -1 Then Return ; Nothing selected
        Local $result = MsgBox(33, "Confirm SymLink Delete", "Are you sure you want to DELETE this symlink item?" & @CRLF & _
            "Click OK to proceed or Cancel to abort.")
        If $result = 1 Then ; OK
            _SymLinks_DeleteItem()
        EndIf
        Return
    ElseIf $msg = $g_SymLinksSaveBtn Then
        _SymLinks_SaveEdits()
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
            GUICtrlSetData($g_SymLinksEdit, $g_SymLinksItems[$selIdx][1])
        Else
            GUICtrlSetData($g_SymLinksEdit, "")
        EndIf
    EndIf
	
	; Handle checkbox state changes
	If $msg = $g_SymLinksAddCheckbox Then
		IniWrite($g_SettingsIni, "SymLinks", "SymLinksAdd", GUICtrlRead($g_SymLinksAddCheckbox) = $GUI_CHECKED ? "1" : "0")
	EndIf
	If $msg = $g_SymLinksRemoveCheckbox Then
		IniWrite($g_SettingsIni, "SymLinks", "SymLinksRemove", GUICtrlRead($g_SymLinksRemoveCheckbox) = $GUI_CHECKED ? "1" : "0")
	EndIf
EndFunc

Func _SymLinks_SaveEdits()
    If $g_SymLinksSelectedIdx >= 0 And $g_SymLinksSelectedIdx < UBound($g_SymLinksItems) Then
        Local $newVal = GUICtrlRead($g_SymLinksEdit)
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
    Local $newVal = GUICtrlRead($g_SymLinksEdit)
    If $newVal = "" Then Return
    Local $key = "SymLink" & (UBound($g_SymLinksItems)+1)
    ReDim $g_SymLinksItems[UBound($g_SymLinksItems)+1][2]
    $g_SymLinksItems[UBound($g_SymLinksItems)-1][0] = $key
    $g_SymLinksItems[UBound($g_SymLinksItems)-1][1] = $newVal
    _SymLinks_ListViewPopulate()
    GUICtrlSetData($g_SymLinksEdit, "")
    $g_SymLinksSelectedIdx = -1
    $g_SymLinksLastSelectedIdx = -1
    SymLinks_Save()
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
        GUICtrlSetData($g_SymLinksEdit, "")
        SymLinks_Save()
    EndIf
EndFunc

Func SymLinks_Save()
    For $i = 1 To 100
        IniDelete($g_SettingsIni, "SymLinks", "SymLink" & $i)
    Next
    For $i = 0 To UBound($g_SymLinksItems) - 1
        IniWrite($g_SettingsIni, "SymLinks", $g_SymLinksItems[$i][0], $g_SymLinksItems[$i][1])
    Next
EndFunc
