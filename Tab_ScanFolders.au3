; Tab_ScanFolders.au3

#include-once
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>

Global $g_ScanListView, $g_ScanFolderEdit, $g_ScanDepthCombo, $g_ScanExtEdit
Global $g_ScanAddBtn, $g_ScanDeleteBtn, $g_ScanSaveBtn, $g_ScanSelectedIdx = -1
Global $g_SettingsIni = @ScriptDir & "\App\Settings.ini"
Global $g_ScanItems[0][3]
Global $g_ScanFoldersBtn
Global $g_ScanFoldersBrowseBtn
Global $g_ScanLastSelectedIdx = -1 ; Track last selected item for selection change detection

; Reset combo box to valid values and no selection
Func _ScanFolders_ResetCombo()
    GUICtrlSetData($g_ScanDepthCombo, "") ; No selection
    GUICtrlSetData($g_ScanDepthCombo, "0|1|2|3|4|5|6|7|8|9|10")
EndFunc

Func Tab_ScanFolders_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("Scan Folders")
    ; Add "#" column to the far left for row index
    $g_ScanListView = GUICtrlCreateListView("#|Folder|Depth|Ext", $listviewX, $listviewY, 390, $guiH-225, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL)
    GUICtrlSetFont($g_ScanListView, 10, 400, 0, "Consolas")
    GUICtrlSendMsg($g_ScanListView, $LVM_SETCOLUMNWIDTH, 0, 30)   ; # column
    GUICtrlSendMsg($g_ScanListView, $LVM_SETCOLUMNWIDTH, 1, 220)  ; Folder
    GUICtrlSendMsg($g_ScanListView, $LVM_SETCOLUMNWIDTH, 2, 50)   ; Depth
    GUICtrlSendMsg($g_ScanListView, $LVM_SETCOLUMNWIDTH, 3, 60)   ; Ext

    _ScanFolders_ReadItems()
    _ScanFolders_ListViewPopulate()
    Local $fieldLeft = 420
    GUICtrlCreateLabel("Folder:", 30, $guiH-160, 60, 22)
    $g_ScanFolderEdit = GUICtrlCreateInput("", 30 + 60, $guiH-164, 270, 20)
	GUICtrlSetTip($g_ScanFolderEdit, "Use '?' for portable drive letter.")
    $g_ScanFoldersBrowseBtn = GUICtrlCreateButton("...", 30 + 340, $guiH-165, 30, 22)
	GUICtrlSetTip($g_ScanFoldersBrowseBtn, "Select Folder to Scan")

    GUICtrlCreateLabel("Depth:", 30, $guiH-130, 60, 22)
    $g_ScanDepthCombo = GUICtrlCreateCombo("", 30+60, $guiH-134, 60, 20)
    For $i = 0 To 10
        GUICtrlSetData($g_ScanDepthCombo, $i)
    Next

    GUICtrlCreateLabel("Extension:", 30, $guiH-100, 60, 22)
    $g_ScanExtEdit = GUICtrlCreateInput("", 90, $guiH-104, 310, 20)
	GUICtrlSetTip($g_ScanExtEdit, "Specify files to scan by Extension")

    $g_ScanAddBtn = GUICtrlCreateButton("Add", 125, $guiH-75, $btnW, $btnH)
	GUICtrlSetTip($g_ScanAddBtn, "Maximum limit: 100 Folders.")
    $g_ScanDeleteBtn = GUICtrlCreateButton("Delete", 195, $guiH-75, $btnW, $btnH)
    $g_ScanSaveBtn = GUICtrlCreateButton("Save", 265, $guiH-75, $btnW, $btnH)
EndFunc

Func _ScanFolders_ReadItems()
    Local $arr[0][3]
    Local $idx = 1
    While True
        Local $folder = IniRead($g_SettingsIni, "ScannedPaths", "Scan" & $idx, "")
        If $folder = "" Then ExitLoop
        Local $depth = IniRead($g_SettingsIni, "ScannedPaths", "Scan" & $idx & "Depth", "")
        Local $ext = IniRead($g_SettingsIni, "ScannedPaths", "Scan" & $idx & "Ext", "")
        ReDim $arr[UBound($arr) + 1][3]
        $arr[UBound($arr) - 1][0] = $folder
        $arr[UBound($arr) - 1][1] = $depth
        $arr[UBound($arr) - 1][2] = $ext
        $idx += 1
    WEnd
    $g_ScanItems = $arr
EndFunc

Func _ScanFolders_ListViewPopulate()
    _GUICtrlListView_DeleteAllItems($g_ScanListView)
    For $i = 0 To UBound($g_ScanItems) - 1
        GUICtrlCreateListViewItem(StringFormat("%d|%s|%s|%s", $i+1, $g_ScanItems[$i][0], $g_ScanItems[$i][1], $g_ScanItems[$i][2]), $g_ScanListView)
    Next
EndFunc

; This function should be called from your GUI event loop with the message param
Func ScanFolders_HandleEvents($msg)
    If $msg = $g_ScanAddBtn Then
        _ScanFolders_AddItem()
        Return
    ElseIf $msg = $g_ScanDeleteBtn Then
		Local $selIdx = -1
		For $i = 0 To UBound($g_ScanItems) - 1
			If _GUICtrlListView_GetItemSelected($g_ScanListView, $i) Then
				$selIdx = $i
				ExitLoop
			EndIf
		Next
		If $selIdx = -1 Then Return ; Nothing selected

		Local $result = MsgBox(33, "Confirm Scan Folder Delete", _
			"Are you sure you want to DELETE this scan folder item?" & @CRLF & _
			"Click OK to proceed or Cancel to abort.")
		If $result = 1 Then ; OK
			_ScanFolders_DeleteItem()
		EndIf
        Return
    ElseIf $msg = $g_ScanSaveBtn Then
        _ScanFolders_SaveEdits()
        Return
    ElseIf $msg = $g_ScanFoldersBrowseBtn Then
        Local $sSelected = FileSelectFolder("Choose a folder...", "")
        If @error = 0 And $sSelected <> "" Then
            GUICtrlSetData($g_ScanFolderEdit, $sSelected)
        EndIf
        Return
    EndIf

    ; Selection logic: Only update fields if selection changed
    Local $selIdx = -1
    For $i = 0 To UBound($g_ScanItems) - 1
        If _GUICtrlListView_GetItemSelected($g_ScanListView, $i) Then
            $selIdx = $i
            ExitLoop
        EndIf
    Next

    If $selIdx <> $g_ScanLastSelectedIdx Then
        $g_ScanSelectedIdx = $selIdx
        $g_ScanLastSelectedIdx = $selIdx

        If $selIdx <> -1 Then
            GUICtrlSetData($g_ScanFolderEdit, $g_ScanItems[$selIdx][0])
            If $g_ScanItems[$selIdx][1] = "" Or Not StringIsDigit($g_ScanItems[$selIdx][1]) Then
                _ScanFolders_ResetCombo()
            Else
                GUICtrlSetData($g_ScanDepthCombo, "0|1|2|3|4|5|6|7|8|9|10", $g_ScanItems[$selIdx][1])
            EndIf
            GUICtrlSetData($g_ScanExtEdit, $g_ScanItems[$selIdx][2])
        Else
            GUICtrlSetData($g_ScanFolderEdit, "")
            _ScanFolders_ResetCombo()
            GUICtrlSetData($g_ScanExtEdit, "")
        EndIf
    EndIf
EndFunc

Func _ScanFolders_SaveEdits()
    If $g_ScanSelectedIdx >= 0 And $g_ScanSelectedIdx < UBound($g_ScanItems) Then
        Local $folder = GUICtrlRead($g_ScanFolderEdit)
        Local $depth = GUICtrlRead($g_ScanDepthCombo)
        Local $ext = GUICtrlRead($g_ScanExtEdit)
        $g_ScanItems[$g_ScanSelectedIdx][0] = $folder
        $g_ScanItems[$g_ScanSelectedIdx][1] = $depth
        $g_ScanItems[$g_ScanSelectedIdx][2] = $ext

        ; Update only the edited row
        _GUICtrlListView_SetItemText($g_ScanListView, $g_ScanSelectedIdx, StringFormat("%d", $g_ScanSelectedIdx+1), 0)
        _GUICtrlListView_SetItemText($g_ScanListView, $g_ScanSelectedIdx, $folder, 1)
        _GUICtrlListView_SetItemText($g_ScanListView, $g_ScanSelectedIdx, $depth, 2)
        _GUICtrlListView_SetItemText($g_ScanListView, $g_ScanSelectedIdx, $ext, 3)

        Tab_ScanFolders_Save()
    EndIf
EndFunc

Func _ScanFolders_FindItemIdx($folder, $depth, $ext)
    For $i = 0 To UBound($g_ScanItems) - 1
        If $g_ScanItems[$i][0] = $folder And $g_ScanItems[$i][1] = $depth And $g_ScanItems[$i][2] = $ext Then Return $i
    Next
    Return -1
EndFunc

Func _ScanFolders_AddItem()
    ; Enforce a limit of 100 scan paths
    If UBound($g_ScanItems) >= 100 Then
        MsgBox(64, "Limit reached", "You can only add up to 100 scan paths.")
        Return
    EndIf

    Local $folder = StringStripWS(GUICtrlRead($g_ScanFolderEdit), 3)
    Local $depth = GUICtrlRead($g_ScanDepthCombo)
    Local $ext   = GUICtrlRead($g_ScanExtEdit)
    If $folder = "" Then Return

    If _ScanFolders_FindItemIdx($folder, $depth, $ext) <> -1 Then Return

    ReDim $g_ScanItems[UBound($g_ScanItems) + 1][3]
    $g_ScanItems[UBound($g_ScanItems) - 1][0] = $folder
    $g_ScanItems[UBound($g_ScanItems) - 1][1] = $depth
    $g_ScanItems[UBound($g_ScanItems) - 1][2] = $ext
    _ScanFolders_ListViewPopulate()
    GUICtrlSetData($g_ScanFolderEdit, "")
    _ScanFolders_ResetCombo()
    GUICtrlSetData($g_ScanExtEdit, "")
    $g_ScanSelectedIdx = -1
    $g_ScanLastSelectedIdx = -1
    Tab_ScanFolders_Save()
EndFunc

Func _ScanFolders_DeleteItem()
    Local $selIdx = -1
    For $i = 0 To UBound($g_ScanItems) - 1
        If _GUICtrlListView_GetItemSelected($g_ScanListView, $i) Then
            $selIdx = $i
            ExitLoop
        EndIf
    Next

    If $selIdx >= 0 And $selIdx < UBound($g_ScanItems) Then
        For $i = $selIdx To UBound($g_ScanItems) - 2
            $g_ScanItems[$i][0] = $g_ScanItems[$i + 1][0]
            $g_ScanItems[$i][1] = $g_ScanItems[$i + 1][1]
            $g_ScanItems[$i][2] = $g_ScanItems[$i + 1][2]
        Next
        ReDim $g_ScanItems[UBound($g_ScanItems) - 1][3]
        _ScanFolders_ListViewPopulate()
        $g_ScanSelectedIdx = -1
        $g_ScanLastSelectedIdx = -1
        GUICtrlSetData($g_ScanFolderEdit, "")
        _ScanFolders_ResetCombo()
        GUICtrlSetData($g_ScanExtEdit, "")
        Tab_ScanFolders_Save()
    EndIf
EndFunc

Func Tab_ScanFolders_Save()
    For $i = 1 To 100
        IniDelete($g_SettingsIni, "ScannedPaths", "Scan" & $i)
        IniDelete($g_SettingsIni, "ScannedPaths", "Scan" & $i & "Depth")
        IniDelete($g_SettingsIni, "ScannedPaths", "Scan" & $i & "Ext")
    Next
    For $i = 0 To UBound($g_ScanItems) - 1
        IniWrite($g_SettingsIni, "ScannedPaths", "Scan" & ($i + 1), $g_ScanItems[$i][0])
        IniWrite($g_SettingsIni, "ScannedPaths", "Scan" & ($i + 1) & "Depth", $g_ScanItems[$i][1])
        IniWrite($g_SettingsIni, "ScannedPaths", "Scan" & ($i + 1) & "Ext", $g_ScanItems[$i][2])
    Next
EndFunc
