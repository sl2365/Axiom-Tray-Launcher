; Tab_Ignore.au3

#include-once
#include "TrayMenu.au3"

Global $guiW, $guiH, $listviewX, $listviewY, $btnW, $btnH, $footer_gap
Global $g_IgnoreListFile = @ScriptDir & "\App\IgnoreList.ini"
Global $g_IgnoreTabListView, $g_IgnoreEdit, $g_IgnoreBrowseBtn, $g_IgnoreAddBtn, $g_IgnoreDeleteBtn
Global $g_IgnoreItems[0][2]
Global Const $MAX_IGNORE_ITEMS = 300

Func Tab_Ignore_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("Ignore List")
    $g_IgnoreTabListView = GUICtrlCreateListView("#|Paths", $listviewX, $listviewY, 360, 320, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL)
    GUICtrlSetFont($g_IgnoreTabListView, 10, 500, 0, "Consolas")
    GUICtrlSendMsg($g_IgnoreTabListView, $LVM_SETCOLUMNWIDTH, 0, 40)
    GUICtrlSendMsg($g_IgnoreTabListView, $LVM_SETCOLUMNWIDTH, 1, 300)
    _Ignore_ReadItems()
    _Ignore_ListViewPopulate()
    $g_IgnoreEdit = GUICtrlCreateInput("", $listviewX, 380, 320, 20)
    $g_IgnoreBrowseBtn = GUICtrlCreateButton("...", 350, 379, 30, 22)
	GUICtrlSetTip($g_IgnoreBrowseBtn, "Select files to exclude from scans.")
    $g_IgnoreAddBtn = GUICtrlCreateButton("Add", $listviewX, 440, $btnW, $btnH)
	GUICtrlSetTip($g_IgnoreAddBtn, "Max: " & $MAX_IGNORE_ITEMS & " items")
    $g_IgnoreDeleteBtn = GUICtrlCreateButton("Delete", 100, 440, $btnW, $btnH)
EndFunc

Func _Ignore_ReadItems()
    Local $arr[0][2]
    If FileExists($g_IgnoreListFile) Then
        Local $lines = FileReadToArray($g_IgnoreListFile)
        If IsArray($lines) Then
            For $i = 0 To UBound($lines) - 1
                Local $val = StringStripWS($lines[$i], 3)
                If $val <> "" Then
                    ReDim $arr[UBound($arr)+1][2]
                    $arr[UBound($arr)-1][0] = "Ignore" & (UBound($arr))
                    $arr[UBound($arr)-1][1] = $val
                EndIf
            Next
        EndIf
    EndIf
    $g_IgnoreItems = $arr
EndFunc

Func _Ignore_ListViewPopulate()
    _GUICtrlListView_DeleteAllItems($g_IgnoreTabListView)
    For $i = 0 To UBound($g_IgnoreItems) - 1
        GUICtrlCreateListViewItem(StringFormat("%d|%s", $i+1, $g_IgnoreItems[$i][1]), $g_IgnoreTabListView)
    Next
EndFunc

Func Ignore_HandleEvents($msg)
    ; ---- Add Button ----
    If $msg = $g_IgnoreAddBtn Then
        _Ignore_AddItem()
        Return
    EndIf

    ; ---- Delete Button ----
    If $msg = $g_IgnoreDeleteBtn Then
        ; Always get currently selected index!
        Local $selIdx = -1
        For $i = 0 To UBound($g_IgnoreItems) - 1
            If _GUICtrlListView_GetItemSelected($g_IgnoreTabListView, $i) Then
                $selIdx = $i
                ExitLoop
            EndIf
        Next
        If $selIdx = -1 Then Return ; Nothing selected
        Local $result = MsgBox(33, "Confirm Ignore Delete", "Are you sure you want to DELETE this ignore path?" & @CRLF & _
            "Click OK to proceed or Cancel to abort.")
        If $result = 1 Then ; OK
            _Ignore_DeleteItem($selIdx)
        EndIf
        Return
    EndIf

    ; ---- Browse Button ----
    If $msg = $g_IgnoreBrowseBtn Then
        Local $browse = FileOpenDialog("Select file to ignore", @ScriptDir, "All files (*.*)", 1 + 2)
        If Not @error And $browse <> "" Then
            Local $path = $browse
            If StringLeft($path, 2) == StringLeft(@ScriptDir, 2) Then
                $path = "?" & StringTrimLeft($path, 2)
            EndIf
            GUICtrlSetData($g_IgnoreEdit, $path)
        EndIf
        Return
    EndIf

    ; ---- ListView selection logic (ALWAYS RUN, just like your other tabs!) ----
    Local $selIdx = -1
    For $i = 0 To UBound($g_IgnoreItems) - 1
        If _GUICtrlListView_GetItemSelected($g_IgnoreTabListView, $i) Then
            $selIdx = $i
            ExitLoop
        EndIf
    Next

    Static $lastIdx = -1
    If $selIdx <> $lastIdx Then
        If $selIdx <> -1 And $selIdx < UBound($g_IgnoreItems) Then
            GUICtrlSetData($g_IgnoreEdit, _ResolvePath($g_IgnoreItems[$selIdx][1], @ScriptDir))
        Else
            GUICtrlSetData($g_IgnoreEdit, "")
        EndIf
        $lastIdx = $selIdx
    EndIf
EndFunc

Func _Ignore_AddItem()
    If UBound($g_IgnoreItems) >= $MAX_IGNORE_ITEMS Then
        MsgBox(64, "Limit reached", "You can only add up to " & $MAX_IGNORE_ITEMS & " ignore paths.")
        Return
    EndIf
    Local $newVal = GUICtrlRead($g_IgnoreEdit)
    If $newVal = "" Then Return
    ; Check for duplicate
    For $i = 0 To UBound($g_IgnoreItems) - 1
        If $g_IgnoreItems[$i][1] == $newVal Then
            MsgBox(64, "Duplicate", "This ignore path already exists.")
            Return
        EndIf
    Next
    Local $key = "Ignore" & (UBound($g_IgnoreItems)+1)
    ReDim $g_IgnoreItems[UBound($g_IgnoreItems)+1][2]
    $g_IgnoreItems[UBound($g_IgnoreItems)-1][0] = $key
    $g_IgnoreItems[UBound($g_IgnoreItems)-1][1] = $newVal
    _Ignore_ListViewPopulate()
    GUICtrlSetData($g_IgnoreEdit, "")
    Ignore_Save()
EndFunc

Func _Ignore_DeleteItem($selIdx)
    If $selIdx >= 0 And $selIdx < UBound($g_IgnoreItems) Then
        For $i = $selIdx To UBound($g_IgnoreItems) - 2
            $g_IgnoreItems[$i][0] = $g_IgnoreItems[$i+1][0]
            $g_IgnoreItems[$i][1] = $g_IgnoreItems[$i+1][1]
        Next
        ReDim $g_IgnoreItems[UBound($g_IgnoreItems)-1][2]
        _Ignore_ListViewPopulate()
        GUICtrlSetData($g_IgnoreEdit, "")
        Ignore_Save()
    EndIf
EndFunc

Func Ignore_Save()
    Local $hFile = FileOpen($g_IgnoreListFile, $FO_OVERWRITE)
    If $hFile = -1 Then Return
    For $i = 0 To UBound($g_IgnoreItems) - 1
        Local $item = StringStripWS($g_IgnoreItems[$i][1], 3)
        If $item <> "" Then
            FileWriteLine($hFile, $item)
        EndIf
    Next
    FileClose($hFile)
EndFunc
