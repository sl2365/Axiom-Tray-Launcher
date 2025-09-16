#include-once
#include "TrayMenu.au3"

Global $guiW, $guiH, $listviewX, $listviewY, $btnW, $btnH, $footer_gap
Global $g_IgnoreListFile = @ScriptDir & "\App\IgnoreList.ini"
Global $g_IgnoreEdit, $g_IgnoreBrowseBtn, $g_IgnoreBrowseFolderBtn, $g_IgnoreEditBtn, $g_IgnoreSaveBtn
Global Const $MAX_IGNORE_ITEMS = 500

Func Tab_Ignore_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("Ignore List")
    ; Multi-line, readonly Edit replaces ListView
    $g_IgnoreEdit = GUICtrlCreateEdit("", $listviewX, $listviewY, $guiW-40, $guiH-155)
    GUICtrlSetState($g_IgnoreEdit, $GUI_DISABLE)
    GUICtrlSetFont($g_IgnoreEdit, 10, 500, 0, "Consolas")
    GUICtrlSetTip($g_IgnoreEdit, "One path per line. eg:" & @CRLF & "Filename.exe" & @CRLF & "Folder" & @CRLF & "Parent\Folder")
    _Ignore_EditPopulate()
    $g_IgnoreEditBtn = GUICtrlCreateButton("✏️", $listviewX, $guiH-90, $btnH, $btnH)
    GUICtrlSetTip($g_IgnoreEditBtn, "Enable editing of the ignore list.")
    $g_IgnoreBrowseBtn = GUICtrlCreateButton("📄 ...", $listviewX + 35, $guiH-90, 40, $btnH)
    GUICtrlSetTip($g_IgnoreBrowseBtn, "Select file to exclude from scans.")
    GUICtrlSetState($g_IgnoreBrowseBtn, $GUI_DISABLE)
    $g_IgnoreBrowseFolderBtn = GUICtrlCreateButton("📁 ...", $listviewX + 85, $guiH-90, 40, $btnH)
    GUICtrlSetTip($g_IgnoreBrowseFolderBtn, "Select folder to exclude from scans.")
    GUICtrlSetState($g_IgnoreBrowseFolderBtn, $GUI_DISABLE)
    $g_IgnoreSaveBtn = GUICtrlCreateButton("💾", $listviewX + 135, $guiH-90, $btnH, $btnH)
    GUICtrlSetTip($g_IgnoreSaveBtn, "Save changes to the ignore list.")
    GUICtrlSetState($g_IgnoreSaveBtn, $GUI_DISABLE)
    _Ignore_EditPopulate()
EndFunc

Func _Ignore_EditPopulate()
    Local $lines = ""
    If FileExists($g_IgnoreListFile) Then
        Local $arr = FileReadToArray($g_IgnoreListFile)
        If IsArray($arr) Then
            For $i = 0 To UBound($arr) - 1
                If StringStripWS($arr[$i], 3) <> "" Then
                    $lines &= $arr[$i] & @CRLF
                EndIf
            Next
        EndIf
    EndIf
    GUICtrlSetData($g_IgnoreEdit, StringTrimRight($lines, 2))
EndFunc

Func Ignore_HandleEvents($msg)
    ; ---- Browse File Button ----
    If $msg = $g_IgnoreBrowseBtn Then
        If GUICtrlGetState($g_IgnoreEdit) = $GUI_DISABLE Or _IsEditReadOnly($g_IgnoreEdit) Then Return
        Local $browse = FileOpenDialog("Select file to ignore", @ScriptDir, "All files (*.*)", 1)
        If Not @error And $browse <> "" Then
            Local $path = $browse
            If StringLeft($path, 2) == StringLeft(@ScriptDir, 2) Then
                $path = "?:" & StringMid($path, 3)
            EndIf
            Local $currText = GUICtrlRead($g_IgnoreEdit)
            Local $arr = StringSplit($currText, @CRLF, 1)
            For $i = 1 To $arr[0]
                If StringStripWS($arr[$i], 3) == $path Then
                    MsgBox(64, "Duplicate", "This ignore path already exists.")
                    Return
                EndIf
            Next
            If $arr[0] >= $MAX_IGNORE_ITEMS Then
                MsgBox(64, "Limit reached", "You can only add up to " & $MAX_IGNORE_ITEMS & " ignore paths.")
                Return
            EndIf
            If $currText <> "" Then
                $currText &= @CRLF & $path
            Else
                $currText = $path
            EndIf
            GUICtrlSetData($g_IgnoreEdit, $currText)
        EndIf
        Return
    EndIf

    ; ---- Browse Folder Button ----
    If $msg = $g_IgnoreBrowseFolderBtn Then
        If GUICtrlGetState($g_IgnoreEdit) = $GUI_DISABLE Or _IsEditReadOnly($g_IgnoreEdit) Then Return
        Local $folder = FileSelectFolder("Select folder to ignore", @ScriptDir, 1)
        If Not @error And $folder <> "" Then
            Local $path = $folder
            If StringLeft($path, 2) == StringLeft(@ScriptDir, 2) Then
                $path = "?:" & StringMid($path, 3)
            EndIf
            Local $currText = GUICtrlRead($g_IgnoreEdit)
            Local $arr = StringSplit($currText, @CRLF, 1)
            For $i = 1 To $arr[0]
                If StringStripWS($arr[$i], 3) == $path Then
                    MsgBox(64, "Duplicate", "This ignore path already exists.")
                    Return
                EndIf
            Next
            If $arr[0] >= $MAX_IGNORE_ITEMS Then
                MsgBox(64, "Limit reached", "You can only add up to " & $MAX_IGNORE_ITEMS & " ignore paths.")
                Return
            EndIf
            If $currText <> "" Then
                $currText &= @CRLF & $path
            Else
                $currText = $path
            EndIf
            GUICtrlSetData($g_IgnoreEdit, $currText)
        EndIf
        Return
    EndIf

    ; ---- Edit Button ----
    If $msg = $g_IgnoreEditBtn Then
        GUICtrlSetState($g_IgnoreEdit, $GUI_ENABLE)
        GUICtrlSetState($g_IgnoreEditBtn, $GUI_DISABLE)
        GUICtrlSetState($g_IgnoreSaveBtn, $GUI_ENABLE)
        GUICtrlSetState($g_IgnoreBrowseBtn, $GUI_ENABLE)
        GUICtrlSetState($g_IgnoreBrowseFolderBtn, $GUI_ENABLE)
        Return
    EndIf

    ; ---- Save Button ----
    If $msg = $g_IgnoreSaveBtn Then
        Local $currText = GUICtrlRead($g_IgnoreEdit)
        Local $arr = StringSplit($currText, @CRLF, 1)
        If $arr[0] > $MAX_IGNORE_ITEMS Then
            MsgBox(64, "Limit reached", "You can only save up to " & $MAX_IGNORE_ITEMS & " ignore paths.")
            Return
        EndIf
        Local $hFile = FileOpen($g_IgnoreListFile, $FO_OVERWRITE)
        If $hFile = -1 Then Return
        For $i = 1 To $arr[0]
            Local $item = StringStripWS($arr[$i], 3)
            If $item <> "" Then
                FileWriteLine($hFile, $item)
            EndIf
        Next
        FileClose($hFile)
        GUICtrlSetState($g_IgnoreEdit, $GUI_DISABLE)
        GUICtrlSetState($g_IgnoreEditBtn, $GUI_ENABLE)
        GUICtrlSetState($g_IgnoreSaveBtn, $GUI_DISABLE)
        GUICtrlSetState($g_IgnoreBrowseBtn, $GUI_DISABLE)
        GUICtrlSetState($g_IgnoreBrowseFolderBtn, $GUI_DISABLE)
        Return
    EndIf
EndFunc

Func _IsEditReadOnly($idEdit)
    Return (BitAND(GUICtrlGetState($idEdit), $GUI_DISABLE) = $GUI_DISABLE)
EndFunc
