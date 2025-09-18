; Tab_Apps.au3

#include-once
#include <GuiTreeView.au3>
#include <File.au3>
#include <Array.au3>
#include <WindowsConstants.au3>
#include <StructureConstants.au3>
#include "Utils.au3"

Global $g_CategoryIniDir = @ScriptDir & "\App"
Global $g_AppTreeView, $g_AppFields, $g_IniFiles, $g_AppSections, $g_SelectedIni, $g_SelectedSection
Global $g_SaveBtn, $g_IgnoreBtn, $g_IgnoreEdit, $g_RefreshBtn
Global $g_AppTreeItemToSection = ObjCreate("Scripting.Dictionary")
Global $g_AppTreeItemToIni = ObjCreate("Scripting.Dictionary")
Global $g_AppLastSelectedItem = 0
Global $guiW, $guiH, $listviewX, $listviewY, $btnW, $btnH, $footer_gap, $hGUI
Global $g_AppCategoryTreeItems = ObjCreate("Scripting.Dictionary")
Global $g_AppsSymlinksEdit
Global Const $ES_MULTILINE = 0x0004
Global Const $ES_WANTRETURN = 0x1000
Global Const $ES_AUTOHSCROLL  = 0x0080
Global Const $ES_AUTOVSCROLL  = 0x0020
Global $g_AppTreeViewMenu, $g_AppMenuDeleteItem, $g_AppMenuEditIniItem
Global $g_RightClickedItemID = 0

Func Tab_Apps_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("Apps")
    Local $x = 290

    ; Right side fields container
    $g_AppFields = ObjCreate("Scripting.Dictionary")
    ; SetEnv comes after Arguments
    Local $fields = ["ButtonText", "RunFile", "RunAsAdmin", "WorkDir", "Arguments", "SetEnv", "SingleInstance", "Sandboxie", "SandboxName", "Category", "Fave", "Hide", "Symlinks"]
    For $i = 0 To UBound($fields) - 1
        GUICtrlCreateLabel($fields[$i] & ":", $x, 60 + $i * 27, 70, 18)
        If StringInStr("RunAsAdmin SingleInstance Sandboxie Fave Hide", $fields[$i]) Then
            $g_AppFields($fields[$i]) = GUICtrlCreateCheckbox("", $x + 90, 56 + $i * 27, 18, 18)
        ElseIf $fields[$i] = "Symlinks" Then
            ; Do not create input here, will be created manually below
            ContinueLoop
        Else
            ; Ensure SetEnv is a single-line input (not multiline edit)
            $g_AppFields($fields[$i]) = GUICtrlCreateInput("", $x + 90, 56 + $i * 27, 250, 18)
        EndIf
    Next

    ; After the loop, create SymLinkCreate checkbox and Symlinks field, both manually positioned
    Local $symlinkY = 60 + (UBound($fields) - 1) * 27
    $g_AppFields("SymLinkCreate") = GUICtrlCreateCheckbox("", $x + 90, $symlinkY - 4, 18, 18)
	$g_AppsSymlinksEdit = GUICtrlCreateEdit("", $x, $symlinkY + 20, 340, 120, _
    BitOR($ES_MULTILINE, $ES_WANTRETURN, $ES_AUTOHSCROLL, $ES_AUTOVSCROLL, $WS_VSCROLL, $WS_HSCROLL))
    $g_AppTreeView = GUICtrlCreateTreeView($listviewX, $listviewY, $x + 90, $guiH-155, BitOR($TVS_HASBUTTONS, $TVS_LINESATROOT, $TVS_SHOWSELALWAYS))
    $g_IgnoreBtn = GUICtrlCreateButton("🚫", 177, 485, $btnH, $btnH)
    GUICtrlSetTip($g_IgnoreBtn, "Add to Ignore List and Delete.")
    $g_SaveBtn = GUICtrlCreateButton("💾", 212, 485, $btnH, $btnH)
    GUICtrlSetTip($g_SaveBtn, "Save before selecting items in Tree.")
	$g_RefreshBtn = GUICtrlCreateButton("🔄", 247, 485, $btnH, $btnH)
    GUICtrlSetTip($g_RefreshBtn, "Save & Refresh Tree")

    ; Create context menu for TreeView and delete item
    $g_AppTreeViewMenu = GUICtrlCreateContextMenu($g_AppTreeView)
    $g_AppMenuDeleteItem = GUICtrlCreateMenuItem("Delete", $g_AppTreeViewMenu)
	$g_AppMenuEditIniItem = GUICtrlCreateMenuItem("Edit INI File", $g_AppTreeViewMenu)
	
    ; Disable all buttons on startup
    _AppTab_Buttons_Disable()
    _AppTab_LoadIniFilesAndApps()  ; Sets up dictionaries and reads INI files
    _AppTab_ClearTree1()           ; Creates TreeView control + context menu
    _AppTab_RecreateTree()         ; Populates the tree with data
    _AppTab_DisableFields()
EndFunc

Func _AppTab_LoadIniFilesAndApps()
    Global $g_AppSections
    $g_AppSections = ObjCreate("Scripting.Dictionary")
    $g_AppTreeItemToSection = ObjCreate("Scripting.Dictionary")
    $g_AppTreeItemToIni = ObjCreate("Scripting.Dictionary")
    $g_AppCategoryTreeItems = ObjCreate("Scripting.Dictionary")
    $g_IniFiles = _FileListToArray($g_CategoryIniDir, "*.ini", 1)
    If @error Or Not IsArray($g_IniFiles) Then Return
EndFunc

Func _AppTab_ClearTree1()
    ; Clear previous tree nodes
    GUICtrlDelete($g_AppTreeView)
    $g_AppTreeView = GUICtrlCreateTreeView($listviewX, $listviewY, 250, $guiH-155, BitOR($TVS_HASBUTTONS, $TVS_LINESATROOT, $TVS_SHOWSELALWAYS))
    
    ; Re-create context menu after deletion
    $g_AppTreeViewMenu = GUICtrlCreateContextMenu($g_AppTreeView)
    $g_AppMenuDeleteItem = GUICtrlCreateMenuItem("Delete", $g_AppTreeViewMenu)
	$g_AppMenuEditIniItem = GUICtrlCreateMenuItem("Edit INI File", $g_AppTreeViewMenu)
EndFunc

Func _AppTab_ClearTree2()
    ; Clear previous tree nodes
    _GUICtrlTreeView_DeleteAll(GUICtrlGetHandle($g_AppTreeView))
EndFunc

Func _AppTab_RecreateTree()
    Local $favourites = ObjCreate("Scripting.Dictionary")
    Local $categories = ObjCreate("Scripting.Dictionary")
    For $f = 1 To $g_IniFiles[0]
        Local $file = $g_IniFiles[$f]
        Local $fileLower = StringLower($file)
        If $fileLower = "settings.ini" Or $fileLower = "ignorelist.ini" Then ContinueLoop

        Local $iniFile = $g_CategoryIniDir & "\" & $file
        Local $sections = IniReadSectionNames($iniFile)
        If Not IsArray($sections) Then ContinueLoop
        For $i = 1 To $sections[0]
            Local $appName = $sections[$i]
            Local $category = IniRead($iniFile, $appName, "Category", "")
            If $category = "" Then $category = "Uncategorized"
            Local $hide = IniRead($iniFile, $appName, "Hide", "0")
            Local $fave = IniRead($iniFile, $appName, "Fave", "0")

            If $fave = "1" Then
                $favourites.Add($appName, $iniFile)
            EndIf

            If Not $categories.Exists($category) Then
                $categories.Add($category, ObjCreate("Scripting.Dictionary"))
            EndIf
            $categories.Item($category).Add($appName, $iniFile)
        Next
    Next

    ; Add Favourites node (alphabetical)
    If $favourites.Count > 0 Then
        Local $favNode = GUICtrlCreateTreeViewItem("Favourites", $g_AppTreeView)
		$g_AppCategoryTreeItems.Add($favNode, "Favourites")
        ; Collect and sort favourite app names
        Local $faveNames[0]
        For $appName In $favourites.Keys
            _ArrayAdd($faveNames, $appName)
        Next
        _ArraySort($faveNames)
        For $i = 0 To UBound($faveNames) - 1
            Local $appName = $faveNames[$i]
            Local $iniFile = $favourites.Item($appName)
            Local $hide = IniRead($iniFile, $appName, "Hide", "0")
            Local $buttonText = IniRead($iniFile, $appName, "ButtonText", $appName)
			Local $label = $buttonText
			If $hide = "1" Then $label = "(H) " & $buttonText
            Local $appNode = GUICtrlCreateTreeViewItem($label, $favNode)
            If Not $g_AppSections.Exists($appName) Then
                $g_AppSections.Add($appName, $iniFile)
            Else
                $g_AppSections.Item($appName) = $iniFile ; Or update, or skip
            EndIf
            $g_AppTreeItemToSection.Add($appNode, $appName)
            $g_AppTreeItemToIni.Add($appNode, $iniFile)
        Next
    EndIf

    ; Add categories (alphabetical)
    Local $catNames[0]
    For $category In $categories.Keys
        _ArrayAdd($catNames, $category)
    Next
    _ArraySort($catNames)
    For $i = 0 To UBound($catNames) - 1
        Local $category = $catNames[$i]
        Local $catNode = GUICtrlCreateTreeViewItem($category, $g_AppTreeView)
        $g_AppCategoryTreeItems.Add($catNode, $category)
        ; Collect and sort app names in this category
        Local $appNames[0]
        For $appName In $categories.Item($category).Keys
            _ArrayAdd($appNames, $appName)
        Next
        _ArraySort($appNames)
        For $j = 0 To UBound($appNames) - 1
            Local $appName = $appNames[$j]
            Local $iniFile = $categories.Item($category).Item($appName)
            Local $hide = IniRead($iniFile, $appName, "Hide", "0")
            Local $fave = IniRead($iniFile, $appName, "Fave", "0")
            ; If Fave=1, already added above
            If $fave = "1" Then ContinueLoop
            Local $buttonText = IniRead($iniFile, $appName, "ButtonText", $appName)
			Local $label = $buttonText
			If $hide = "1" Then $label = "(H) " & $buttonText
            Local $appNode = GUICtrlCreateTreeViewItem($label, $catNode)
            If Not $g_AppSections.Exists($appName) Then
                $g_AppSections.Add($appName, $iniFile)
            Else
                $g_AppSections.Item($appName) = $iniFile ; Or update, or skip
            EndIf
            $g_AppTreeItemToSection.Add($appNode, $appName)
            $g_AppTreeItemToIni.Add($appNode, $iniFile)
        Next
    Next
EndFunc

; Enable all buttons (Ignore, Save)
Func _AppTab_Buttons_Enable()
    GUICtrlSetState($g_IgnoreBtn, $GUI_ENABLE)
    GUICtrlSetState($g_SaveBtn, $GUI_ENABLE)
    GUICtrlSetState($g_RefreshBtn, $GUI_ENABLE)
EndFunc

; Disable all buttons (Ignore, Save)
Func _AppTab_Buttons_Disable()
    GUICtrlSetState($g_IgnoreBtn, $GUI_DISABLE)
    GUICtrlSetState($g_SaveBtn, $GUI_DISABLE)
    GUICtrlSetState($g_RefreshBtn, $GUI_DISABLE)
EndFunc

Func Tab_App_HandleEvents($msg)
    ; Handle buttons and context menu
    Switch $msg
        Case $g_SaveBtn
            If $g_SelectedIni <> "" And $g_SelectedSection <> "" Then
                _AppTab_SaveFields($g_SelectedIni, $g_SelectedSection)
            EndIf
		Case $g_RefreshBtn
			If $g_SelectedIni <> "" And $g_SelectedSection <> "" Then
				_AppTab_SaveFields($g_SelectedIni, $g_SelectedSection)
				_AppTab_LoadIniFilesAndApps()  ; Re-reads INI files to get latest data
				_AppTab_ClearTree2()           ; Just clears tree items (keeps control intact)
				_AppTab_RecreateTree()         ; Repopulates with fresh data
				; Update TreeView item text if ButtonText changed
				Local $newButtonText = GUICtrlRead($g_AppFields("ButtonText"))
				If $newButtonText <> "" Then
					Local $selectedItem = GUICtrlRead($g_AppTreeView)
					If $selectedItem <> 0 Then
						GUICtrlSetData($selectedItem, $newButtonText)
					EndIf
				EndIf
			EndIf
        Case $g_IgnoreBtn
            _AppTab_IgnoreSelected()
        Case $g_AppMenuDeleteItem
            _AppTab_DeleteSelected()
		Case $g_AppMenuEditIniItem
			Local $targetItem = ($g_RightClickedItemID <> 0) ? $g_RightClickedItemID : GUICtrlRead($g_AppTreeView)
			If $targetItem = 0 Or Not $g_AppTreeItemToSection.Exists($targetItem) Then Return
			Local $appName = $g_AppTreeItemToSection.Item($targetItem)
			Local $iniFile = $g_AppTreeItemToIni.Item($targetItem)
			If FileExists($iniFile) Then ShellExecute($iniFile)
        Case $g_AppFields("SymLinkCreate")
            ; Enable/disable symlinks field based on checkbox
            If GUICtrlRead($g_AppFields("SymLinkCreate")) = $GUI_CHECKED Then
                GUICtrlSetState($g_AppsSymlinksEdit, $GUI_ENABLE)
            Else
                GUICtrlSetState($g_AppsSymlinksEdit, $GUI_DISABLE)
            EndIf
        Case $g_AppFields("Sandboxie")
            ; Enable/disable SandboxName field based on checkbox
            If GUICtrlRead($g_AppFields("Sandboxie")) = $GUI_CHECKED Then
                GUICtrlSetState($g_AppFields("SandboxName"), $GUI_ENABLE)
            Else
                GUICtrlSetState($g_AppFields("SandboxName"), $GUI_DISABLE)
            EndIf
    EndSwitch

    Local $selectedItem = GUICtrlRead($g_AppTreeView)
    ; Only act if the selection changed
    If $selectedItem <> $g_AppLastSelectedItem Then
        ; If a category node is selected
        If $selectedItem <> 0 And $g_AppCategoryTreeItems.Exists($selectedItem) Then
            _AppTab_ClearFields()
            _AppTab_DisableFields()
            GUICtrlSetState($g_SaveBtn, $GUI_DISABLE)
            GUICtrlSetState($g_IgnoreBtn, $GUI_DISABLE)
            $g_SelectedIni = ""
            $g_SelectedSection = ""
        ; If last selected was a valid app item and now whitespace/node: disable buttons
        ElseIf $g_AppLastSelectedItem <> 0 And $g_AppTreeItemToSection.Exists($g_AppLastSelectedItem) _
            And ($selectedItem = 0 Or Not $g_AppTreeItemToSection.Exists($selectedItem)) Then
            _AppTab_ClearFields()
            _AppTab_Buttons_Disable()
            $g_SelectedIni = ""
            $g_SelectedSection = ""
        ; If selection changed to a valid app item: enable buttons
        ElseIf $selectedItem <> 0 And $g_AppTreeItemToSection.Exists($selectedItem) Then
            $g_SelectedSection = $g_AppTreeItemToSection.Item($selectedItem)
            $g_SelectedIni = $g_AppTreeItemToIni.Item($selectedItem)
            _AppTab_PopulateFields($g_SelectedIni, $g_SelectedSection)
            _AppTab_EnableFields()
            _AppTab_Buttons_Enable()
        ; If selection changed to whitespace/node, disable buttons
        Else
            _AppTab_ClearFields()
            _AppTab_Buttons_Disable()
            $g_SelectedIni = ""
            $g_SelectedSection = ""
        EndIf
        $g_AppLastSelectedItem = $selectedItem
    EndIf
EndFunc

Func _AppTab_PopulateFields($iniFile, $section)
    For $key In $g_AppFields.Keys()
        ; Populate standard fields
        Local $val = IniRead($iniFile, $section, $key, "")
        If StringInStr("RunAsAdmin SingleInstance Sandboxie SymLinkCreate Fave Hide", $key) Then
            GUICtrlSetState($g_AppFields($key), ($val = "1") ? $GUI_CHECKED : $GUI_UNCHECKED)
        Else
            GUICtrlSetData($g_AppFields($key), $val)
        EndIf
        GUICtrlSetState($g_AppFields($key), $GUI_ENABLE)
    Next
    ; SetEnv value (from SetEnv1 in INI)
    Local $setenv_val = IniRead($iniFile, $section, "SetEnv1", "")
    GUICtrlSetData($g_AppFields("SetEnv"), $setenv_val)
    ; Populate symlinks multiline edit
    Local $symlinks = ""
    Local $idx = 1
    While 1
        Local $val = IniRead($iniFile, $section, "Symlink" & $idx, "")
        If $val = "" Then ExitLoop
        $symlinks &= $val & @CRLF
        $idx += 1
    WEnd
    GUICtrlSetData($g_AppsSymlinksEdit, StringTrimRight($symlinks, 2))
    ; Enable/disable symlinks field according to SymLinkCreate checkbox
    If GUICtrlRead($g_AppFields("SymLinkCreate")) = $GUI_CHECKED Then
        GUICtrlSetState($g_AppsSymlinksEdit, $GUI_ENABLE)
    Else
        GUICtrlSetState($g_AppsSymlinksEdit, $GUI_DISABLE)
    EndIf
    ; Enable/disable SandboxName field according to Sandboxie checkbox
    If GUICtrlRead($g_AppFields("Sandboxie")) = $GUI_CHECKED Then
        GUICtrlSetState($g_AppFields("SandboxName"), $GUI_ENABLE)
    Else
        GUICtrlSetState($g_AppFields("SandboxName"), $GUI_DISABLE)
    EndIf
EndFunc

Func _AppTab_SaveFields($iniFile, $section)
    For $key In $g_AppFields.Keys()
        If StringInStr("RunAsAdmin SingleInstance Sandboxie SymLinkCreate Fave Hide", $key) Then
            Local $v = (GUICtrlRead($g_AppFields($key)) = $GUI_CHECKED) ? "1" : "0"
            IniWrite($iniFile, $section, $key, $v)
        Else
            ; For SetEnv, always write to SetEnv1 as a single line with pipes (never multiline)
            If $key = "SetEnv" Then
                Local $setenv_val = GUICtrlRead($g_AppFields("SetEnv"))
                ; Replace any accidental line breaks with pipes
                $setenv_val = StringReplace($setenv_val, @CRLF, "|")
                $setenv_val = StringReplace($setenv_val, @LF, "|")
                $setenv_val = StringReplace($setenv_val, @CR, "|")
                IniWrite($iniFile, $section, "SetEnv1", $setenv_val)
            Else
                IniWrite($iniFile, $section, $key, GUICtrlRead($g_AppFields($key)))
            EndIf
        EndIf
        GUICtrlSetState($g_AppFields($key), $GUI_DISABLE)
    Next
    ; Save symlinks from multiline edit as Symlink1=..., Symlink2=..., ...
    Local $symlinksText = GUICtrlRead($g_AppsSymlinksEdit)
    Local $symlinksArr = StringSplit(StringStripCR($symlinksText), @LF)
    ; First, delete any old SymlinkX keys
    Local $idx = 1
    While 1
        Local $exists = IniRead($iniFile, $section, "Symlink" & $idx, "")
        If $exists = "" Then ExitLoop
        IniDelete($iniFile, $section, "Symlink" & $idx)
        $idx += 1
    WEnd
    ; Write new symlinks
    If $symlinksArr[0] > 0 Then
        For $i = 1 To $symlinksArr[0]
            If StringStripWS($symlinksArr[$i], 3) <> "" Then
                IniWrite($iniFile, $section, "Symlink" & $i, $symlinksArr[$i])
            EndIf
        Next
    EndIf
    GUICtrlSetState($g_AppsSymlinksEdit, $GUI_DISABLE)
    GUICtrlSetState($g_AppFields("SandboxName"), $GUI_DISABLE)
    RemoveExtraBlankLines($iniFile)
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
    GUICtrlSetData($g_AppFields("SetEnv"), "")
    GUICtrlSetData($g_AppsSymlinksEdit, "")
    ; Enable/disable Symlinks field according to SymLinkCreate checkbox
    If GUICtrlRead($g_AppFields("SymLinkCreate")) = $GUI_CHECKED Then
        GUICtrlSetState($g_AppsSymlinksEdit, $GUI_ENABLE)
    Else
        GUICtrlSetState($g_AppsSymlinksEdit, $GUI_DISABLE)
    EndIf
    ; Enable/disable SandboxName field according to Sandboxie checkbox
    If GUICtrlRead($g_AppFields("Sandboxie")) = $GUI_CHECKED Then
        GUICtrlSetState($g_AppFields("SandboxName"), $GUI_ENABLE)
    Else
        GUICtrlSetState($g_AppFields("SandboxName"), $GUI_DISABLE)
    EndIf
EndFunc

Func _AppTab_EnableFields()
    For $key In $g_AppFields.Keys()
        GUICtrlSetState($g_AppFields($key), $GUI_ENABLE)
    Next
    ; Enable/disable symlinks field according to SymLinkCreate
    If GUICtrlRead($g_AppFields("SymLinkCreate")) = $GUI_CHECKED Then
        GUICtrlSetState($g_AppsSymlinksEdit, $GUI_ENABLE)
    Else
        GUICtrlSetState($g_AppsSymlinksEdit, $GUI_DISABLE)
    EndIf
    ; Enable/disable SandboxName field according to Sandboxie checkbox
    If GUICtrlRead($g_AppFields("Sandboxie")) = $GUI_CHECKED Then
        GUICtrlSetState($g_AppFields("SandboxName"), $GUI_ENABLE)
    Else
        GUICtrlSetState($g_AppFields("SandboxName"), $GUI_DISABLE)
    EndIf
EndFunc

Func _AppTab_DisableFields()
    For $key In $g_AppFields.Keys()
        GUICtrlSetState($g_AppFields($key), $GUI_DISABLE)
    Next
    GUICtrlSetState($g_AppsSymlinksEdit, $GUI_DISABLE)
    GUICtrlSetState($g_AppFields("SandboxName"), $GUI_DISABLE)
EndFunc

Func _AppTab_DeleteSelected($skipConfirm = False)
    ; Use the right-clicked item if we have one, otherwise fall back to selected
    Local $targetItem = ($g_RightClickedItemID <> 0) ? $g_RightClickedItemID : GUICtrlRead($g_AppTreeView)
    
;~     ConsoleWrite("Target item for deletion: " & $targetItem & " (right-clicked: " & $g_RightClickedItemID & ")" & @CRLF)
    
    ; Clear the right-clicked item after using it
    $g_RightClickedItemID = 0
    
    If $targetItem = 0 Then 
        MsgBox(48, "No Selection", "Please select an item first.")
        Return
    EndIf

    ; Get the text of the item to show in confirmation
    Local $hTreeView = GUICtrlGetHandle($g_AppTreeView)
    Local $itemText = _GUICtrlTreeView_GetText($hTreeView, $targetItem)
    
    ; Debug output
;~     ConsoleWrite("About to delete item: '" & $itemText & "' (ID: " & $targetItem & ")" & @CRLF)
    
    ; --- CATEGORY NODE DELETION ---
    If $g_AppCategoryTreeItems.Exists($targetItem) Then
        Local $category = $g_AppCategoryTreeItems.Item($targetItem)
;~         ConsoleWrite("Found in category dictionary: " & $category & @CRLF)
        Local $iniFile = $g_CategoryIniDir & "\" & $category & ".ini"
        ; SAFEGUARD: Never delete Settings.ini or IgnoreList.ini
        If StringInStr(StringLower($iniFile), "settings.ini") Or StringInStr(StringLower($iniFile), "ignorelist.ini") Then Return

        If Not $skipConfirm Then
            Local $answer = MsgBox(1 + 32, "Delete Category", _
                "Are you sure you want to delete the entire category '" & $itemText & "' (" & $category & ")?" & @CRLF & _
                "This will remove the INI file (" & $category & ".ini) and all its apps." & @CRLF & _
                "This cannot be undone.", 0)
            If $answer <> 1 Then Return
        EndIf

        If FileExists($iniFile) Then FileDelete($iniFile)
        GUICtrlDelete($targetItem)
        $g_AppCategoryTreeItems.Remove($targetItem)
        $g_SelectedIni = ""
        $g_SelectedSection = ""
        $g_AppLastSelectedItem = 0
        _AppTab_ClearFields()
        _AppTab_Buttons_Disable()
        Return
    EndIf

    ; --- APP NODE DELETION ---
    If Not $g_AppTreeItemToSection.Exists($targetItem) Then 
;~         ConsoleWrite("Item not found in app dictionaries! ID: " & $targetItem & @CRLF)
        Return
    EndIf

    Local $appName = $g_AppTreeItemToSection.Item($targetItem)
;~     ConsoleWrite("Found in app dictionary: " & $appName & @CRLF)
    Local $iniFile = $g_AppSections.Item($appName)
    ; SAFEGUARD: Never delete from Settings.ini
    If StringInStr(StringLower($iniFile), "settings.ini") Then Return

    If Not $skipConfirm Then
        Local $answer = MsgBox(1 + 32, "Delete App", "Are you sure you want to delete '" & $itemText & "' (" & $appName & ")?" & @CRLF & "This cannot be undone.", 0)
        If $answer <> 1 Then Return
    EndIf

    IniDelete($iniFile, $appName)
	RemoveExtraBlankLines($iniFile)
    GUICtrlDelete($targetItem)
    $g_AppSections.Remove($appName)
    $g_AppTreeItemToSection.Remove($targetItem)
    $g_AppTreeItemToIni.Remove($targetItem)
    _AppTab_ClearFields()
    $g_SelectedIni = ""
    $g_SelectedSection = ""
    $g_AppLastSelectedItem = 0
    _AppTab_Buttons_Disable()
EndFunc

Func _AppTab_IgnoreSelected()
    Local $selectedItem = GUICtrlRead($g_AppTreeView)
    If $selectedItem = 0 Or Not $g_AppTreeItemToSection.Exists($selectedItem) Then Return

    Local $appName = $g_AppTreeItemToSection.Item($selectedItem)
    Local $iniFile = $g_AppSections.Item($appName)
    ; SAFEGUARD: Never ignore from Settings.ini
    If StringInStr(StringLower($iniFile), "settings.ini") Then Return

    ; Get the RunFile (path) to ignore
    Local $ignorePath = IniRead($iniFile, $appName, "RunFile", "")
    If $ignorePath = "" Then Return

    ; --- Custom Ignore Confirmation ---
    Local $answer = MsgBox(1 + 32, "Ignore App", _
        "Confirm adding '" & $appName & "' (" & $ignorePath & ") to the Ignore List and delete from App list." & @CRLF & _
        @CRLF & _
        "You can restore ignored apps via the Ignore tab.", 0)
    If $answer <> 1 Then Return

    ; Add the path to IgnoreList.ini if not already present
    Local $ignoreIni = @ScriptDir & "\App\IgnoreList.ini"
    Local $ignoreArr = FileReadToArray($ignoreIni)
    Local $exists = False
    If IsArray($ignoreArr) Then
        For $i = 0 To UBound($ignoreArr) - 1
            If StringCompare(StringStripWS($ignoreArr[$i], 3), StringStripWS($ignorePath, 3), 1) = 0 Then
                $exists = True
                ExitLoop
            EndIf
        Next
    EndIf
    If Not $exists Then
        Local $hFile = FileOpen($ignoreIni, $FO_APPEND)
        If $hFile <> -1 Then
            FileWriteLine($hFile, $ignorePath)
            FileClose($hFile)
        EndIf
    EndIf

    ; Store the selected item before deletion
    Local $itemToDelete = $selectedItem
    
    ; Clear selection state first
    $g_SelectedIni = ""
    $g_SelectedSection = ""
    $g_AppLastSelectedItem = 0
    _AppTab_ClearFields()
    _AppTab_Buttons_Disable()
    
    ; Delete from INI file
    IniDelete($iniFile, $appName)
	RemoveExtraBlankLines($iniFile)
    
    ; Remove from TreeView and clean up dictionaries
    GUICtrlDelete($itemToDelete)
    $g_AppSections.Remove($appName)
    $g_AppTreeItemToSection.Remove($itemToDelete)
    $g_AppTreeItemToIni.Remove($itemToDelete)
    
    ; Update Ignore tab view
    GUICtrlSetState($g_IgnoreEdit, $GUI_ENABLE)
    _Ignore_EditPopulate()
    GUICtrlSetState($g_IgnoreEdit, $GUI_DISABLE)
EndFunc

Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
    Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
    If $tNMHDR.code <> $NM_RCLICK Then Return $GUI_RUNDEFMSG
    If $tNMHDR.idFrom <> $g_AppTreeView Then Return $GUI_RUNDEFMSG
    
;~     ConsoleWrite("WM_NOTIFY: Right-click detected on TreeView" & @CRLF)
    
    ; Get cursor position and control position
    Local $cursor = GUIGetCursorInfo($hWnd)
    Local $ctrlPos = ControlGetPos($hWnd, "", $g_AppTreeView)

    Local $relX = $cursor[0] - $ctrlPos[0]
    Local $relY = $cursor[1] - $ctrlPos[1]
    
    ; Get the item handle that was hit
    Local $hTreeView = GUICtrlGetHandle($g_AppTreeView)
    Local $hItem = _GUICtrlTreeView_HitTestItem($hTreeView, $relX, $relY)
    If Not $hItem Then 
;~         ConsoleWrite("No item hit" & @CRLF)
        Return $GUI_RUNDEFMSG
    EndIf
    
    Local $itemText = _GUICtrlTreeView_GetText($hTreeView, $hItem)
;~     ConsoleWrite("Hit item text: '" & $itemText & "'" & @CRLF)
    
    ; NEW APPROACH: Enumerate all TreeView items and find the matching one
    $g_RightClickedItemID = 0
    
    ; Get the first (root) item
    Local $hFirstItem = _GUICtrlTreeView_GetFirstItem($hTreeView)
    If $hFirstItem Then
        _FindItemInTree($hTreeView, $hFirstItem, $hItem, $itemText)
    EndIf
    
;~     ConsoleWrite("Final g_RightClickedItemID: " & $g_RightClickedItemID & @CRLF)
    
    If $g_RightClickedItemID = 0 Then
;~         ConsoleWrite("Could not find matching control ID for hit item!" & @CRLF)
    EndIf
    
    Return $GUI_RUNDEFMSG
EndFunc

Func _FindItemInTree($hTreeView, $hCurrentItem, $hTargetItem, $targetText)
    ; Check if current item matches our target
    If $hCurrentItem = $hTargetItem Then
        ; Found the matching item! Now find its control ID
;~         ConsoleWrite("Found matching TreeView item handle!" & @CRLF)
        
        ; Search through our dictionaries to find which control ID corresponds to this item
        ; Try app items first - compare against ButtonText from INI files
        For $itemID In $g_AppTreeItemToSection.Keys
            Local $appName = $g_AppTreeItemToSection.Item($itemID)
            Local $iniFile = $g_AppSections.Item($appName)
            
            ; Read the ButtonText from the INI file (this is what's displayed in TreeView)
            Local $buttonText = IniRead($iniFile, $appName, "ButtonText", $appName)
            Local $hide = IniRead($iniFile, $appName, "Hide", "0")
            Local $label = $buttonText
            If $hide = "1" Then $label = "(H) " & $buttonText
            
;~             ConsoleWrite("Checking app: " & $appName & " (Label: '" & $label & "') (ID: " & $itemID & ")" & @CRLF)
            
            ; Compare the clicked text with the actual label (including hide prefix)
            If $label = $targetText Then
                $g_RightClickedItemID = $itemID
;~                 ConsoleWrite("*** FOUND MATCHING APP BY LABEL! ID: " & $itemID & " ***" & @CRLF)
                Return
            EndIf
        Next
        
        ; Try category items (these might not have ButtonText, so use original logic)
        For $itemID In $g_AppCategoryTreeItems.Keys
            Local $categoryName = $g_AppCategoryTreeItems.Item($itemID)
;~             ConsoleWrite("Checking category: " & $categoryName & " (ID: " & $itemID & ")" & @CRLF)
            
            If $categoryName = $targetText Then
                $g_RightClickedItemID = $itemID
;~                 ConsoleWrite("*** FOUND MATCHING CATEGORY BY NAME! ID: " & $itemID & " ***" & @CRLF)
                Return
            EndIf
        Next
        Return
    EndIf
    
    ; Check child items
    Local $hChild = _GUICtrlTreeView_GetFirstChild($hTreeView, $hCurrentItem)
    While $hChild
        _FindItemInTree($hTreeView, $hChild, $hTargetItem, $targetText)
        If $g_RightClickedItemID <> 0 Then Return ; Found it, stop searching
        $hChild = _GUICtrlTreeView_GetNextSibling($hTreeView, $hChild)
    WEnd
    
    ; Check sibling items
    Local $hSibling = _GUICtrlTreeView_GetNextSibling($hTreeView, $hCurrentItem)
    If $hSibling Then
        _FindItemInTree($hTreeView, $hSibling, $hTargetItem, $targetText)
    EndIf
EndFunc
