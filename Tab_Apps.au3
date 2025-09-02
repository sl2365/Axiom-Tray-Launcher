; Tab_Apps.au3

#include-once
#include <GuiTreeView.au3>
#include <File.au3>

Global $g_CategoryIniDir = @ScriptDir & "\App"
Global $g_AppTreeView, $g_AppFields, $g_IniFiles, $g_AppSections, $g_SelectedIni, $g_SelectedSection
Global $g_SaveBtn, $g_DeleteBtn
Global $g_AppTreeItemToSection = ObjCreate("Scripting.Dictionary") ; maps treeview item ID to section name
Global $g_AppTreeItemToIni = ObjCreate("Scripting.Dictionary")     ; maps treeview item ID to ini file
Global $g_AppLastSelectedItem = 0

Global $g_ViewSwitch, $g_TreeViewMode = "tray" ; "tray" or "folder"

Func Tab_Apps_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("Apps")

    ; Add switch to toggle view style
    Local $x = 290
    $g_ViewSwitch = GUICtrlCreateCheckbox("Show Tree as folder structure", $x, 413, 160, 18)
    GUICtrlSetState($g_ViewSwitch, $GUI_UNCHECKED)

    ; Right side fields container
    $g_AppFields = ObjCreate("Scripting.Dictionary")
    Local $fields = ["ButtonText", "RunFile", "RunAsAdmin", "WorkDir", "Arguments", "SingleInstance", "Sandboxie", "SandboxName", "Category", "SymLinkCreate", "SymLink1", "Fave", "Hide"]
    For $i = 0 To UBound($fields) - 1
        GUICtrlCreateLabel($fields[$i] & ":", $x, 60 + $i * 27, 70, 18)
        If StringInStr("RunAsAdmin SingleInstance Sandboxie SymLinkCreate Fave Hide", $fields[$i]) Then
            $g_AppFields($fields[$i]) = GUICtrlCreateCheckbox("", $x + 90, 56 + $i * 27, 18, 18)
        Else
            $g_AppFields($fields[$i]) = GUICtrlCreateInput("", $x + 90, 56 + $i * 27, 220, 18)
        EndIf
    Next

    $g_AppTreeView = GUICtrlCreateTreeView($listviewX, $listviewY, $x + 90, $guiH-155, BitOR($TVS_HASBUTTONS, $TVS_LINESATROOT, $TVS_SHOWSELALWAYS))
    $g_DeleteBtn = GUICtrlCreateButton("Delete", $x + 180, 410, $btnW, $btnH)
    $g_SaveBtn = GUICtrlCreateButton("Save", $x + 250, 410, $btnW, $btnH)

    _AppTab_LoadIniFilesAndApps()
EndFunc

Func _AppTab_LoadIniFilesAndApps()
    Global $g_AppSections
    $g_AppSections = ObjCreate("Scripting.Dictionary")
    $g_AppTreeItemToSection = ObjCreate("Scripting.Dictionary")
    $g_AppTreeItemToIni = ObjCreate("Scripting.Dictionary")
    $g_IniFiles = _FileListToArray($g_CategoryIniDir, "*.ini", 1)
    If @error Or Not IsArray($g_IniFiles) Then Return

    ; Clear previous tree nodes
    GUICtrlDelete($g_AppTreeView)
    $g_AppTreeView = GUICtrlCreateTreeView($listviewX, $listviewY, 250, $guiH-155, BitOR($TVS_HASBUTTONS, $TVS_LINESATROOT, $TVS_SHOWSELALWAYS))

    Local $categoryNodes = ObjCreate("Scripting.Dictionary")

    If $g_TreeViewMode = "tray" Then
        ; TrayMenu logic: group by Category key, show all apps, mark hidden, add "Favourites" group

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

        ; Add Favourites node (British English)
        If $favourites.Count > 0 Then
            Local $favNode = GUICtrlCreateTreeViewItem("Favourites", $g_AppTreeView)
            For $appName In $favourites.Keys
                Local $iniFile = $favourites.Item($appName)
                Local $hide = IniRead($iniFile, $appName, "Hide", "0")
                Local $label = $appName
                If $hide = "1" Then $label = "(H) " & $appName
                Local $appNode = GUICtrlCreateTreeViewItem($label, $favNode)
                $g_AppSections.Add($appName, $iniFile)
                $g_AppTreeItemToSection.Add($appNode, $appName)
                $g_AppTreeItemToIni.Add($appNode, $iniFile)
            Next
        EndIf

        ; Add categories
        For $category In $categories.Keys
            Local $catNode = GUICtrlCreateTreeViewItem($category, $g_AppTreeView)
            For $appName In $categories.Item($category).Keys
                Local $iniFile = $categories.Item($category).Item($appName)
                Local $hide = IniRead($iniFile, $appName, "Hide", "0")
                Local $fave = IniRead($iniFile, $appName, "Fave", "0")
                ; If Fave=1, already added above
                If $fave = "1" Then ContinueLoop
                Local $label = $appName
                If $hide = "1" Then $label = "(H) " & $appName
                Local $appNode = GUICtrlCreateTreeViewItem($label, $catNode)
                $g_AppSections.Add($appName, $iniFile)
                $g_AppTreeItemToSection.Add($appNode, $appName)
                $g_AppTreeItemToIni.Add($appNode, $iniFile)
            Next
        Next

    Else
        ; Folder structure logic: group by INI file name, show all apps, mark hidden
        For $f = 1 To $g_IniFiles[0]
            Local $file = $g_IniFiles[$f]
            Local $fileLower = StringLower($file)
            If $fileLower = "settings.ini" Or $fileLower = "ignorelist.ini" Then ContinueLoop

            Local $iniFile = $g_CategoryIniDir & "\" & $file
            Local $catName = StringTrimRight($file, 4) ; Remove ".ini"
            Local $sections = IniReadSectionNames($iniFile)
            If Not IsArray($sections) Or $sections[0] = 0 Then ContinueLoop
            Local $catNode = GUICtrlCreateTreeViewItem($catName, $g_AppTreeView)
            For $i = 1 To $sections[0]
                Local $appName = $sections[$i]
                Local $hide = IniRead($iniFile, $appName, "Hide", "0")
                Local $label = $appName
                If $hide = "1" Then $label = "(H) " & $appName
                Local $appNode = GUICtrlCreateTreeViewItem($label, $catNode)
                $g_AppSections.Add($appName, $iniFile)
                $g_AppTreeItemToSection.Add($appNode, $appName)
                $g_AppTreeItemToIni.Add($appNode, $iniFile)
            Next
        Next
    EndIf
EndFunc

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

    ; Poll TreeView for selection change (single-click works)
    Local $selectedItem = GUICtrlRead($g_AppTreeView)
    If $selectedItem <> $g_AppLastSelectedItem Then
        $g_AppLastSelectedItem = $selectedItem
        If $selectedItem <> 0 And $g_AppTreeItemToSection.Exists($selectedItem) Then
            $g_SelectedSection = $g_AppTreeItemToSection.Item($selectedItem)
            $g_SelectedIni = $g_AppTreeItemToIni.Item($selectedItem)
            _AppTab_PopulateFields($g_SelectedIni, $g_SelectedSection)
            _AppTab_EnableFields()
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
    Local $selectedItem = GUICtrlRead($g_AppTreeView)
    If $selectedItem = 0 Or Not $g_AppTreeItemToSection.Exists($selectedItem) Then Return

    Local $appName = $g_AppTreeItemToSection.Item($selectedItem)
    Local $iniFile = $g_AppSections.Item($appName)
    ; SAFEGUARD: Never delete from Settings.ini
    If StringInStr(StringLower($iniFile), "settings.ini") Then Return

    Local $answer = MsgBox(1 + 32, "Delete App", "Are you sure you want to delete '" & $appName & "'?" & @CRLF & "This cannot be undone.", 0)
    If $answer = 1 Then ; OK pressed
        IniDelete($iniFile, $appName)
        GUICtrlDelete($selectedItem)
        $g_AppSections.Remove($appName)
        $g_AppTreeItemToSection.Remove($selectedItem)
        $g_AppTreeItemToIni.Remove($selectedItem)
        _AppTab_ClearFields()
        $g_SelectedIni = ""
        $g_SelectedSection = ""
        $g_AppLastSelectedItem = 0
    EndIf
EndFunc
