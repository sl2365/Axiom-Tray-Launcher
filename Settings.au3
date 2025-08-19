#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <GuiListView.au3>
#include <File.au3>
#include <WinAPIFiles.au3>
#include <GuiTab.au3>
#include <GuiEdit.au3>
#include <GuiButton.au3>
#include <GuiComboBox.au3>
#include <GuiListBox.au3>
#include <GuiSlider.au3>
#include <GuiMenu.au3>
#include <Misc.au3>

Global $__SettingsGUIStandaloneRun = false
Global $g_hGUI, $g_hTab
Global $g_aCheckBoxCtrlIDs[0], $g_aCheckBoxNames[0]
Global $g_hSandBoxieInput, $hSandBoxieBrowse
Global $hScanLV, $hScanLV_Handle, $btnAddScan, $btnEditScan, $btnDeleteScan, $btnDebugLV, $btnScan
Global $g_aScanPaths[0], $g_aScanDepths[0]
Global $g_AboutLinkURL = "https://www.portablefreeware.com/forums/viewtopic.php?p=109314#p109314"
Global $g_AboutLinkCtrl
Global $btnOK, $btnCancel

Func _DebugListView()
    Local $selText = GUICtrlRead($hScanLV)
    MsgBox(64, "DEBUG", "GUICtrlRead result: " & $selText)
    If $selText <> "" Then
        Local $parts = StringSplit($selText, "|")
        If $parts[0] >= 3 Then
            MsgBox(64, "DEBUG", "Selected index: " & (Number($parts[1]) - 1) & @CRLF & _
                "Path: " & $parts[2] & @CRLF & "Depth: " & $parts[3])
        EndIf
    EndIf
EndFunc

Func ShowSettingsGUI()
    $g_hGUI = GUICreate("Settings", 700, 420, -1, -1, $WS_SYSMENU)
    $g_hTab = GUICtrlCreateTab(10, 10, 680, 380)
    
    ; --- Global Tab ---
    GUICtrlCreateTabItem("Global")
    SetupGlobalTab()
    ; --- Buttons Tab (NEW TAB #2) ---
    GUICtrlCreateTabItem("Buttons")
    SetupButtonsTab()
    ; --- About Tab (now TAB #3) ---
    GUICtrlCreateTabItem("About")
    SetupAboutTab()
    GUICtrlCreateTabItem("") ; End tab items
    ; OK and Cancel buttons - Always visible, not part of tabs
    $btnOK     = GUICtrlCreateButton("OK", 510, 350, 80, 30)
    $btnCancel = GUICtrlCreateButton("Cancel", 600, 350, 80, 30)
    GUISetState(@SW_SHOW, $g_hGUI)

    While 1
        Local $msg = GUIGetMsg()
        Switch $msg
            Case $GUI_EVENT_CLOSE
                ExitLoop
            ; Add any button handlers for the Buttons tab here, e.g.:
            ; Case $btnMyButton
            ;     ; Do something
            Case $btnAddScan
                _AddScanPath()
            Case $btnEditScan
                _EditScanPath()
            Case $btnDeleteScan
                _DeleteScanPath()
            Case $btnDebugLV
                _DebugListView()
			Case $btnScan
				ScanAppsFolders()
            Case $hSandBoxieBrowse
				Local $exePath = FileOpenDialog("Select SandMan.exe (Sandboxie)", @ScriptDir, "Executable (*.exe)", 1)
				If @error Or $exePath = "" Then ContinueLoop
				GUICtrlSetData($g_hSandBoxieInput, $exePath)
            Case $g_AboutLinkCtrl
                ShellExecute($g_AboutLinkURL)
            Case $btnOK
                SaveGlobalTab()
                SaveScanPathsToINI()
                ExitLoop
            Case $btnCancel
                ExitLoop
        EndSwitch
    WEnd
    GUIDelete($g_hGUI)
EndFunc

Func SetupGlobalTab()
    Local $iY = 50
    GUICtrlCreateLabel("Scan Paths", 20, $iY, 100, 20)
    $hScanLV = GUICtrlCreateListView("#|Path|Depth", 20, $iY + 20, 600, 180, $LVS_SHOWSELALWAYS + $LVS_SINGLESEL)
    $hScanLV_Handle = GUICtrlGetHandle($hScanLV)
	_GUICtrlListView_SetColumnWidth($hScanLV_Handle, 0, 40)
    _GUICtrlListView_SetColumnWidth($hScanLV_Handle, 1, 480) ; Path column enlarged!
    _GUICtrlListView_SetColumnWidth($hScanLV_Handle, 2, 50)
    $btnAddScan    = GUICtrlCreateButton("Add",      630, $iY + 25, 50, 25)
    $btnEditScan   = GUICtrlCreateButton("Edit",     630, $iY + 60, 50, 25)
    $btnDeleteScan = GUICtrlCreateButton("Delete",   630, $iY + 95, 50, 25)
    $btnDebugLV    = GUICtrlCreateButton("Debug LV", 630, $iY + 130, 60, 25)
	$btnScan       = GUICtrlCreateButton("Scan",     630, $iY + 170, 50, 25)
	GUICtrlSetState($btnDebugLV, $GUI_HIDE)

    Local $iY2 = $iY + 210
    Local $aSettings = IniReadSection(@ScriptDir & "\App\Settings.ini", "GLOBAL")
    If @error Then Return
    Global $g_aCheckBoxCtrlIDs[0], $g_aCheckBoxNames[0]
    For $i = 1 To $aSettings[0][0]
        Local $key = $aSettings[$i][0]
        Local $val = $aSettings[$i][1]
        If StringRegExp($key, "^Scan\d+$") Or StringRegExp($key, "^Scan\d+Depth$") Then
            ContinueLoop
        EndIf
        If $val = "1" Or $val = "0" Then
            Local $id = GUICtrlCreateCheckbox($key, 30, $iY2, 200, 20)
            GUICtrlSetState($id, ($val = "1") ? $GUI_CHECKED : $GUI_UNCHECKED)
            _ArrayAdd($g_aCheckBoxCtrlIDs, $id)
            _ArrayAdd($g_aCheckBoxNames, $key)
            $iY2 += 30
        EndIf
    Next

    GUICtrlCreateLabel("Sandboxie Executable Path (SandMan.exe):", 30, $iY2, 250, 20)
    Local $sandboxieValue = IniRead(@ScriptDir & "\App\Settings.ini", "GLOBAL", "SandboxiePath", "")
    $g_hSandBoxieInput = GUICtrlCreateInput($sandboxieValue, 30, $iY2 + 25, 250, 20)
    $hSandBoxieBrowse = GUICtrlCreateButton("...", 250, $iY2 -4, 30, 22)

    _LoadScanPathsFromINI()
    _RefreshScanListView()
EndFunc

Func SaveGlobalTab()
    For $i = 0 To UBound($g_aCheckBoxCtrlIDs) - 1
        Local $state = GUICtrlRead($g_aCheckBoxCtrlIDs[$i])
        Local $val = ($state = $GUI_CHECKED) ? "1" : "0"
        IniWrite(@ScriptDir & "\App\Settings.ini", "GLOBAL", $g_aCheckBoxNames[$i], $val)
    Next
    Local $sandboxiePath = GUICtrlRead($g_hSandBoxieInput)
    IniWrite(@ScriptDir & "\App\Settings.ini", "GLOBAL", "SandboxiePath", $sandboxiePath)
EndFunc

Func _LoadScanPathsFromINI()
    Global $g_aScanPaths[0], $g_aScanDepths[0]
    Local $i = 1
    While True
        Local $path = IniRead(@ScriptDir & "\App\Settings.ini", "GLOBAL", "Scan" & $i, "")
        If $path = "" Then ExitLoop
        Local $depth = IniRead(@ScriptDir & "\App\Settings.ini", "GLOBAL", "Scan" & $i & "Depth", "1")
        _ArrayAdd($g_aScanPaths, $path)
        _ArrayAdd($g_aScanDepths, $depth)
        $i += 1
    WEnd
EndFunc

; --- New combined dialog for path/depth ---
Func ShowScanPathDialog($sPath = "", $sDepth = "1", $bEdit = False)
    Local $title = $bEdit ? "Edit Scan Path" : "Add Scan Path"
    Local $hDlg = GUICreate($title, 365, 150, -1, -1, $WS_SYSMENU)
    GUICtrlCreateLabel("Scan Path:", 10, 14, 70, 22)
    Local $inpPath = GUICtrlCreateInput($sPath, 80, 10, 220, 20)
    Local $btnBrowse = GUICtrlCreateButton("...", 310, 9, 40, 22)
    GUICtrlCreateLabel("Scan Depth:", 10, 50, 70, 22)
    Local $cboDepth = GUICtrlCreateCombo("", 80, 46, 60, 22)
	GUICtrlCreateLabel("1 = Scan path folder, 2 = Sub folder, etc.", 150, 50, 200, 22)
    For $i = 1 To 10
        GUICtrlSetData($cboDepth, $i)
    Next
    GUICtrlSetData($cboDepth, $sDepth)
    Local $btnOK = GUICtrlCreateButton("OK", 180, 85, 80, 30)
    Local $btnCancel = GUICtrlCreateButton("Cancel", 270, 85, 80, 30)
    GUISetState(@SW_SHOW, $hDlg)

    Local $result[2] = ["", ""]
    While 1
        Local $msg = GUIGetMsg()
        Select
            Case $msg = $GUI_EVENT_CLOSE Or $msg = $btnCancel
                ExitLoop
            Case $msg = $btnBrowse
                Local $selectedFolder = FileSelectFolder("Select folder for Scan Path", "", 1)
                If Not @error And $selectedFolder <> "" Then
                    GUICtrlSetData($inpPath, $selectedFolder)
                EndIf
            Case $msg = $btnOK
                $result[0] = GUICtrlRead($inpPath)
                Local $depth = GUICtrlRead($cboDepth)
                ; Auto-correct depth between 1 and 10
                If Not StringIsInt($depth) Then
                    $depth = 1
                ElseIf $depth < 1 Then
                    $depth = 1
                ElseIf $depth > 10 Then
                    $depth = 10
                EndIf
                $result[1] = $depth
                ExitLoop
        EndSelect
    WEnd
    GUIDelete($hDlg)
    Return $result
EndFunc

Func _AddScanPath()
    Local $vals = ShowScanPathDialog("", "1", False)
    If $vals[0] = "" Then Return
    If $vals[1] = "" Then $vals[1] = "1"
    _ArrayAdd($g_aScanPaths, $vals[0])
    _ArrayAdd($g_aScanDepths, $vals[1])
    _RefreshScanListView()
EndFunc

Func _GetSelectedIndex()
    Local $count = _GUICtrlListView_GetItemCount($hScanLV_Handle)
    For $i = 0 To $count - 1
        If _GUICtrlListView_GetItemSelected($hScanLV_Handle, $i) Then
            Return $i
        EndIf
    Next
    Return -1
EndFunc

Func _EditScanPath()
    Local $selIndex = _GetSelectedIndex()
    If $selIndex = -1 Then
        MsgBox(16, "Edit Scan Path", "Please select a scan path to edit.")
        Return
    EndIf
    Local $currPath = $g_aScanPaths[$selIndex]
    Local $currDepth = $g_aScanDepths[$selIndex]
    Local $vals = ShowScanPathDialog($currPath, $currDepth, True)
    If $vals[0] = "" Then Return
    If $vals[1] = "" Then $vals[1] = "1"
    $g_aScanPaths[$selIndex] = $vals[0]
    $g_aScanDepths[$selIndex] = $vals[1]
    _RefreshScanListView()
EndFunc

Func _DeleteScanPath()
    Local $selIndex = _GetSelectedIndex()
    If $selIndex = -1 Then
        MsgBox(16, "Delete Scan Path", "Please select a scan path to delete.")
        Return
    EndIf
    _ArrayDelete($g_aScanPaths, $selIndex)
    _ArrayDelete($g_aScanDepths, $selIndex)
    _RefreshScanListView()
EndFunc

Func _RefreshScanListView()
    _GUICtrlListView_DeleteAllItems($hScanLV_Handle)
    For $i = 0 To UBound($g_aScanPaths) - 1
        GUICtrlCreateListViewItem(StringFormat("%d|%s|%s", $i+1, $g_aScanPaths[$i], $g_aScanDepths[$i]), $hScanLV)
    Next
EndFunc

Func SaveScanPathsToINI()
    Local $i = 1
    While True
        Local $old = IniRead(@ScriptDir & "\App\Settings.ini", "GLOBAL", "Scan" & $i, "")
        If $old = "" Then ExitLoop
        IniDelete(@ScriptDir & "\App\Settings.ini", "GLOBAL", "Scan" & $i)
        IniDelete(@ScriptDir & "\App\Settings.ini", "GLOBAL", "Scan" & $i & "Depth")
        $i += 1
    WEnd
    For $j = 0 To UBound($g_aScanPaths) - 1
        IniWrite(@ScriptDir & "\App\Settings.ini", "GLOBAL", "Scan" & ($j + 1), $g_aScanPaths[$j])
        IniWrite(@ScriptDir & "\App\Settings.ini", "GLOBAL", "Scan" & ($j + 1) & "Depth", $g_aScanDepths[$j])
    Next
EndFunc

Func SetupButtonsTab()
    Local $iY = 50
    GUICtrlCreateLabel("Button Settings go here!", 30, $iY, 300, 20)
    ; Example: add more controls here as needed
    ; Local $btnExample = GUICtrlCreateButton("Example", 30, $iY + 30, 80, 25)
EndFunc

Func SetupAboutTab()
    Local $startY = 50 ; Move About tab content down so it doesn't overlap tab header!
    GUICtrlCreatePic(@ScriptDir & "\App\AxiomTrayJPG.jpg", 10, $startY, 64, 64)
    Local $lblTitle = GUICtrlCreateLabel("Axiom Tray Launcher", 90, $startY, 250, 30)
    GUICtrlSetFont($lblTitle, 16, 800, 0, "Segoe UI")
    Local $aAboutLabels[6]
    $aAboutLabels[0] = GUICtrlCreateLabel("Version 1.0.0.169", 90, $startY + 35, 250, 20)
    $aAboutLabels[1] = GUICtrlCreateLabel("Author: sl23", 90, $startY + 60, 250, 20)
    $aAboutLabels[2] = GUICtrlCreateLabel("Build date: 2025-08-17", 90, $startY + 85, 250, 20)
    $aAboutLabels[3] = GUICtrlCreateLabel("A tray menu to launch your portable apps.", 90, $startY + 110, 250, 20)
    $aAboutLabels[4] = GUICtrlCreateLabel("Website: " & $g_AboutLinkURL, 90, $startY + 135, 550, 20)
    $aAboutLabels[5] = GUICtrlCreateButton("Visit Website", 145, $startY + 160, 180, 20)
    $g_AboutLinkCtrl = $aAboutLabels[5] ; set global button ID
    For $i = 0 To UBound($aAboutLabels) - 1
        GUICtrlSetFont($aAboutLabels[$i], 10)
    Next
EndFunc

If $__SettingsGUIStandaloneRun Then ShowSettingsGUI()
