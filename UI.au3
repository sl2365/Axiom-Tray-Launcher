#include-once
#include <TrayConstants.au3>
#include <Array.au3>
#include <StaticConstants.au3>

; ----------------- Minimal _MySingleton (no Misc.au3 dependency) -----------------
; Returns a non-zero handle if this is the first instance; returns 0 if another instance exists.
Global $g_hSingleton = 0

Func _MySingleton($sOccurrenceName, $iFlag = 0)
    Local $aCreate = DllCall("kernel32.dll", "handle", "CreateMutexW", "ptr", 0, "bool", True, "wstr", $sOccurrenceName)
    If @error Or Not IsArray($aCreate) Then Return SetError(1, 0, 0)
    $g_hSingleton = $aCreate[0]

    Local $aGLE = DllCall("kernel32.dll", "dword", "GetLastError")
    If @error Or Not IsArray($aGLE) Then Return SetError(2, 0, 0)

    If $aGLE[0] = 183 Then
        DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $g_hSingleton)
        $g_hSingleton = 0
        Return 0
    EndIf

    OnAutoItExitRegister("__MySingleton_CloseMutex")
    Return $g_hSingleton
EndFunc

Func __MySingleton_CloseMutex()
    If $g_hSingleton <> 0 Then
        DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $g_hSingleton)
        $g_hSingleton = 0
    EndIf
EndFunc

; ----------------- In-memory button catalog -----------------
Global $g_ButtonCatalog[0][2]

Func Buttons_Reset()
    ReDim $g_ButtonCatalog[0][2]
EndFunc

Func Buttons_Add($label, $handler)
    Local $n = UBound($g_ButtonCatalog)
    ReDim $g_ButtonCatalog[$n + 1][2]
    $g_ButtonCatalog[$n][0] = $label
    $g_ButtonCatalog[$n][1] = $handler
    Return $n
EndFunc

Func Buttons_Count()
    Return UBound($g_ButtonCatalog)
EndFunc

Func Buttons_Label($i)
    If $i < 0 Or $i >= UBound($g_ButtonCatalog) Then Return ""
    Return $g_ButtonCatalog[$i][0]
EndFunc

Func Buttons_Update()
    Local $versionInfo = InetRead($g_UpdateInfoURL, 1)
    If @error Or $versionInfo = "" Then
        MsgBox(16, "Update", "Unable to check for updates.")
        Return
    EndIf

    Local $infoStr = BinaryToString($versionInfo)
    Local $parts = StringSplit($infoStr, "|")
    If $parts[0] < 2 Then
        MsgBox(16, "Update", "Invalid update info format.")
        Return
    EndIf

    Local $latestVersion = $parts[1]
    Local $downloadURL = $parts[2]
    If $latestVersion = $g_CurrentVersion Then
        MsgBox(64, "Update", "You have the latest version.")
        Return
    EndIf

    Local $answer = MsgBox(65, "Update", "New version " & $latestVersion & " available. Update now?")
    If $answer <> 1 Then Return ; Cancel

    Local $tempDir = @TempDir & "\AxiomTrayLauncherUpdate"
    DirCreate($tempDir)
    Local $zipPath = $tempDir & "\update.zip"
    InetGet($downloadURL, $zipPath, 1, 1)
;~     While @InetGetActive
;~         Sleep(500)
;~     WEnd

    Local $extractCmd = '"' & @ScriptDir & '\7za.exe" x "' & $zipPath & '" -o"' & $tempDir & '" -y'
    RunWait($extractCmd, @ScriptDir, @SW_HIDE)

    Local $newExe = $tempDir & "\" & $g_UpdateExeName
    If Not FileExists($newExe) Then
        MsgBox(16, "Update", "Update failed: new EXE not found.")
        Return
    EndIf

    ShellExecute($newExe, "/updatehelper " & @ScriptFullPath, @ScriptDir)
    Exit
EndFunc

Func _CheckForUpdateOnStartup()
    Local $versionInfo = InetRead($g_UpdateInfoURL, 1)
    If @error Or $versionInfo = "" Then Return

    Local $infoStr = BinaryToString($versionInfo)
    Local $parts = StringSplit($infoStr, "|")
    If $parts[0] < 2 Then Return

    Local $latestVersion = $parts[1]
    Local $downloadURL = $parts[2]

    If $latestVersion = $g_CurrentVersion Then Return

    Local $answer = MsgBox(65, "Update", "New version " & $latestVersion & " is available. Update now?")
    If $answer <> 1 Then Return

    Buttons_Update()
EndFunc

; Robust invoker: try IsFunc/Call first, then fall back to Execute()
Func Buttons_Invoke($i)
    If $i < 0 Or $i >= UBound($g_ButtonCatalog) Then Return SetError(1, 0, False)
    Local $fn = $g_ButtonCatalog[$i][1]
    If $fn = "" Then Return SetError(2, 0, False)

    ; Try normal route
    If IsFunc($fn) Then
        Call($fn)
        Return True
    EndIf

    ; Fallback: Execute allows calling even when IsFunc oddly returns 0 on some setups
    Local $res = Execute($fn & "()")
    If @error = 0 Then Return True

    MsgBox(16, "AxiomTrayLauncher", "Handler not found: " & $fn)
    Return SetError(3, 0, False)
EndFunc

Func Buttons_Init()
    Buttons_Reset()
    Buttons_Add("Reload", "Buttons_Reload")
    Buttons_Add("Generate Links", "Buttons_GenerateLinks")
    Buttons_Add("Scan", "Buttons_Scan")
    Buttons_Add("Close SandBoxie", "Buttons_CloseSandboxie")
    Buttons_Add("Cmd Syntax", "Buttons_ShowCmdSyntax")
    Buttons_Add("Env Var's", "Buttons_ShowEnvVars")
    Buttons_Add("Run As Admin", "Buttons_RunAsAdmin")
    Buttons_Add("Settings", "Buttons_OpenSettings")
	Buttons_Add("Update", "Buttons_Update")
EndFunc

; ----------------- CLI handling (/skiptobutton) -----------------
Func _CLS_ParseRunNumber($text, $keyword)
    Local $t = StringLower(StringStripWS($text, 3))
    Local $pos = StringInStr($t, "/" & $keyword)
    If $pos = 0 Then Return 0
    Local $rest = StringStripWS(StringMid($t, $pos + 1 + StringLen($keyword)), 3)
    If $rest = "" Then Return 0
    If StringLeft($rest, 1) = "=" Or StringLeft($rest, 1) = ":" Then $rest = StringStripWS(StringMid($rest, 2), 3)
    If StringIsInt($rest) Then Return Number($rest)
    Return 0
EndFunc

Func Buttons_HandleSecondaryInstance()
;~ 	MsgBox(0, "DEBUG", "Buttons_HandleSecondaryInstance called")
    If _MySingleton("AxiomTrayLauncher", 1) <> 0 Then Return False ; we are primary
    Local $id = 0
    For $i = 1 To $CMDLINE[0]
        If StringLeft(StringLower($CMDLINE[$i]), 13) = "/skiptobutton" Then
            Local $rest = StringStripWS(StringMid($CMDLINE[$i], 14), 3)
            If $rest = "" And $i < $CMDLINE[0] Then $rest = $CMDLINE[$i + 1]
            If StringLeft($rest, 1) = "=" Or StringLeft($rest, 1) = ":" Then $rest = StringStripWS(StringMid($rest, 2), 3)
            If StringIsInt($rest) Then
                $id = Number($rest)
                ExitLoop
            EndIf
        EndIf
    Next
    If $id <= 0 Then
        MsgBox(48, "AxiomTrayLauncher (secondary)", "No /skiptobutton N argument found.")
        Return True
    EndIf
    If Not IPC_PostRunBtn($id) Then
        MsgBox(16, "AxiomTrayLauncher", "Main instance not found or message failed.")
    EndIf
    Return True
EndFunc

Func Buttons_HandleStartupArgs()
    For $i = 1 To $CMDLINE[0]
        If StringLeft(StringLower($CMDLINE[$i]), 13) = "/skiptobutton" Then
            Local $cmd = ""
            For $k = $i To $CMDLINE[0]
                If $cmd <> "" Then $cmd &= " "
                $cmd &= $CMDLINE[$k]
            Next
            Local $n = _CLS_ParseRunNumber($cmd, "skiptobutton")
            If $n > 0 Then RunButtonById($n)
            ExitLoop
        EndIf
    Next
EndFunc

; ----------------- Shortcuts (stable IDs) -----------------
Func GenerateShortcutsForButtons()
    Local $shortcutFolder = @ScriptDir & "\Shortcuts"
    If Not FileExists($shortcutFolder) Then DirCreate($shortcutFolder)
    Local $mainExe = @ScriptFullPath
    Local $count = 0
    Local $ub = UBound($appsData) - 1
    If $ub < 0 Then
        MsgBox(48, "Shortcuts", "No apps to create shortcuts for.")
        Return
    EndIf
    For $i = 0 To $ub
        If Not _IsMainButton($i) Then ContinueLoop
        If $appsData[$i][$APP_ID] <= 0 Then ContinueLoop
        Local $btnText = StringRegExpReplace($appsData[$i][$APP_NAME], '[\\/:*?"<>|]', "_")
        Local $shortcutPath = $shortcutFolder & "\" & $btnText & ".lnk"
        Local $args = "/skiptobutton " & $appsData[$i][$APP_ID]
        ; Silently overwrite existing shortcut
        If FileExists($shortcutPath) Then
            FileDelete($shortcutPath)
        EndIf
        FileCreateShortcut($mainExe, $shortcutPath, @ScriptDir, $args, "Launch '" & $btnText & "' from the tray menu", $mainExe)
        $count += 1
    Next
    MsgBox(64, "Shortcuts", "Created/updated " & $count & " shortcuts in 'Shortcuts' folder.")
EndFunc

; ----------------- Default button handlers -----------------
Func Buttons_Reload()
;~     EnsureIniDefaults()
    ReadSettings()
    ReadApps()
	LoadScannedApps()
    LoadPortableAppsList()
    TrayUI_Destroy()
    TrayUI_BuildTrayMenu()
    If $DEBUG_IPC Then MsgBox(64, "AxiomTrayLauncher", "Configuration reloaded.")
EndFunc

Func Buttons_Scan()
    ScanAppsFolders()
	LoadScannedApps()
    LoadPortableAppsList()
    TrayUI_Destroy()
    TrayUI_BuildTrayMenu()
EndFunc

Func Buttons_OpenSettings()
	ShowSettingsGUI()
;~     ShellExecute($INI_FILE)
EndFunc

Func Buttons_CloseSandboxie()
    If $SandboxieConfigured Then _StopSandboxieService()
EndFunc

Func Buttons_ShowCmdSyntax()
    MsgBox(64, "Help", "Command: /skiptobutton N (stable BUTTON ID)")
EndFunc

Func Buttons_ShowEnvVars()
    ShowEnvVarListView()
EndFunc

Func Buttons_GenerateLinks()
    GenerateShortcutsForButtons()
EndFunc

Func Buttons_RunAsAdmin()
    RestartAsAdmin()
EndFunc

; ----------------- Tray UI -----------------
Global $trayTitleMenu = 0
Global $trayButtonItems[1] = [0]
Global $traySeparator = 0
Global $trayCatMenus[1] = [0]
Global $trayAppItems[1] = [0]
Global $trayFaveSeparator = 0
Global $trayPreExitSeparator = 0
Global $trayExit = 0

Func TrayUI_Destroy()
    If IsArray($trayButtonItems) Then
        For $i = 0 To UBound($trayButtonItems) - 1
            If $trayButtonItems[$i] Then TrayItemDelete($trayButtonItems[$i])
        Next
    EndIf
    If IsArray($trayCatMenus) Then
        For $i = 0 To UBound($trayCatMenus) - 1
            If $trayCatMenus[$i] Then TrayItemDelete($trayCatMenus[$i])
        Next
    EndIf
    If IsArray($trayAppItems) Then
        For $i = 0 To UBound($trayAppItems) - 1
            If $trayAppItems[$i] Then TrayItemDelete($trayAppItems[$i])
        Next
    EndIf
    If $trayTitleMenu Then TrayItemDelete($trayTitleMenu)
    If $traySeparator Then TrayItemDelete($traySeparator)
    If $trayFaveSeparator Then TrayItemDelete($trayFaveSeparator)
    If $trayPreExitSeparator Then TrayItemDelete($trayPreExitSeparator)
    If $trayExit Then TrayItemDelete($trayExit)
EndFunc

Func TrayUI_BuildTrayMenu()
    Local $TitleText = IniRead($INI_FILE, "GLOBAL", "TitleText", "AxiomTrayLauncher")

    ; Gather unique categories from appsData
    Local $categories[0]
    For $i = 0 To UBound($appsData) - 1
        If $appsData[$i][$APP_HIDE] = "1" Then ContinueLoop
        Local $category = StringStripWS($appsData[$i][$APP_CAT], 3)
        If $category = "" Then
            $appsData[$i][$APP_CAT] = "Uncategorised"
            $category = "Uncategorised"
        EndIf
        If StringLower($category) = "fave" Then ContinueLoop
        Local $found = False
        For $j = 0 To UBound($categories) - 1
            If StringLower($categories[$j]) = StringLower($category) Then
                $found = True
                ExitLoop
            EndIf
        Next
        If Not $found Then
            ReDim $categories[UBound($categories) + 1]
            $categories[UBound($categories) - 1] = $category
        EndIf
    Next
    If UBound($categories) > 1 Then _ArraySort($categories, 0, 0)

    $trayTitleMenu = TrayCreateMenu("★ " & $TitleText)

    ReDim $trayButtonItems[Buttons_Count()]
    For $i = 0 To Buttons_Count() - 1
        $trayButtonItems[$i] = TrayCreateItem(Buttons_Label($i), $trayTitleMenu)
    Next

    $traySeparator = TrayCreateItem("")

    ; Make sure trayCatMenus and trayAppItems match the current data size
    ReDim $trayCatMenus[UBound($categories)]
    ReDim $trayAppItems[UBound($appsData)]

    ; Create a menu for each category
    For $i = 0 To UBound($categories) - 1
        $trayCatMenus[$i] = TrayCreateMenu("● " & $categories[$i])
    Next

    ; Map every app to its category menu and make sure trayAppItems index matches appsData index
    For $i = 0 To UBound($appsData) - 1
        If $appsData[$i][$APP_HIDE] = "1" Then ContinueLoop
        Local $category = $appsData[$i][$APP_CAT]
        If StringLower($category) = "fave" Then ContinueLoop
        Local $catIdx = -1
        For $j = 0 To UBound($categories) - 1
            If StringLower($categories[$j]) = StringLower($category) Then
                $catIdx = $j
                ExitLoop
            EndIf
        Next
        If $catIdx >= 0 Then
            $trayAppItems[$i] = TrayCreateItem($appsData[$i][$APP_NAME], $trayCatMenus[$catIdx])
        Else
            ; If not found, create in root tray (fallback, optional)
            $trayAppItems[$i] = TrayCreateItem($appsData[$i][$APP_NAME])
        EndIf
    Next

    $trayFaveSeparator = TrayCreateItem("")
    For $i = 0 To UBound($appsData) - 1
        If $appsData[$i][$APP_HIDE] = "1" Then ContinueLoop
        If StringLower($appsData[$i][$APP_CAT]) = "fave" Then
            $trayAppItems[$i] = TrayCreateItem($appsData[$i][$APP_NAME])
        EndIf
    Next

    $trayPreExitSeparator = TrayCreateItem("")
    $trayExit = TrayCreateItem("✖ Exit")
    TraySetIcon(@ScriptDir & "\App\AxiomTrayIcon.ico")
    TraySetState()
EndFunc

Func TrayUI_HandleTrayMenu()
    While 1
        Local $msg = TrayGetMsg()
        If $msg = 0 Then
            Sleep(100)
            ContinueLoop
        EndIf

        For $i = 0 To UBound($trayButtonItems) - 1
            If $msg = $trayButtonItems[$i] Then
                Buttons_Invoke($i)
                ContinueLoop 2
            EndIf
        Next

        For $i = 0 To UBound($appsData) - 1
            If $appsData[$i][$APP_HIDE] = "1" Then ContinueLoop
            If $msg = $trayAppItems[$i] Then
                ; Defensive: Only launch if app path exists
                If $appsData[$i][$APP_PATH] <> "" And FileExists($appsData[$i][$APP_PATH]) Then
                    LaunchApp($i)
                Else
                    MsgBox(48, "AxiomTrayLauncher", "App not found: " & $appsData[$i][$APP_NAME] & @CRLF & $appsData[$i][$APP_PATH])
                EndIf
                ContinueLoop 2
            EndIf
        Next

        If $msg = $trayExit Then Exit
    WEnd
EndFunc
