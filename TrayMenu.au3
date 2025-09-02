; TrayMenu.au3
; Builds and manages tray menu structure, refresh logic, supports cross-INI variable expansion in Arguments/SetEnvN
; Now supports absolute and relative paths directly, no Utils.au3 needed for path logic.

#include-once
#include "Sandboxie.au3"
#include "Favorites.au3"
#include "ScanFolders.au3"
#include "SymLinks.au3"
#include "Updates.au3"
 
Global $hSettingsGUI, $g_OKBtn
Global $g_TrayMenuTitle, $g_TrayReload, $g_TraySettings, $g_TrayGenLinks, $g_TraySeparator1, $g_TraySeparator2, $g_TraySeparator3, $g_TrayExit
Global $g_CategoryMenus = ObjCreate("Scripting.Dictionary")
Global $g_TrayItemMap = ObjCreate("Scripting.Dictionary")
Global $g_TrayScan
Global $g_TrayCreateGlobalLinks, $g_TrayRemoveGlobalLinks
Global $g_TrayCheckUpdates ; <-- ADDED: Menu item for update check

Local $globalIni = _ResolvePath("App\Settings.ini", @ScriptDir)

; --- Path Handling Functions (embedded here) ---
;~ Func _PathIsAbsolute($path)
;~     ; Returns True if path is absolute (starts with drive letter or UNC)
;~     Return StringRegExp($path, "^[A-Za-z]:\\|^\\\\")
;~ EndFunc

Func _ResolvePath($path, $baseDir)
    ; Trim whitespace
    $path = StringStripWS($path, 3)
    $baseDir = StringStripWS($baseDir, 3)
    $path = StringReplace($path, "/", "\")
    $baseDir = StringReplace($baseDir, "/", "\")

    ; If empty, return baseDir
    If $path = "" Then Return $baseDir

    ; Handle special "?" drive
    If StringLeft($path, 3) = "?:\" Then
        Local $driveLetter = StringLeft(@ScriptDir, 2) ; e.g. "F:"
        $path = $driveLetter & StringTrimLeft($path, 2)
        ; Now $path is "F:\rest..."
    EndIf

    ; UNC and absolute drive paths
    If StringRegExp($path, '^(\\\\|[A-Za-z]:\\)') Then
        Return $path
    EndIf

    ; Remove leading backslash from relative path, if present
    If StringLeft($path, 1) = "\" Then $path = StringTrimLeft($path, 1)

    ; Combine base and path
    Local $fullPath = $baseDir & "\" & $path

    ; Normalize "..\" and ".\" using AutoIt's _PathFull (requires #include <File.au3>)
    If FileExists($fullPath) Then
        $fullPath = _PathFull($fullPath)
    EndIf

    ; Remove duplicate backslashes (except for UNC root)
    $fullPath = StringRegExpReplace($fullPath, '(?<!^)\\{2,}', '\')

    Return $fullPath
EndFunc

Func _GetFileName($path)
    Local $parts = StringSplit($path, "\", 2)
    Return $parts[UBound($parts)-1]
EndFunc

Func _GetAppCategory($appName, $categories)
    For $catName In $categories.Keys
        Local $catApps = $categories.Item($catName)
        If $catApps.Exists($appName) Then
            Return $catName
        EndIf
    Next
    Return ""
EndFunc

Func _GetFolderName($path)
    Local $parts = StringSplit($path, "\", 2)
    If UBound($parts) > 0 Then
        Return $parts[UBound($parts)-1]
    Else
        Return $path
    EndIf
EndFunc

; --- Helper: Resolve cross-INI variable value ---
Func _TrayMenu_ResolveEnv($envValue)
    If StringInStr($envValue, "|") Then
        Local $parts = StringSplit($envValue, "|")
        If $parts[0] = 3 Then
            Local $iniFile = _ResolvePath("App\" & $parts[1] & ".ini", @ScriptDir)
            Local $section = $parts[2]
            Local $key = $parts[3]
            If FileExists($iniFile) Then
                Local $val = IniRead($iniFile, $section, $key, "__NOT_FOUND__")
                If $val <> "__NOT_FOUND__" Then
                    Return $val
                Else
                    MsgBox(16, "Tray Launcher Error", _
                        "INI found: " & $iniFile & @CRLF & _
                        "Section or key NOT found: [" & $section & "] " & $key)
                EndIf
            Else
                MsgBox(16, "Tray Launcher Error", _
                    "INI file NOT found: " & $iniFile)
            EndIf
            Return ""
        EndIf
        MsgBox(16, "Tray Launcher Error", "Malformed SetEnv value: " & $envValue)
        Return ""
    Else
        Return $envValue
    EndIf
EndFunc

; Create global symlinks on startup ---
_SymLink_CreateGlobalSymlinks($globalIni)

Func _TrayMenu_Build(ByRef $categories, ByRef $apps, ByRef $settings)
    TrayItemDelete(0)
    $g_TrayItemMap.RemoveAll()
    $g_CategoryMenus.RemoveAll()
    ; 1. Utility submenu
    $g_TrayMenuTitle = TrayCreateMenu("Axiom Menu")
    $g_TrayReload    = TrayCreateItem("🔄 Reload Menu", $g_TrayMenuTitle)
    $g_TraySettings  = TrayCreateItem("⚙️ Settings", $g_TrayMenuTitle)
    $g_TrayGenLinks  = TrayCreateItem("🔗 Generate Links", $g_TrayMenuTitle)
;~     $g_TrayScan      = TrayCreateItem("🧐 Scan", $g_TrayMenuTitle)
    ; --- Add buttons for global symlink management ---
    $g_TrayCreateGlobalLinks = TrayCreateItem("➕ Create Global Symlinks", $g_TrayMenuTitle)
    $g_TrayRemoveGlobalLinks = TrayCreateItem("➖ Remove Global Symlinks", $g_TrayMenuTitle)
    ; --- Add Check for Updates button ---
    $g_TrayCheckUpdates = TrayCreateItem("🔍 Check for Updates", $g_TrayMenuTitle) ; ADDED

    ; --- Determine what to show ---
    Local $hasCategories = False
    For $catName In $categories
        If $catName = "Fave" Then ContinueLoop
        If $catName = "Other" Then
            If Not $apps.Exists("Other") Or UBound($apps.Item("Other")) = 0 Then ContinueLoop
        EndIf
        If $apps.Exists($catName) Then
            Local $arr = $apps.Item($catName)
            If IsArray($arr) And UBound($arr) > 0 Then
                $hasCategories = True
                ExitLoop
            EndIf
        EndIf
    Next
    Local $hasFaves = False
    If $apps.Exists("Fave") Then
        Local $faveArr = $apps.Item("Fave")
        If IsArray($faveArr) And UBound($faveArr) > 0 Then
            $hasFaves = True
        EndIf
    EndIf
    ; --- Separator1: only if categories ---
    If $hasCategories Then
        $g_TraySeparator1 = TrayCreateItem("")
    EndIf
    ; 2. Category submenus (skip Fave and empty Other)
    For $catName In $categories
        If $catName = "Fave" Then ContinueLoop
        If $catName = "Other" Then
            If Not $apps.Exists("Other") Or UBound($apps.Item("Other")) = 0 Then ContinueLoop
        EndIf
        If $apps.Exists($catName) Then
            Local $arr = $apps.Item($catName)
            If IsArray($arr) And UBound($arr) > 0 Then
                Local $catMenu = TrayCreateMenu("📁 " & $catName)
                $g_CategoryMenus.Item($catName) = $catMenu
                For $i = 0 To UBound($arr) - 1
                    Local $appName = $arr[$i][0]
                    Local $buttonText = $arr[$i][1]
                    Local $appItem = TrayCreateItem($buttonText, $catMenu)
                    $g_TrayItemMap.Item($appName) = $appItem
                Next
            EndIf
        EndIf
    Next
    ; --- Separator2 logic ---
    If $hasFaves Then
        $g_TraySeparator2 = TrayCreateItem("")
    EndIf
    ; 3. Fave flat list
    If $hasFaves Then
        Local $faveArr = $apps.Item("Fave")
        For $i = 0 To UBound($faveArr) - 1
            Local $appName = $faveArr[$i]
            Local $catName = _GetAppCategory($appName, $categories)
            Local $catIni = _ResolvePath("App\" & $catName & ".ini", @ScriptDir)
            Local $buttonText = IniRead($catIni, $appName, "ButtonText", $appName)
            Local $faveItem = TrayCreateItem($buttonText)
            $g_TrayItemMap.Item($appName) = $faveItem
        Next
    EndIf
    ; --- Separator3: ALWAYS before Exit ---
    $g_TraySeparator3 = TrayCreateItem("")
    $g_TrayExit = TrayCreateItem("❌ Exit")
EndFunc

Func _TrayMenu_RefreshFull()
    If @Compiled Then
        Run(@ScriptFullPath)
    Else
        Run(@AutoItExe & ' "' & @ScriptFullPath & '"')
    EndIf
    Exit
EndFunc

Func TrayUI_Destroy()
    For $appName In $g_TrayItemMap.Keys
        TrayItemDelete($g_TrayItemMap.Item($appName))
    Next
    For $catName In $g_CategoryMenus.Keys
        TrayItemDelete($g_CategoryMenus.Item($catName))
    Next

    If IsDeclared("g_TrayMenuTitle") And $g_TrayMenuTitle Then TrayItemDelete($g_TrayMenuTitle)
    If IsDeclared("g_TrayReload") And $g_TrayReload Then TrayItemDelete($g_TrayReload)
    If IsDeclared("g_TraySettings") And $g_TraySettings Then TrayItemDelete($g_TraySettings)
    If IsDeclared("g_TrayGenLinks") And $g_TrayGenLinks Then TrayItemDelete($g_TrayGenLinks)
    If IsDeclared("g_TrayScan") And $g_TrayScan Then TrayItemDelete($g_TrayScan)
    If IsDeclared("g_TrayCreateGlobalLinks") And $g_TrayCreateGlobalLinks Then TrayItemDelete($g_TrayCreateGlobalLinks)
    If IsDeclared("g_TrayRemoveGlobalLinks") And $g_TrayRemoveGlobalLinks Then TrayItemDelete($g_TrayRemoveGlobalLinks)
    If IsDeclared("g_TraySeparator1") And $g_TraySeparator1 Then TrayItemDelete($g_TraySeparator1)
    If IsDeclared("g_TraySeparator2") And $g_TraySeparator2 Then TrayItemDelete($g_TraySeparator2)
    If IsDeclared("g_TraySeparator3") And $g_TraySeparator3 Then TrayItemDelete($g_TraySeparator3)
    If IsDeclared("g_TrayExit") And $g_TrayExit Then TrayItemDelete($g_TrayExit)
    If IsDeclared("g_TrayCheckUpdates") And $g_TrayCheckUpdates Then TrayItemDelete($g_TrayCheckUpdates) ; ADDED

    $g_TrayItemMap.RemoveAll()
    $g_CategoryMenus.RemoveAll()
EndFunc

Func _TrayMenu_AppLauncher($appName, $catIni)
    ; --- Read all app variables ---
    Local $keys = IniReadSection($catIni, $appName)
    Local $vars = ObjCreate("Scripting.Dictionary")
    If IsArray($keys) Then
        For $i = 1 To UBound($keys) - 1
            Local $k = $keys[$i][0]
            Local $v = $keys[$i][1]
            If StringRegExp($k, "^SetEnv\d+$") Then
                $v = _TrayMenu_ResolveEnv($v)
            EndIf
            $vars.Item($k) = $v
        Next
    EndIf

    ; --- Get Arguments and substitute variables ---
    Local $argsRaw = ""
    If $vars.Exists("Arguments") Then $argsRaw = $vars.Item("Arguments")
    Local $args = $argsRaw
    Local $matches = StringRegExp($argsRaw, "%([^%]+)%", 3)
    If IsArray($matches) Then
        For $i = 0 To UBound($matches) - 1
            Local $vn = $matches[$i]
            If $vars.Exists($vn) Then
                $args = StringReplace($args, "%" & $vn & "%", $vars.Item($vn))
            EndIf
        Next
    EndIf

    ; --- Get executable (prefer SetEnv1, fallback to RunFile) ---
    Local $exe = ""
    If $vars.Exists("SetEnv1") And $vars.Item("SetEnv1") <> "" Then
        $exe = $vars.Item("SetEnv1")
    ElseIf $vars.Exists("RunFile") And $vars.Item("RunFile") <> "" Then
        $exe = $vars.Item("RunFile")
    EndIf
    $exe = _ResolvePath($exe, @ScriptDir)

    ; --- Determine workDir ---
    Local $workDir = ""
    If $vars.Exists("WorkDir") And $vars.Item("WorkDir") <> "" Then
        $workDir = _ResolvePath($vars.Item("WorkDir"), @ScriptDir)
    Else
        Local $lastSlash = StringInStr($exe, "\", 0, -1)
        If $lastSlash > 0 Then
            $workDir = StringLeft($exe, $lastSlash - 1)
        Else
            $workDir = @ScriptDir
        EndIf
    EndIf

    ; --- Check exe is not a folder and is a valid executable ---
    Local $ext = StringLower(StringRight($exe, 4))
    If $exe = "" _
        Or Not FileExists($exe) _
        Or ($ext <> ".exe" And $ext <> ".bat" And $ext <> ".cmd") Then
        MsgBox(16, $appName & " Launch Error", _
            "Executable not found or not a valid exe/batch/cmd file:" & @CRLF & $exe)
        Return
    EndIf

    ; --- SingleInstance logic (per app via INI) ---
    Local $exeName = StringRegExpReplace($exe, "^.*\\", "")
    Local $singleInstance = 0
    If $vars.Exists("SingleInstance") Then
        $singleInstance = Number($vars.Item("SingleInstance"))
    EndIf

    If $singleInstance = 1 Then
        Local $isRunning = ProcessExists($exeName)
        If $isRunning Then
            MsgBox(64, $appName & " Already Running", _
                "There is an instance of this app/file open already.")
            Return
        EndIf
    EndIf

    ; --- Create per-app/category symlinks BEFORE launch ---
    _SymLink_CreateAppSymlinks($catIni, $appName, $globalIni)

    ; --- RunAsAdmin logic (per app via INI) ---
    Local $runAsAdmin = "0"
    If $vars.Exists("RunAsAdmin") Then
        $runAsAdmin = $vars.Item("RunAsAdmin")
    EndIf

    ; --- Sandboxie logic (per app via INI) ---
    Local $sandboxie = "0", $sandboxName = ""
    If $vars.Exists("Sandboxie") Then $sandboxie = $vars.Item("Sandboxie")
    If $vars.Exists("SandboxName") Then $sandboxName = $vars.Item("SandboxName")
    Local $settingsIni = @ScriptDir & "\App\Settings.ini"

    ; --- Strip off executable path if present at start of Arguments ---
    If StringLeft($args, StringLen($exe)) = $exe Then
        $args = StringTrimLeft($args, StringLen($exe))
        $args = StringStripWS($args, 1)
    Else
        Local $exeQuoted = '"' & $exe & '"'
        If StringLeft($args, StringLen($exeQuoted)) = $exeQuoted Then
            $args = StringTrimLeft($args, StringLen($exeQuoted))
            $args = StringStripWS($args, 1)
        EndIf
    EndIf

    ; --- Launch ---
    Local $appExited = False
    If $sandboxie = "1" And $sandboxName <> "" Then
        Local $pid = _RunWithSandboxie($exe, $args, $workDir, $sandboxName, $settingsIni)
        If $pid <> 0 Then
            ProcessWaitClose($pid)
            $appExited = True
        EndIf
    ElseIf $runAsAdmin = "1" Then
        ShellExecute($exe, $args, $workDir, "runas")
        ; Can't reliably detect process exit for ShellExecute/runas
        $appExited = False
    Else
        Local $pid = Run('"' & $exe & '" ' & $args, $workDir)
        ; Do NOT block with ProcessWaitClose, tray must stay responsive!
        ; $appExited = False   ; If you need to do cleanup, handle it elsewhere (e.g., tray exit or background watcher)
    EndIf

    ; --- Remove per-app/category symlinks after app closes ---
    If $appExited Then
        _SymLink_RemoveAppSymlinks($catIni, $appName, $globalIni)
    EndIf
EndFunc

Func _TrayMenu_HandleEvents(ByRef $settings, ByRef $categories, ByRef $apps)
    Local $msg = TrayGetMsg()
    Switch $msg
        Case $g_TrayReload
            _TrayMenu_RefreshFull()
        Case $g_TrayGenLinks
            _Shortcuts_GenerateLinks($categories, $apps, $settings)
        Case $g_TraySettings
            ShowSettingsGUI()
			While $hSettingsGUI <> 0
				SettingsGUI_HandleEvents()
			WEnd
        Case $g_TrayCreateGlobalLinks
            Local $globalIni = _ResolvePath("App\Settings.ini", @ScriptDir)
            _SymLink_ManualCreateGlobalSymlinks($globalIni)
            MsgBox(64, "Global Symlinks", "Global symlinks created.")
        Case $g_TrayRemoveGlobalLinks
            Local $globalIni = _ResolvePath("App\Settings.ini", @ScriptDir)
            _SymLink_ManualRemoveGlobalSymlinks($globalIni)
            MsgBox(64, "Global Symlinks", "Global symlinks removed.")
        Case $g_TrayCheckUpdates
            Updates_Check(True)
        Case $g_TrayExit
			_SymLink_TrayExitCleanup()
            _SymLink_RemoveGlobalSymlinks($globalIni)
            Exit
        Case Else
            For $appName In $g_TrayItemMap.Keys
                If $msg = $g_TrayItemMap.Item($appName) Then
                    Local $catName = _GetAppCategory($appName, $categories)
                    Local $catIni = _ResolvePath("App\" & $catName & ".ini", @ScriptDir)
                    If Not FileExists($catIni) Then
                        MsgBox(16, "INI Error", "INI file not found: " & $catIni)
                        ExitLoop
                    EndIf
                    Local $keys = IniReadSection($catIni, $appName)
                    If Not IsArray($keys) Then
                        MsgBox(16, "INI Error", "Section [" & $appName & "] not found in " & $catIni)
                        ExitLoop
                    EndIf
                    _TrayMenu_AppLauncher($appName, $catIni)
                    ExitLoop
                EndIf
            Next
    EndSwitch
EndFunc
