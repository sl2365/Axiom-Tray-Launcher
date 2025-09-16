; TrayMenu.au3
; Builds and manages tray menu structure, refresh logic, supports cross-INI variable expansion in Arguments/SetEnvN
; Now supports absolute and relative paths directly, no Utils.au3 needed for path logic.

#include-once
#include <Array.au3>
#include "Sandboxie.au3"
#include "Favorites.au3"
#include "ScanFolders.au3"
#include "SymLinks.au3"
 
Global $hSettingsGUI, $g_OKBtn
Global $g_TrayMenuTitle, $g_TrayReload, $g_TraySettings, $g_TraySeparator1, $g_TraySeparator2, $g_TraySeparator3, $g_TrayExit
Global $g_CategoryMenus = ObjCreate("Scripting.Dictionary")
Global $g_TrayItemMap = ObjCreate("Scripting.Dictionary")
Global $g_TrayScan
Global $g_TrayCreateGlobalLinks, $g_TrayRemoveGlobalLinks
Global $g_MonitoredApps
Local $globalIni = _ResolvePath("App\Settings.ini", @ScriptDir)

; --- Path Handling Functions (embedded here) ---
Func _ResolvePath($path, $baseDir)
    $path = StringStripWS($path, 3)
    $baseDir = StringStripWS($baseDir, 3)
    $path = StringReplace($path, "/", "\")
    $baseDir = StringReplace($baseDir, "/", "\")

    If $path = "" Then Return $baseDir

    If StringLeft($path, 2) = "?:" Then
		Local $driveLetter = StringLeft(@ScriptDir, 2)
		$path = $driveLetter & StringMid($path, 3)
	EndIf

    If StringRegExp($path, '^(\\\\|[A-Za-z]:\\)') Then
        Return $path
    EndIf

    If StringLeft($path, 1) = "\" Then $path = StringTrimLeft($path, 1)

    Local $fullPath = $baseDir & "\" & $path

    If FileExists($fullPath) Then
        $fullPath = _PathFull($fullPath)
    EndIf

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

_SymLink_CreateGlobalSymlinks($globalIni)

Func _TrayMenu_Build(ByRef $categories, ByRef $apps, ByRef $settings)
    TrayItemDelete(0)
    $g_TrayItemMap.RemoveAll()
    $g_CategoryMenus.RemoveAll()
    $g_TrayMenuTitle = TrayCreateMenu("♾️ Axiom Menu")
    $g_TrayReload    = TrayCreateItem("🔄 Reload Menu", $g_TrayMenuTitle)
    $g_TraySettings  = TrayCreateItem("⚙️ Settings", $g_TrayMenuTitle)
    $g_TrayCreateGlobalLinks = TrayCreateItem("➕ Create Global Symlinks", $g_TrayMenuTitle)
    $g_TrayRemoveGlobalLinks = TrayCreateItem("➖ Remove Global Symlinks", $g_TrayMenuTitle)

    Local $hasCategories = False
    Local $sortedCategories[0]
    For $catName In $categories
        If $catName = "Fave" Then ContinueLoop
        If $catName = "Other" Then
            If Not $apps.Exists("Other") Or UBound($apps.Item("Other")) = 0 Then ContinueLoop
        EndIf
        If $apps.Exists($catName) Then
            Local $arr = $apps.Item($catName)
            If IsArray($arr) And UBound($arr) > 0 Then
                $hasCategories = True
                _ArrayAdd($sortedCategories, $catName)
            EndIf
        EndIf
    Next
    If UBound($sortedCategories) > 0 Then _ArraySort($sortedCategories)
    Local $hasFaves = False
    If $apps.Exists("Fave") Then
        Local $faveArr = $apps.Item("Fave")
        If IsArray($faveArr) And UBound($faveArr) > 0 Then
            $hasFaves = True
        EndIf
    EndIf
    If $hasCategories Then
        $g_TraySeparator1 = TrayCreateItem("")
    EndIf

    For $i = 0 To UBound($sortedCategories) - 1
        Local $catName = $sortedCategories[$i]
        If $catName = "Fave" Then ContinueLoop
        If $catName = "Other" Then
            If Not $apps.Exists("Other") Or UBound($apps.Item("Other")) = 0 Then ContinueLoop
        EndIf
        If $apps.Exists($catName) Then
            Local $arr = $apps.Item($catName)
            If IsArray($arr) And UBound($arr) > 0 Then
                Local $catMenu = TrayCreateMenu("📁 " & $catName)
                $g_CategoryMenus.Item($catName) = $catMenu
                Local $sortedArr[UBound($arr)][2]
                For $j = 0 To UBound($arr) - 1
                    $sortedArr[$j][0] = $arr[$j][0]
                    $sortedArr[$j][1] = $arr[$j][1]
                Next
                _ArraySort($sortedArr, 0, 0, 0, 1)
                For $j = 0 To UBound($sortedArr) - 1
                    Local $appName = $sortedArr[$j][0]
                    Local $buttonText = $sortedArr[$j][1]
                    Local $appItem = TrayCreateItem($buttonText, $catMenu)
                    $g_TrayItemMap.Item($appName) = $appItem
                Next
            EndIf
        EndIf
    Next

    If $hasFaves Then
        $g_TraySeparator2 = TrayCreateItem("")
    EndIf

    If $hasFaves Then
        Local $faveArr = $apps.Item("Fave")
        Local $sortedFaves[UBound($faveArr)]
        For $i = 0 To UBound($faveArr) - 1
            $sortedFaves[$i] = $faveArr[$i]
        Next
        _ArraySort($sortedFaves)
        For $i = 0 To UBound($sortedFaves) - 1
            Local $appName = $sortedFaves[$i]
            Local $catName = _GetAppCategory($appName, $categories)
            Local $catIni = _ResolvePath("App\" & $catName & ".ini", @ScriptDir)
            Local $buttonText = IniRead($catIni, $appName, "ButtonText", $appName)
            Local $faveItem = TrayCreateItem($buttonText)
            $g_TrayItemMap.Item($appName) = $faveItem
        Next
    EndIf
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
    If IsDeclared("g_TrayScan") And $g_TrayScan Then TrayItemDelete($g_TrayScan)
    If IsDeclared("g_TrayCreateGlobalLinks") And $g_TrayCreateGlobalLinks Then TrayItemDelete($g_TrayCreateGlobalLinks)
    If IsDeclared("g_TrayRemoveGlobalLinks") And $g_TrayRemoveGlobalLinks Then TrayItemDelete($g_TrayRemoveGlobalLinks)
    If IsDeclared("g_TraySeparator1") And $g_TraySeparator1 Then TrayItemDelete($g_TraySeparator1)
    If IsDeclared("g_TraySeparator2") And $g_TraySeparator2 Then TrayItemDelete($g_TraySeparator2)
    If IsDeclared("g_TraySeparator3") And $g_TraySeparator3 Then TrayItemDelete($g_TraySeparator3)
    If IsDeclared("g_TrayExit") And $g_TrayExit Then TrayItemDelete($g_TrayExit)

    $g_TrayItemMap.RemoveAll()
    $g_CategoryMenus.RemoveAll()
EndFunc

Func _TrayMenu_AppLauncher($appName, $catIni)
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

    Local $exe = ""
    If $vars.Exists("SetEnv1") And $vars.Item("SetEnv1") <> "" Then
        $exe = $vars.Item("SetEnv1")
    ElseIf $vars.Exists("RunFile") And $vars.Item("RunFile") <> "" Then
        $exe = $vars.Item("RunFile")
    EndIf
    $exe = _ResolvePath($exe, @ScriptDir)

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

    Local $ext = StringLower(StringRight($exe, 4))
    If $exe = "" _
        Or Not FileExists($exe) _
        Or ($ext <> ".exe" And $ext <> ".bat" And $ext <> ".cmd") Then
        MsgBox(16, $appName & " Launch Error", _
            "Executable not found or not a valid exe/batch/cmd file:" & @CRLF & $exe)
        Return
    EndIf

    Local $exeName = $exe
    If StringInStr($exeName, '\') Then $exeName = StringRegExpReplace($exeName, '^.*\\', '')
    If StringInStr($exeName, ' ') Then $exeName = StringLeft($exeName, StringInStr($exeName, ' ') - 1)
    $exeName = StringReplace($exeName, '"', '')

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

    _SymLink_CreateAppSymlinks($catIni, $appName, $globalIni)

    Local $runAsAdmin = "0"
    If $vars.Exists("RunAsAdmin") Then
        $runAsAdmin = $vars.Item("RunAsAdmin")
    EndIf

    Local $sandboxie = "0", $sandboxName = ""
    If $vars.Exists("Sandboxie") Then $sandboxie = $vars.Item("Sandboxie")
    If $vars.Exists("SandboxName") Then $sandboxName = $vars.Item("SandboxName")
    Local $settingsIni = @ScriptDir & "\App\Settings.ini"

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

	Local $appExited = False
	If $sandboxie = "1" And $sandboxName <> "" Then
		Local $pidArr = _RunWithSandboxie($exe, $args, $workDir, $sandboxName, $settingsIni)
		Local $sandboxedPid = $pidArr[1]
		If $sandboxedPid <> 0 And $sandboxedPid <> "" Then
			Local $newLen = UBound($g_MonitoredApps) + 1
			ReDim $g_MonitoredApps[$newLen][4]
			$g_MonitoredApps[$newLen - 1][0] = $sandboxedPid
			$g_MonitoredApps[$newLen - 1][1] = 1 ; Sandboxie
			$g_MonitoredApps[$newLen - 1][2] = $catIni
			$g_MonitoredApps[$newLen - 1][3] = $appName
		EndIf
	ElseIf $runAsAdmin = "1" Then
		ShellExecute($exe, $args, $workDir, "runas")
	Else
		Local $pid = Run('"' & $exe & '" ' & $args, $workDir)
		If $pid <> 0 And $pid <> "" Then
			Local $newLen = UBound($g_MonitoredApps) + 1
			ReDim $g_MonitoredApps[$newLen][4]
			$g_MonitoredApps[$newLen - 1][0] = $pid
			$g_MonitoredApps[$newLen - 1][1] = 0 ; Not sandboxie
			$g_MonitoredApps[$newLen - 1][2] = $catIni
			$g_MonitoredApps[$newLen - 1][3] = $appName
		EndIf
	EndIf

    ; DO NOT remove symlinks here; monitoring/cleanup is handled in main file via AdlibRegister("MonitorApps", 500)
EndFunc

Func _TrayMenu_HandleEvents(ByRef $settings, ByRef $categories, ByRef $apps)
    Local $msg = TrayGetMsg()
    Switch $msg
        Case $g_TrayReload
            _TrayMenu_RefreshFull()
        Case $g_TraySettings
            ShowSettingsGUI()
			While $hSettingsGUI <> 0
				SettingsGUI_HandleEvents()
			WEnd
        Case $g_TrayCreateGlobalLinks
            Local $globalIni = _ResolvePath("App\Settings.ini", @ScriptDir)
            _SymLink_ManualCreateGlobalSymlinks($globalIni)
        Case $g_TrayRemoveGlobalLinks
            Local $globalIni = _ResolvePath("App\Settings.ini", @ScriptDir)
            _SymLink_ManualRemoveGlobalSymlinks($globalIni)
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
