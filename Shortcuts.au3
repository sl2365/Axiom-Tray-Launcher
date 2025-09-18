; Shortcuts.au3
; Generates Windows shortcuts for menu items with full variable substitution, including SetEnvN and cross-INI references

#include "Utils.au3"
#include "TrayMenu.au3"

Func _Shortcuts_CreateShortcut($target, $lnkPath, $args, $workDir)
    Local $oShell = ObjCreate("WScript.Shell")
    Local $oShort = $oShell.CreateShortcut($lnkPath)
    $oShort.TargetPath = $target
    $oShort.Arguments = $args
    $oShort.WorkingDirectory = $workDir
    $oShort.Save()
EndFunc

; --- Helper: Get all app variables from INI section ---
Func _Shortcuts_GetAppVariables($catIni, $appName)
    Local $vars = ObjCreate("Scripting.Dictionary")
    Local $keys = IniReadSection($catIni, $appName)
    If IsArray($keys) Then
        For $i = 1 To UBound($keys) - 1
            Local $k = $keys[$i][0]
            Local $v = $keys[$i][1]
            ; If this is a SetEnvN, check for cross-INI syntax
            If StringRegExp($k, "^SetEnv\d+$") Then
                $v = _Shortcuts_ResolveEnv($v)
            EndIf
            $vars.Item($k) = $v
        Next
    EndIf
    Return $vars
EndFunc

; --- Helper: Substitute %VarName% with value in string ---
Func _Shortcuts_ResolveVars($text, $vars)
    Local $resolved = $text
    Local $matches = StringRegExp($text, "%([^%]+)%", 3)
    If IsArray($matches) Then
        For $i = 0 To UBound($matches) - 1
            Local $name = $matches[$i]
            If $vars.Exists($name) Then
                Local $value = $vars.Item($name)
                ; If this variable looks like a path, resolve it
                If StringLeft($value, 2) = "?:" Then
                    $value = _ResolvePath($value, @ScriptDir)
                EndIf
                $resolved = StringReplace($resolved, "%" & $name & "%", $value)
            EndIf
        Next
    EndIf
    Return $resolved
EndFunc

; --- Helper: Resolve cross-INI variable value ---
Func _Shortcuts_ResolveEnv($envValue)
    If StringInStr($envValue, "|") Then
        Local $parts = StringSplit($envValue, "|")
        If $parts[0] = 3 Then
            Local $iniFile = @ScriptDir & "\App\" & $parts[1] & ".ini"
            Local $section = $parts[2]
            Local $key = $parts[3]
            If FileExists($iniFile) Then
                Return IniRead($iniFile, $section, $key, "")
            Else
                MsgBox(16, "Shortcut Generator Error", "INI file NOT found: " & $iniFile)
            EndIf
        Else
            MsgBox(16, "Shortcut Generator Error", "Malformed SetEnv value: " & $envValue)
        EndIf
        Return ""
    Else
        Return $envValue
    EndIf
EndFunc

Func _Shortcuts_GenerateLinks($categories, $apps, $settings)
    Local $shortcutDir = @ScriptDir & "\Shortcuts"
    If Not FileExists($shortcutDir) Then DirCreate($shortcutDir)
    Local $summary = "Links generated for these folders:" & @CRLF
    For $catName In $categories.Keys
        If $catName = "Settings" Then ContinueLoop
        Local $catFolder = $shortcutDir & "\" & $catName
        If Not FileExists($catFolder) Then DirCreate($catFolder)
        Local $catApps = $categories.Item($catName)
        Local $count = 0

        For $appName In $catApps.Keys
            Local $catIni = @ScriptDir & "\App\" & $catName & ".ini"
            Local $vars = _Shortcuts_GetAppVariables($catIni, $appName)
            ; --- Check Hide setting before creating shortcut ---
            If $vars.Exists("Hide") And $vars.Item("Hide") = "1" Then
                ContinueLoop ; Skip this app, do not create shortcut
            EndIf

            ; --- Target: Use SetEnv1, fallback to RunFile ---
            Local $target = ""
            If $vars.Exists("SetEnv1") And $vars.Item("SetEnv1") <> "" Then
                $target = _Shortcuts_ResolveVars($vars.Item("SetEnv1"), $vars)
            ElseIf $vars.Exists("RunFile") Then
                $target = $vars.Item("RunFile")
            Else
                $target = @ScriptFullPath ; fallback: launcher itself
            EndIf
            $target = _ResolvePath($target, @ScriptDir)

            ; --- Arguments: Use full user-specified string, substitute variables. DO NOT use _ResolvePath! ---
            Local $argsRaw = ""
            If $vars.Exists("Arguments") Then $argsRaw = $vars.Item("Arguments")
            Local $args = ""
            If $argsRaw <> "" Then
                $args = _Shortcuts_ResolveVars($argsRaw, $vars)
                ; --- Strip off Target if present at start of Arguments ---
                If StringLeft($args, StringLen($target)) = $target Then
                    $args = StringTrimLeft($args, StringLen($target))
                    $args = StringStripWS($args, 1)
                Else
                    Local $targetQuoted = '"' & $target & '"'
                    If StringLeft($args, StringLen($targetQuoted)) = $targetQuoted Then
                        $args = StringTrimLeft($args, StringLen($targetQuoted))
                        $args = StringStripWS($args, 1)
                    EndIf
                EndIf
            EndIf
            ; DO NOT call _ResolvePath on $args!

            ; --- WorkDir: resolve if present, fallback to @ScriptDir ---
            Local $workDir = @ScriptDir
            If $vars.Exists("WorkDir") And $vars.Item("WorkDir") <> "" Then
                $workDir = _Shortcuts_ResolveVars($vars.Item("WorkDir"), $vars)
                $workDir = _ResolvePath($workDir, @ScriptDir)
            EndIf

            ; --- Shortcut path ---
            Local $lnkPath = $catFolder & "\" & $appName & ".lnk"

            ; --- Debug output ---
;~             ConsoleWrite("Shortcut: " & $lnkPath & @CRLF)
;~             ConsoleWrite("Target: " & $target & @CRLF)
;~             ConsoleWrite("Arguments: " & $args & @CRLF)
;~             ConsoleWrite("WorkDir: " & $workDir & @CRLF)

            ; --- Create shortcut ---
            _Shortcuts_CreateShortcut($target, $lnkPath, $args, $workDir)
            $count += 1
        Next
        $summary &= $catName & ": " & $count & @CRLF
    Next
    MsgBox(64, "Shortcut Generation Complete", $summary)
EndFunc
