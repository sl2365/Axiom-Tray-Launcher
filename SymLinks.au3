; SymLinks.au3

#include-once

; --------- Symlink Variable Expansion ---------
Func _SymLink_ExpandVarsInPath($path, $vars)
    ;ConsoleWrite("[_SymLink_ExpandVarsInPath] Input path: " & $path & @CRLF)
    For $key In $vars.Keys()
        ;ConsoleWrite("[_SymLink_ExpandVarsInPath] Replacing %" & $key & "% with " & $vars.Item($key) & @CRLF)
        $path = StringReplace($path, "%" & $key & "%", $vars.Item($key))
    Next
    Local $re = StringRegExp($path, "%([A-Za-z0-9_]+)%", 3)
    If IsArray($re) Then
        For $i = 0 To UBound($re) - 1
            ;ConsoleWrite("[_SymLink_ExpandVarsInPath] Replacing env %" & $re[$i] & "% with " & EnvGet($re[$i]) & @CRLF)
            $path = StringReplace($path, "%" & $re[$i] & "%", EnvGet($re[$i]))
        Next
    EndIf
    ;ConsoleWrite("[_SymLink_ExpandVarsInPath] Output path: " & $path & @CRLF)
    Return $path
EndFunc

Func _SymLink_ExpandEnv($str)
    ;ConsoleWrite("[_SymLink_ExpandEnv] Input string: " & $str & @CRLF)
    Local $re = StringRegExp($str, "%([A-Za-z0-9_]+)%", 3)
    If IsArray($re) Then
        For $i = 0 To UBound($re) - 1
            ;ConsoleWrite("[_SymLink_ExpandEnv] Replacing env %" & $re[$i] & "% with " & EnvGet($re[$i]) & @CRLF)
            $str = StringReplace($str, "%" & $re[$i] & "%", EnvGet($re[$i]))
        Next
    EndIf
    ;ConsoleWrite("[_SymLink_ExpandEnv] Output string: " & $str & @CRLF)
    Return $str
EndFunc

Func _SymLink_LoadVariables($iniFile)
    ;ConsoleWrite("[_SymLink_LoadVariables] Loading variables from: " & $iniFile & @CRLF)
    Local $vars = ObjCreate("Scripting.Dictionary")
    Local $section = "Variables"
    Local $varNames = IniReadSection($iniFile, $section)
    If IsArray($varNames) Then
        For $i = 1 To $varNames[0][0]
            ;ConsoleWrite("[_SymLink_LoadVariables] Found var: " & $varNames[$i][0] & " = " & $varNames[$i][1] & @CRLF)
            $vars.Add($varNames[$i][0], _SymLink_ExpandEnv($varNames[$i][1]))
        Next
    Else
        ;ConsoleWrite("[_SymLink_LoadVariables] No [Variables] section found!" & @CRLF)
    EndIf
    Return $vars
EndFunc

; --------- Symlink Creation & Removal ---------
Func _SymLink_CreateSymlink($symlink_path, $target_path)
    ;MsgBox(64, "DEBUG", "_SymLink_CreateSymlink called!" & @CRLF & _
    ;    "Symlink Path: " & $symlink_path & @CRLF & "Target Path: " & $target_path)
    ;ConsoleWrite("[_SymLink_CreateSymlink] Called with: " & $symlink_path & " -> " & $target_path & @CRLF)
    Local $symlinkParent = StringLeft($symlink_path, StringInStr($symlink_path, "\", 0, -1) - 1)
    ;ConsoleWrite("[_SymLink_CreateSymlink] Parent directory: " & $symlinkParent & @CRLF)
    If Not FileExists($symlinkParent) Then
        ;ConsoleWrite("[_SymLink_CreateSymlink] Creating parent directory: " & $symlinkParent & @CRLF)
        DirCreate($symlinkParent)
    EndIf
    If FileExists($symlink_path) Then
        If _SymLink_IsSymlinkOrJunction($symlink_path) Then
            ;ConsoleWrite("[_SymLink_CreateSymlink] Symlink already exists: " & $symlink_path & @CRLF)
            MsgBox(48, "Symlink Exists", "Symlink already exists: " & $symlink_path)
            Return
        Else
            ;ConsoleWrite("[_SymLink_CreateSymlink] Directory exists and is NOT a symlink: " & $symlink_path & @CRLF)
            MsgBox(16, "Symlink Conflict", "Directory exists and is NOT a symlink: " & $symlink_path)
            Return
        EndIf
    EndIf
    ;ConsoleWrite("[_SymLink_CreateSymlink] Creating symlink: " & $symlink_path & " -> " & $target_path & @CRLF)
    ;MsgBox(0, "Symlink Create", "Requesting admin: mklink /D " & @CRLF & $symlink_path & @CRLF & "->" & $target_path)
    Local $cmd = 'mklink /D "' & $symlink_path & '" "' & $target_path & '"'
    ;ConsoleWrite("[_SymLink_CreateSymlink] ShellExecute Command: " & @ComSpec & " /c " & $cmd & @CRLF)
    ShellExecute(@ComSpec, " /c " & $cmd, "", "runas")
    Sleep(500)
    If FileExists($symlink_path) Then
        MsgBox(64, "Symlink Created", "Symlink created successfully: " & $symlink_path)
    Else
        MsgBox(16, "Symlink NOT Created", "Symlink was NOT created: " & $symlink_path)
    EndIf
EndFunc

Func _SymLink_DeleteSymlink($symlink_path)
    ;ConsoleWrite("[_SymLink_DeleteSymlink] Called with: " & $symlink_path & @CRLF)
    If Not _SymLink_IsSymlinkOrJunction($symlink_path) Then
        ;ConsoleWrite("[_SymLink_DeleteSymlink] Not a symlink/junction: " & $symlink_path & @CRLF)
        MsgBox(16, "Delete Symlink", "Not a symlink/junction: " & $symlink_path)
        Return
    EndIf
    ;ConsoleWrite("[_SymLink_DeleteSymlink] Deleting symlink: " & $symlink_path & @CRLF)
    ;MsgBox(0, "Symlink Delete", "Requesting admin: rmdir " & @CRLF & $symlink_path)
    Local $cmd = 'rmdir "' & $symlink_path & '"'
    ;ConsoleWrite("[_SymLink_DeleteSymlink] ShellExecute Command: " & @ComSpec & " /c " & $cmd & @CRLF)
    ShellExecute(@ComSpec, " /c " & $cmd, "", "runas")
    Sleep(500)
    If Not FileExists($symlink_path) Then
        MsgBox(64, "Symlink Deleted", "Symlink deleted successfully: " & $symlink_path)
    Else
        MsgBox(16, "Symlink NOT Deleted", "Symlink was NOT deleted: " & $symlink_path)
    EndIf
EndFunc

Func _SymLink_IsSymlinkOrJunction($path)
    ;ConsoleWrite("[_SymLink_IsSymlinkOrJunction] Testing: " & $path & @CRLF)
    Local $cmd = 'fsutil reparsepoint query "' & $path & '"'
    Local $pid = Run(@ComSpec & " /c " & $cmd, "", @SW_HIDE, $STDOUT_CHILD)
    Local $output = ""
    While True
        $output &= StdoutRead($pid)
        If @error Then ExitLoop
    WEnd
    ;ConsoleWrite("[_SymLink_IsSymlinkOrJunction] Output: " & $output & @CRLF)
    If StringInStr($output, "Symbolic Link") Or StringInStr($output, "Junction") Then Return True
    Return False
EndFunc

; --------- GLOBAL SYMLINKS ONLY ---------
Func _SymLink_CreateGlobalSymlinks($globalIni, $force = False)
    ;MsgBox(64, "DEBUG", "_SymLink_CreateGlobalSymlinks called! INI: " & $globalIni)
    Local $doAdd = IniRead($globalIni, "SymLinks", "SymLinksAdd", "0")
    If (Not $force) And (StringStripWS($doAdd, 3) <> "1") Then
        ;ConsoleWrite("[_SymLink_CreateGlobalSymlinks] SymLinksAdd not enabled, returning. Actual value: '" & $doAdd & "'" & @CRLF)
        Return
    ElseIf (Not $force) Then
        ;ConsoleWrite("[_SymLink_CreateGlobalSymlinks] SymLinksAdd enabled, proceeding." & @CRLF)
    EndIf
    Local $vars = _SymLink_LoadVariables($globalIni)
    Local $symlinks = IniReadSection($globalIni, "SymLinks")
    If Not IsArray($symlinks) Then
        ;ConsoleWrite("[_SymLink_CreateGlobalSymlinks] No [SymLinks] section found!" & @CRLF)
        Return
    EndIf
    ;ConsoleWrite("[_SymLink_CreateGlobalSymlinks] Found " & $symlinks[0][0] & " symlink entries." & @CRLF)
    For $i = 1 To $symlinks[0][0]
        Local $key = $symlinks[$i][0]
        If StringLeft($key, 7) <> "SymLink" Then
            ;ConsoleWrite("[_SymLink_CreateGlobalSymlinks] Ignoring key: " & $key & @CRLF)
            ContinueLoop
        EndIf
        Local $val = $symlinks[$i][1]
        ;ConsoleWrite("[_SymLink_CreateGlobalSymlinks] Processing: " & $key & " = " & $val & @CRLF)
        Local $split = StringSplit($val, "~", 2) ; <--- CHANGED FROM "|" TO "~"
        If UBound($split) = 2 Then
            Local $symlink_path = _SymLink_ExpandVarsInPath($split[0], $vars)
            Local $target_path  = _SymLink_ExpandVarsInPath($split[1], $vars)
            ;MsgBox(64, "DEBUG", "Global Symlink Spec" & @CRLF & "Key: " & $key & @CRLF & _
            ;    "Symlink Path: " & $symlink_path & @CRLF & "Target Path: " & $target_path)
            _SymLink_CreateSymlink($symlink_path, $target_path)
        EndIf
    Next
EndFunc

Func _SymLink_RemoveGlobalSymlinks($globalIni, $force = False)
    ;MsgBox(64, "DEBUG", "_SymLink_RemoveGlobalSymlinks called! INI: " & $globalIni)
    Local $doRemove = IniRead($globalIni, "SymLinks", "SymLinksRemove", "0")
    If (Not $force) And (StringStripWS($doRemove, 3) <> "1") Then
        ;ConsoleWrite("[_SymLink_RemoveGlobalSymlinks] SymLinksRemove not enabled, returning. Actual value: '" & $doRemove & "'" & @CRLF)
        Return
    ElseIf (Not $force) Then
        ;ConsoleWrite("[_SymLink_RemoveGlobalSymlinks] SymLinksRemove enabled, proceeding." & @CRLF)
    EndIf
    Local $vars = _SymLink_LoadVariables($globalIni)
    Local $symlinks = IniReadSection($globalIni, "SymLinks")
    If Not IsArray($symlinks) Then
        ;ConsoleWrite("[_SymLink_RemoveGlobalSymlinks] No [SymLinks] section found!" & @CRLF)
        Return
    EndIf
    ;ConsoleWrite("[_SymLink_RemoveGlobalSymlinks] Found " & $symlinks[0][0] & " symlink entries." & @CRLF)
    For $i = 1 To $symlinks[0][0]
        Local $key = $symlinks[$i][0]
        If StringLeft($key, 7) <> "SymLink" Then
            ;ConsoleWrite("[_SymLink_RemoveGlobalSymlinks] Ignoring key: " & $key & @CRLF)
            ContinueLoop
        EndIf
        Local $val = $symlinks[$i][1]
        ;ConsoleWrite("[_SymLink_RemoveGlobalSymlinks] Processing: " & $key & " = " & $val & @CRLF)
        Local $split = StringSplit($val, "~", 2) ; <--- CHANGED FROM "|" TO "~"
        If UBound($split) = 2 Then
            Local $symlink_path = _SymLink_ExpandVarsInPath($split[0], $vars)
            ;MsgBox(64, "DEBUG", "Global Symlink Remove Spec" & @CRLF & "Key: " & $key & @CRLF & "Symlink Path: " & $symlink_path)
            _SymLink_DeleteSymlink($symlink_path)
        EndIf
    Next
EndFunc

Func _SymLink_ManualCreateGlobalSymlinks($globalIni)
    ;MsgBox(64, "DEBUG", "_SymLink_ManualCreateGlobalSymlinks called")
    Local $result = MsgBox(33, "Manual Global Symlinks", "Create global symlinks now?" & @CRLF & "Click OK to proceed or Cancel to abort.")
    If $result = 1 Then ; OK
        _SymLink_CreateGlobalSymlinks($globalIni, True)
    Else ; Cancel
        ; Operation aborted
    EndIf
EndFunc

Func _SymLink_ManualRemoveGlobalSymlinks($globalIni)
    ;MsgBox(64, "DEBUG", "_SymLink_ManualRemoveGlobalSymlinks called")
    Local $result = MsgBox(33, "Manual Global Symlinks", "Remove global symlinks now?" & @CRLF & "Click OK to proceed or Cancel to abort.")
    If $result = 1 Then ; OK
        _SymLink_RemoveGlobalSymlinks($globalIni, True)
    Else ; Cancel
        ; Operation aborted
    EndIf
EndFunc

; --------- PER-APP SYMLINKS ONLY ---------
Func _SymLink_CreateAppSymlinks($catIni, $appName, $globalIni)
    ;MsgBox(64, "DEBUG", "_SymLink_CreateAppSymlinks called! INI: " & $catIni & " Section: " & $appName)
    Local $doAdd = IniRead($catIni, $appName, "SymLinkCreate", "0")
    ;ConsoleWrite("[_SymLink_CreateAppSymlinks] SymLinkCreate: '" & $doAdd & "' | HEX: " & Hex(Asc(StringLeft($doAdd, 1)), 2) & @CRLF)
    If StringStripWS($doAdd, 3) <> "1" Then
        ;ConsoleWrite("[_SymLink_CreateAppSymlinks] SymLinkCreate not enabled, returning. Actual value: '" & $doAdd & "'" & @CRLF)
        Return
    Else
        ;ConsoleWrite("[_SymLink_CreateAppSymlinks] SymLinkCreate enabled, proceeding." & @CRLF)
    EndIf
    ; Merge variables from global INI and app section
    Local $vars = _SymLink_LoadVariables($globalIni)
    Local $appVars = ObjCreate("Scripting.Dictionary")
    Local $appSectionVars = IniReadSection($catIni, $appName)
    If IsArray($appSectionVars) Then
        For $i = 1 To $appSectionVars[0][0]
            $appVars.Add($appSectionVars[$i][0], $appSectionVars[$i][1])
        Next
    EndIf
    For $key In $appVars.Keys()
        $vars.Item($key) = $appVars.Item($key)
    Next
    Local $symlinks = IniReadSection($catIni, $appName)
    If Not IsArray($symlinks) Then
        ;ConsoleWrite("[_SymLink_CreateAppSymlinks] No symlink section found!" & @CRLF)
        Return
    EndIf
    For $i = 1 To $symlinks[0][0]
        Local $key = $symlinks[$i][0]
        If StringLeft($key, 7) <> "SymLink" Then
            ContinueLoop
        EndIf
        Local $val = $symlinks[$i][1]
        Local $split = StringSplit($val, "~", 2) ; <--- CHANGED FROM "|" TO "~"
        If UBound($split) = 2 Then
            Local $target = _SymLink_ExpandVarsInPath($split[0], $vars)
            Local $link = _SymLink_ExpandVarsInPath($split[1], $vars)
            ;MsgBox(64, "DEBUG", "App Symlink Spec" & @CRLF & "Key: " & $key & @CRLF & "Target: " & $target & @CRLF & "Link: " & $link)
            _SymLink_CreateSymlink($target, $link)
        EndIf
    Next
EndFunc

Func _SymLink_RemoveAppSymlinks($catIni, $appName, $globalIni = "")
    ;MsgBox(64, "DEBUG", "_SymLink_RemoveAppSymlinks called: " & $catIni & ", " & $appName)
    If StringStripWS(IniRead($catIni, $appName, "SymLinkCreate", "0"), 3) <> "1" Then
        ;ConsoleWrite("[_SymLink_RemoveAppSymlinks] SymLinkCreate not enabled for section: " & $appName & @CRLF)
        Return
    EndIf
    Local $vars = ($globalIni <> "") ? _SymLink_LoadVariables($globalIni) : ObjCreate("Scripting.Dictionary")
    Local $appVars = ObjCreate("Scripting.Dictionary")
    Local $appSectionVars = IniReadSection($catIni, $appName)
    If IsArray($appSectionVars) Then
        For $i = 1 To $appSectionVars[0][0]
            $appVars.Add($appSectionVars[$i][0], $appSectionVars[$i][1])
        Next
    EndIf
    For $key In $appVars.Keys()
        $vars.Item($key) = $appVars.Item($key)
    Next
    For $j = 1 To 10
        Local $symKey = "SymLink" & $j
        Local $symSpec = IniRead($catIni, $appName, $symKey, "")
        If $symSpec = "" Then ContinueLoop
        ;ConsoleWrite("[_SymLink_RemoveAppSymlinks] Processing: " & $symKey & " = " & $symSpec & @CRLF)
        Local $parts = StringSplit($symSpec, "~", 2) ; <--- CHANGED FROM "|" TO "~"
        If UBound($parts) = 2 Then
            Local $symlink_path = _SymLink_ExpandVarsInPath($parts[0], $vars)
            ;MsgBox(64, "DEBUG", "App Symlink Remove Spec" & @CRLF & "Key: " & $symKey & @CRLF & "Symlink Path: " & $symlink_path)
            _SymLink_DeleteSymlink($symlink_path)
        EndIf
    Next
EndFunc

; --------- Optional Debug ---------
Func _SymLink_DumpVariables($globalIni)
    ;ConsoleWrite("[_SymLink_DumpVariables] Global INI: " & $globalIni & @CRLF)
    Local $vars = _SymLink_LoadVariables($globalIni)
    For $key In $vars.Keys()
        ;ConsoleWrite($key & " = " & $vars.Item($key) & @CRLF)
    Next
EndFunc

; --------- Cleanup On Exit ---------
Global $settingsIni = @ScriptDir & "\App\Settings.ini"
Global $categoryIniDir = @ScriptDir & "\App"

; Checks if any app in any category ini has SymLinkCreate=1
Func _SymLink_ShouldCleanupOnExit()
    Local $aFiles = _FileListToArray($categoryIniDir, "*.ini", 1)
    If @error Then Return False
    For $i = 1 To $aFiles[0]
        Local $catIni = $categoryIniDir & "\" & $aFiles[$i]
        Local $appSections = IniReadSectionNames($catIni)
        If @error Or Not IsArray($appSections) Then ContinueLoop
        For $j = 1 To $appSections[0]
            Local $appName = $appSections[$j]
            If IniRead($catIni, $appName, "SymLinkCreate", "0") == "1" Then
                Return True ; Found at least one app needing cleanup
            EndIf
        Next
    Next
    Return False
EndFunc

; Removes all per-app symlinks for all apps with SymLinkCreate=1 in all categories
Func _SymLink_CleanupAllPerAppSymlinks()
    Local $aFiles = _FileListToArray($categoryIniDir, "*.ini", 1)
    If @error Then Return
    For $i = 1 To $aFiles[0]
        Local $catIni = $categoryIniDir & "\" & $aFiles[$i]
        Local $appSections = IniReadSectionNames($catIni)
        If @error Or Not IsArray($appSections) Then ContinueLoop
        For $j = 1 To $appSections[0]
            Local $appName = $appSections[$j]
            If IniRead($catIni, $appName, "SymLinkCreate", "0") == "1" Then
                _SymLink_RemoveAppSymlinks($catIni, $appName, $settingsIni)
            EndIf
        Next
    Next
EndFunc

; Call this on tray exit!
Func _SymLink_TrayExitCleanup()
    If _SymLink_ShouldCleanupOnExit() Then
        MsgBox(64, "Symlink Cleanup", "Cleaning up per-app symlinks before exit...")
        _SymLink_CleanupAllPerAppSymlinks()
    EndIf
    ; Otherwise, just exit silently
EndFunc

; To use: call OnAutoItExitRegister("_SymLink_TrayExitCleanup") in your tray script
