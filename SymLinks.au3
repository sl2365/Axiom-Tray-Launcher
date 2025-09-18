; SymLinks.au3

#include-once
Global Const $SS_CENTER = 0x0001

; --------- Symlink Variable Expansion ---------
Func _SymLink_ExpandVarsInPath($path, $vars)
    For $key In $vars.Keys()
        $path = StringReplace($path, "%" & $key & "%", $vars.Item($key))
    Next
    Local $re = StringRegExp($path, "%([A-Za-z0-9_]+)%", 3)
    If IsArray($re) Then
        For $i = 0 To UBound($re) - 1
            $path = StringReplace($path, "%" & $re[$i] & "%", EnvGet($re[$i]))
        Next
    EndIf
    Return $path
EndFunc

Func _SymLink_ExpandEnv($str)
    Local $re = StringRegExp($str, "%([A-Za-z0-9_]+)%", 3)
    If IsArray($re) Then
        For $i = 0 To UBound($re) - 1
            $str = StringReplace($str, "%" & $re[$i] & "%", EnvGet($re[$i]))
        Next
    EndIf
    Return $str
EndFunc

Func _SymLink_LoadVariables($iniFile)
    Local $vars = ObjCreate("Scripting.Dictionary")
    Local $section = "Variables"
    Local $varNames = IniReadSection($iniFile, $section)
    If IsArray($varNames) Then
        For $i = 1 To $varNames[0][0]
            $vars.Add($varNames[$i][0], _SymLink_ExpandEnv($varNames[$i][1]))
        Next
    EndIf
    Return $vars
EndFunc

Func _SymLink_IsSymlinkOrJunction($path)
    Local $cmd = 'fsutil reparsepoint query "' & $path & '"'
    Local $pid = Run(@ComSpec & " /c " & $cmd, "", @SW_HIDE, $STDOUT_CHILD)
    Local $output = ""
    While True
        $output &= StdoutRead($pid)
        If @error Then ExitLoop
    WEnd
    If StringInStr($output, "Symbolic Link") Or StringInStr($output, "Junction") Then Return True
    Return False
EndFunc

; --------- Batch Symlink Creation & Removal (PowerShell Version) ---------
Func _SymLink_BatchCreateSymlinks($symlink_specs)
    Local $settingsFile = @ScriptDir & "\Settings.ini"
    Local $symLinksType = IniRead($settingsFile, "SymLinks", "SymLinksType", "1")
    Local $itemType = ($symLinksType = "0") ? "Junction" : "SymbolicLink"
    Local $createdCount = 0
    For $i = 0 To UBound($symlink_specs) - 1
        Local $symlink_path = _ResolvePath($symlink_specs[$i][0], @ScriptDir)
        Local $target_path  = _ResolvePath($symlink_specs[$i][1], @ScriptDir)
        Local $symlinkParent = StringLeft($symlink_path, StringInStr($symlink_path, "\", 0, -1) - 1)
        If Not FileExists($symlinkParent) Then DirCreate($symlinkParent)
        If Not FileExists($symlink_path) Or Not _SymLink_IsSymlinkOrJunction($symlink_path) Then
            Local $psCmdArgs = '-Command "New-Item -ItemType {0} -Path ''{1}'' -Target ''{2}''"'
            $psCmdArgs = StringReplace($psCmdArgs, "{0}", $itemType)
            $psCmdArgs = StringReplace($psCmdArgs, "{1}", $symlink_path)
            $psCmdArgs = StringReplace($psCmdArgs, "{2}", $target_path)
            ShellExecute("powershell.exe", $psCmdArgs, "", "runas", @SW_MINIMIZE)
            $createdCount += 1
        EndIf
    Next
    If $createdCount > 0 Then
        MsgBox(64, "Symlinks/Junctions Created", $createdCount & " links processed.")
    Else
        MsgBox(48, "Symlinks/Junctions", "No new links to create." & @CRLF & "They already exist.")
    EndIf
EndFunc

Func _SymLink_BatchDeleteSymlinks($symlink_paths)
    Local $deletedCount = 0
    For $i = 0 To UBound($symlink_paths) - 1
        Local $symlink_path = _ResolvePath($symlink_paths[$i], @ScriptDir)
        If FileExists($symlink_path) And _SymLink_IsSymlinkOrJunction($symlink_path) Then
            Local $psDelArgs = '-Command "Remove-Item -Path ''{0}'' -Force"'
			$psDelArgs = StringReplace($psDelArgs, "{0}", $symlink_path)
			ShellExecute("powershell.exe", $psDelArgs, "", "runas", @SW_MINIMIZE)
            $deletedCount += 1
        EndIf
    Next
    If $deletedCount > 0 Then
        MsgBox(64, "Symlinks Removed", $deletedCount & " symlinks processed.")
    Else
        MsgBox(48, "Symlinks", "No symlinks found to delete.")
    EndIf
EndFunc

; --------- GLOBAL SYMLINKS ONLY ---------
Func _SymLink_CreateGlobalSymlinks($globalIni, $force = False)
    Local $doAdd = IniRead($globalIni, "SymLinks", "SymLinksAdd", "0")
    If (Not $force) And (StringStripWS($doAdd, 3) <> "1") Then Return
    Local $vars = _SymLink_LoadVariables($globalIni)
    Local $symlinks = IniReadSection($globalIni, "SymLinks")
    If Not IsArray($symlinks) Then Return

    ; Gather all specs for batch creation
    Local $specs[0][2]
    For $i = 1 To $symlinks[0][0]
        Local $key = $symlinks[$i][0]
        If StringLeft($key, 7) <> "SymLink" Then ContinueLoop
        Local $val = $symlinks[$i][1]
        Local $split = StringSplit($val, "~", 2)
        If UBound($split) = 2 Then
            Local $symlink_path = _SymLink_ExpandVarsInPath($split[0], $vars)
            Local $target_path  = _SymLink_ExpandVarsInPath($split[1], $vars)
            ReDim $specs[UBound($specs)+1][2]
            $specs[UBound($specs)-1][0] = $symlink_path
            $specs[UBound($specs)-1][1] = $target_path
        EndIf
    Next
    If UBound($specs) > 0 Then
        _SymLink_BatchCreateSymlinks($specs)
    Else
        MsgBox(48, "Symlinks", "No valid symlinks found to create.")
    EndIf
EndFunc

Func _SymLink_RemoveGlobalSymlinks($globalIni, $force = False)
    Local $doRemove = IniRead($globalIni, "SymLinks", "SymLinksRemove", "0")
    If (Not $force) And (StringStripWS($doRemove, 3) <> "1") Then Return
    Local $vars = _SymLink_LoadVariables($globalIni)
    Local $symlinks = IniReadSection($globalIni, "SymLinks")
    If Not IsArray($symlinks) Then Return

    ; Gather all paths for batch deletion
    Local $delArr[0]
    For $i = 1 To $symlinks[0][0]
        Local $key = $symlinks[$i][0]
        If StringLeft($key, 7) <> "SymLink" Then ContinueLoop
        Local $val = $symlinks[$i][1]
        Local $split = StringSplit($val, "~", 2)
        If UBound($split) = 2 Then
            Local $symlink_path = _SymLink_ExpandVarsInPath($split[0], $vars)
			$symlink_path = _ResolvePath($symlink_path, @ScriptDir)
            ReDim $delArr[UBound($delArr)+1]
            $delArr[UBound($delArr)-1] = $symlink_path
        EndIf
    Next
    If UBound($delArr) > 0 Then
        _SymLink_BatchDeleteSymlinks($delArr)
    Else
        MsgBox(48, "Symlinks", "No symlinks found to delete.")
    EndIf
EndFunc

Func _SymLink_ManualCreateGlobalSymlinks($globalIni)
    Local $result = MsgBox(33, "Manual Global Symlinks", "Create global symlinks now?" & @CRLF & "Click OK to proceed or Cancel to abort.")
    If $result = 1 Then _SymLink_CreateGlobalSymlinks($globalIni, True)
EndFunc

Func _SymLink_ManualRemoveGlobalSymlinks($globalIni)
    Local $result = MsgBox(33, "Manual Global Symlinks", "Remove global symlinks now?" & @CRLF & "Click OK to proceed or Cancel to abort.")
    If $result = 1 Then _SymLink_RemoveGlobalSymlinks($globalIni, True)
EndFunc

; --------- PER-APP SYMLINKS ONLY ---------
Func _SymLink_CreateAppSymlinks($catIni, $appName, $globalIni)
    Local $doAdd = IniRead($catIni, $appName, "SymLinkCreate", "0")
    If StringStripWS($doAdd, 3) <> "1" Then Return
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
    If Not IsArray($symlinks) Then Return
    Local $specs[0][2]
    For $i = 1 To $symlinks[0][0]
        Local $key = $symlinks[$i][0]
        If StringLeft($key, 7) <> "SymLink" Then ContinueLoop
        Local $val = $symlinks[$i][1]
        Local $split = StringSplit($val, "~", 2)
        If UBound($split) = 2 Then
            Local $target = _SymLink_ExpandVarsInPath($split[0], $vars)
            Local $link = _SymLink_ExpandVarsInPath($split[1], $vars)
            ReDim $specs[UBound($specs)+1][2]
            $specs[UBound($specs)-1][0] = $target
            $specs[UBound($specs)-1][1] = $link
        EndIf
    Next
    If UBound($specs) > 0 Then
        _SymLink_BatchCreateSymlinks($specs)
    Else
        MsgBox(48, "Symlinks", "No valid symlinks found to create for app.")
    EndIf
EndFunc

Func _SymLink_RemoveAppSymlinks($catIni, $appName, $globalIni = "")
    If StringStripWS(IniRead($catIni, $appName, "SymLinkCreate", "0"), 3) <> "1" Then Return
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
    Local $delArr[0]
    For $j = 1 To 10
        Local $symKey = "SymLink" & $j
        Local $symSpec = IniRead($catIni, $appName, $symKey, "")
        If $symSpec = "" Then ContinueLoop
        Local $parts = StringSplit($symSpec, "~", 2)
        If UBound($parts) = 2 Then
            Local $symlink_path = _SymLink_ExpandVarsInPath($parts[0], $vars)
            ReDim $delArr[UBound($delArr)+1]
            $delArr[UBound($delArr)-1] = $symlink_path
        EndIf
    Next
    If UBound($delArr) > 0 Then
        _SymLink_BatchDeleteSymlinks($delArr)
    Else
;~         MsgBox(48, "Symlinks", "No symlinks found to delete for app.")
    EndIf
EndFunc

; --------- Optional Debug ---------
Func _SymLink_DumpVariables($globalIni)
    Local $vars = _SymLink_LoadVariables($globalIni)
    For $key In $vars.Keys()
        ;ConsoleWrite($key & " = " & $vars.Item($key) & @CRLF)
    Next
EndFunc

; --------- Cleanup On Exit ---------
Global $settingsIni = @ScriptDir & "\App\Settings.ini"
Global $categoryIniDir = @ScriptDir & "\App"

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
                Return True
            EndIf
        Next
    Next
    Return False
EndFunc

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

Func _SymLink_TrayExitCleanup()
    If _SymLink_ShouldCleanupOnExit() Then
        MsgBox(64, "Symlink Cleanup", "Cleaning up per-app symlinks before exit...", 3)
        _SymLink_CleanupAllPerAppSymlinks()
    EndIf
EndFunc

; To use: call OnAutoItExitRegister("_SymLink_TrayExitCleanup") in your tray script
