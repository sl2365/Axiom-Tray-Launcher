; Config.au3
; Handles INI loading, validation, saving

#include "Utils.au3"

Func _Config_ReadIniSection($file, $section)
    Local $dict = ObjCreate("Scripting.Dictionary")
    Local $data = IniReadSection($file, $section)
    If IsArray($data) Then
        For $i = 1 To $data[0][0]
            $dict.Item($data[$i][0]) = $data[$i][1]
        Next
    EndIf
    Return $dict
EndFunc

Func _Config_LoadAndValidate($settingsFile)
    Local $settings = ObjCreate("Scripting.Dictionary")
    ; Create INI with defaults if missing
    If Not FileExists($settingsFile) Then
        _Config_WriteDefaults($settingsFile)
    EndIf
    ; Read and validate INI
    If Not IniReadSection($settingsFile, "GLOBAL") Then
        _Config_WriteDefaults($settingsFile)
    EndIf
    ; Load GLOBAL section
    Local $global = _Config_ReadIniSection($settingsFile, "GLOBAL")
    $settings.Item("GLOBAL") = $global
    ; Load ScannedPaths section
    Local $scanned = _Config_ReadIniSection($settingsFile, "ScannedPaths")
    $settings.Item("ScannedPaths") = $scanned
    ; Load Variables section
    Local $vars = _Config_ReadIniSection($settingsFile, "Variables")
    $settings.Item("Variables") = $vars
    Return $settings
EndFunc

Func _Config_WriteDefaults($settingsFile)
    If FileExists($settingsFile) Then
        ; File already exists, do not overwrite
        Return
    EndIf
    Local $text = _
        "[GLOBAL]" & @CRLF & _
        "Theme=0" & @CRLF & _
        "UpdateOnStart=0" & @CRLF & _
        "LastUpdateCheck=" & @CRLF & _
        "SandboxiePath=" & @CRLF & _
        @CRLF & _
        "[ScannedPaths]" & @CRLF & _
        "Scan1=C:\Windows" & @CRLF & _
        "Scan1Depth=0" & @CRLF & _
        "Scan1Ext=exe" & @CRLF & _
        "Scan2=" & @CRLF & _
        "Scan2Depth=0" & @CRLF & _
        "Scan2Ext=exe" & @CRLF & _
        @CRLF & _
        "[Variables]" & @CRLF & _
        "SetVar1=" & @CRLF & _
        "SetVar2=" & @CRLF & _
        @CRLF & _
        "[SymLinks]" & @CRLF & _
        "SymLinksAdd=" & @CRLF & _
        "SymLinksRemove=" & @CRLF & _
        "SymLink1=" & @CRLF & _
        "SymLink2=" & @CRLF

    Local $hFile = FileOpen($settingsFile, 2)
    If $hFile = -1 Then
        MsgBox(16, "Error", "Unable to create settings file: " & $settingsFile)
        Return
    EndIf
    FileWrite($hFile, $text)
    FileClose($hFile)
EndFunc

Func _Config_Save($settings, $settingsFile)
    For $section In $settings.Keys
        Local $list = $settings.Item($section)
        For $key In $list.Keys
            IniWrite($settingsFile, $section, $key, $list.Item($key))
        Next
    Next
EndFunc

Func _Config_GetScanExts($settings)
    ; Returns a dictionary mapping ScanN => their Exts
    Local $exts = ObjCreate("Scripting.Dictionary")
    Local $scanned = $settings.Item("ScannedPaths")
    For $key In $scanned.Keys
        If StringRegExp($key, "^Scan\d+Ext$") Then
            Local $scanN = StringLeft($key, StringLen($key) - 3) ; "ScanN" from "ScanNExt"
            $exts.Item($scanN) = $scanned.Item($key)
        EndIf
    Next
    Return $exts
EndFunc

Func _Config_GetScanPaths($settings)
    ; Returns a dictionary mapping ScanN => paths
    Local $paths = ObjCreate("Scripting.Dictionary")
    Local $scanned = $settings.Item("ScannedPaths")
    For $key In $scanned.Keys
        If StringRegExp($key, "^Scan\d+$") Then
            $paths.Item($key) = $scanned.Item($key)
        EndIf
    Next
    Return $paths
EndFunc

Func _Config_GetScanDepths($settings)
    ; Returns a dictionary mapping ScanN => depths
    Local $depths = ObjCreate("Scripting.Dictionary")
    Local $scanned = $settings.Item("ScannedPaths")
    For $key In $scanned.Keys
        If StringRegExp($key, "^Scan\d+Depth$") Then
            Local $scanN = StringLeft($key, StringLen($key) - 5) ; "ScanN" from "ScanNDepth"
            $depths.Item($scanN) = $scanned.Item($key)
        EndIf
    Next
    Return $depths
EndFunc
