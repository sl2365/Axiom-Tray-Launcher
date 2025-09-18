; Utils.au3
; Logging, error handling, helper functions

#include-once
#include "SymLinks.au3"
#include <File.au3>

Func _Utils_Log($msg)
    Local $logFile = @ScriptDir & "\App\debug.log"
    FileWrite($logFile, @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & $msg & @CRLF)
EndFunc

Func _Utils_Error($msg)
    MsgBox(16, "Error", $msg)
    _Utils_Log("ERROR: " & $msg)
EndFunc

Func _Utils_FileListToArrayEx($path)
    ; Returns array of file names in folder (non-hidden/system)
    Return _FileListToArray($path, "*", 1)
EndFunc

Func _Utils_KeysSorted($dict)
    Local $arr[ $dict.Count ]
    Local $i = 0
    For $k In $dict.Keys
        $arr[$i] = $k
        $i += 1
    Next
    _ArraySort($arr)
    Return $arr
EndFunc

Func RemoveExtraBlankLines($sFile)
    Local $txt = FileRead($sFile)
    If @error Or $txt = "" Then Return SetError(1,0,0)
    ; Split into lines
    Local $lines = StringSplit($txt, @CRLF, 1)
    If Not IsArray($lines) Then Return SetError(2,0,0)
    Local $out = ""
    Local $firstSection = True
    Local $previousLineBlank = True
    For $i = 1 To $lines[0]
        Local $line = $lines[$i]
        ; Section header?
        If StringRegExp($line, "^\[.+\]$") Then
            If Not $firstSection And Not $previousLineBlank Then $out &= @CRLF
            $firstSection = False
        EndIf
        ; Only one blank line at a time
        If $line = "" Then
            If Not $previousLineBlank Then $out &= @CRLF
            $previousLineBlank = True
        Else
            $out &= $line & @CRLF
            $previousLineBlank = False
        EndIf
    Next
    ; Ensure a blank line at end
    If StringRight($out, StringLen(@CRLF)) <> @CRLF Then $out &= @CRLF
    ; If last line isn't blank, add one
    Local $hFile = FileOpen($sFile, 2)
    If $hFile = -1 Then Return SetError(3,0,0)
    FileWrite($hFile, $out)
    FileClose($hFile)
    Return 1
EndFunc
