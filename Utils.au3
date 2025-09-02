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
