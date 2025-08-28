; Favorites.au3
; Manages favorites logic (list/flag, menu build)

#include-once

Func _Favorites_IsFavorite($catIni, $appName)
    Return (IniRead($catIni, $appName, "Fave", "0") = "1")
EndFunc

Func _Favorites_GetFavorites($categories)
    Local $faveList = ObjCreate("Scripting.Dictionary")
    For $catName In $categories.Keys
        Local $catIni = @ScriptDir & "\App\" & $catName & ".ini"
        Local $catApps = $categories.Item($catName)
        For $appName In $catApps.Keys
            If _Favorites_IsFavorite($catIni, $appName) Then
                $faveList.Item($appName) = $catApps.Item($appName)
            EndIf
        Next
    Next
    Return $faveList
EndFunc
