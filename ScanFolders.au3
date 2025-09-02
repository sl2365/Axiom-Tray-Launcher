; ScanFolders.au3
; Scans folders from [ScannedPaths] in Settings.ini, discovers apps, creates per-category INIs, shows results.
; Uses path utility functions from TrayMenu.au3 and folder name util from Utils.au3.

#include-once
#include "TrayMenu.au3"
#include "Utils.au3"

; ---------------------- IgnoreList Support -------------------------
Global $g_IgnoreList = _ScanFolders_LoadIgnoreList()

Func _ScanFolders_LoadIgnoreList()
    Local $ignoreArr = FileReadToArray(_ResolvePath("App\IgnoreList.ini", @ScriptDir))
    Local $dict = ObjCreate("Scripting.Dictionary")
    If IsArray($ignoreArr) Then
        For $i = 0 To UBound($ignoreArr) - 1
            Local $entry = StringStripWS($ignoreArr[$i], 3)
            If $entry <> "" Then $dict.Item(StringLower($entry)) = True
        Next
    EndIf
    Return $dict
EndFunc

Func _ScanFolders_IsIgnored($fullPath)
    Local $resolvedFullPath = _ResolvePath($fullPath, @ScriptDir)
    Local $pathLower = StringLower($resolvedFullPath)
    Local $fileName = _GetFileName($fullPath)
    Local $fileLower = StringLower($fileName)
    Local $parentFolder = _GetParentFolder($fullPath)
    Local $parentPlusFile = ""
    If $parentFolder <> "" Then
        $parentPlusFile = StringLower($parentFolder & "\" & $fileName)
    EndIf
    ; Also check resolved parent+file
    Local $resolvedParentPlusFile = ""
    If $parentFolder <> "" Then
        $resolvedParentPlusFile = StringLower(_ResolvePath($parentFolder & "\" & $fileName, @ScriptDir))
    EndIf
    ; Check against all forms in ignore list
    If $g_IgnoreList.Exists($pathLower) Then Return True
    If $parentPlusFile <> "" And $g_IgnoreList.Exists($parentPlusFile) Then Return True
    If $resolvedParentPlusFile <> "" And $g_IgnoreList.Exists($resolvedParentPlusFile) Then Return True
    If $g_IgnoreList.Exists($fileLower) Then Return True
    Return False
EndFunc

Func _GetParentFolder($fullPath)
    Local $parts = StringSplit($fullPath, "\")
    If IsArray($parts) And $parts[0] > 1 Then
        Return $parts[$parts[0] - 1]
    EndIf
    Return ""
EndFunc

Func _MakePortablePath($path, $baseDir)
    ; Replace drive letter with ? if same as baseDir
    Local $driveBase = StringLeft($baseDir, 2)
    If StringLeft($path, 2) = $driveBase Then
        Return "?:" & StringTrimLeft($path, 2)
    EndIf
    Return $path
EndFunc

Func _LoadScanPathsFromINI($iniFile)
    Local $dict = ObjCreate("Scripting.Dictionary")
    Local $section = IniReadSection($iniFile, "ScannedPaths")
    If IsArray($section) And UBound($section) > 1 Then
        For $i = 1 To UBound($section) - 1
            $dict.Item($section[$i][0]) = $section[$i][1]
        Next
    EndIf
    Return $dict
EndFunc

Func _ScanFolders_Scan($settings)
    $g_IgnoreList = _ScanFolders_LoadIgnoreList() ; Reload ignore list at scan time for freshness
    Local $categories = ObjCreate("Scripting.Dictionary")
    Local $scanned = _LoadScanPathsFromINI(_ResolvePath("App\Settings.ini", @ScriptDir))
    Local $catNames = ObjCreate("Scripting.Dictionary")
    Local $catCounter = ObjCreate("Scripting.Dictionary")

    For $key In $scanned.Keys
        If StringRegExp($key, "^Scan\d+$") And $scanned.Item($key) <> "" Then
            Local $scanNum = StringTrimLeft($key, 4)
            Local $scanPath = _ResolvePath($scanned.Item($key), @ScriptDir)
            Local $depthKey = "Scan" & $scanNum & "Depth"
            Local $extKey = "Scan" & $scanNum & "Ext"
            Local $depth = 1
            If $scanned.Exists($depthKey) Then $depth = Number($scanned.Item($depthKey))
            If $depth > 10 Then $depth = 10
            If $depth < 0 Then $depth = 0
            Local $filters = "exe"
            If $scanned.Exists($extKey) And $scanned.Item($extKey) <> "" Then $filters = $scanned.Item($extKey)
            Local $catName = _GetFolderName($scanPath)
            If $catNames.Exists($catName) Then
                $catCounter.Item($catName) += 1
                $catName &= $catCounter.Item($catName)
            Else
                $catNames.Item($catName) = 1
                $catCounter.Item($catName) = 1
            EndIf
            Local $apps = _ScanFolders_DiscoverApps($scanPath, $depth, $filters)
            $categories.Item($catName) = $apps
            Local $catIni = _ResolvePath("App\" & $catName & ".ini", @ScriptDir)
            _ScanFolders_WriteOrUpdateCategoryIni($catIni, $apps, $catName)
        EndIf
    Next
    Return $categories
EndFunc

Func _ScanFolders_ScanSingle($scanPath, $depth, $filters, $catName)
    $g_IgnoreList = _ScanFolders_LoadIgnoreList() ; Reload ignore list for single scan
    Local $apps = _ScanFolders_DiscoverApps($scanPath, $depth, $filters)
    Local $newApps = 0
    Local $catIni = _ResolvePath("App\" & $catName & ".ini", @ScriptDir)
    If Not FileExists($catIni) Then
        _ScanFolders_WriteOrUpdateCategoryIni($catIni, $apps, $catName)
        $newApps = $apps.Count
    Else
        $newApps = 0
        For $appName In $apps.Keys
            If IniRead($catIni, $appName, "RunFile", "") == "" Then $newApps += 1
        Next
        _ScanFolders_WriteOrUpdateCategoryIni($catIni, $apps, $catName)
    EndIf
    Return $newApps
EndFunc

Func _ScanFolders_AndShowResults($settings)
    $g_IgnoreList = _ScanFolders_LoadIgnoreList() ; Reload ignore list for scan
    Local $scanned = _LoadScanPathsFromINI(_ResolvePath("App\Settings.ini", @ScriptDir))
    Local $results = ""
    Local $notFound = ""
    Local $catNames = ObjCreate("Scripting.Dictionary")
    Local $catCounter = ObjCreate("Scripting.Dictionary")
    Local $pathIndex = 1
    Local $notFoundIndex = 1

    For $key In $scanned.Keys
        If StringRegExp($key, "^Scan\d+$") And $scanned.Item($key) <> "" Then
            Local $scanNum = StringTrimLeft($key, 4)
            Local $scanPath = _ResolvePath($scanned.Item($key), @ScriptDir)
            Local $depthKey = "Scan" & $scanNum & "Depth"
            Local $extKey = "Scan" & $scanNum & "Ext"
            Local $depth = 1
            If $scanned.Exists($depthKey) Then $depth = Number($scanned.Item($depthKey))
            If $depth > 10 Then $depth = 10
            If $depth < 0 Then $depth = 0
            Local $filters = "exe"
            If $scanned.Exists($extKey) And $scanned.Item($extKey) <> "" Then $filters = $scanned.Item($extKey)
            Local $catName = _GetFolderName($scanPath)
            If $catNames.Exists($catName) Then
                $catCounter.Item($catName) += 1
                $catName &= $catCounter.Item($catName)
            Else
                $catNames.Item($catName) = 1
                $catCounter.Item($catName) = 1
            EndIf
            If FileExists($scanPath) Then
                Local $numNewApps = _ScanFolders_ScanSingle($scanPath, $depth, $filters, $catName)
                $results &= "Path " & $pathIndex & " - " & $scanPath & @CRLF
                $results &= "           - " & $numNewApps & " new apps found." & @CRLF & @CRLF
                $pathIndex += 1
            Else
                $catNames.Item("notfound_" & $notFoundIndex) = $scanPath
                $catCounter.Item("notfound_" & $notFoundIndex) = $scanNum
                $notFoundIndex += 1
            EndIf
        EndIf
    Next
    If $notFoundIndex > 1 Then
        $notFound &= "Folders not found:" & @CRLF
        For $i = 1 To $notFoundIndex - 1
            Local $nfPath = $catNames.Item("notfound_" & $i)
            $notFound &= "Path " & $i & " - " & $nfPath & @CRLF
            $notFound &= "           - Folder not found." & @CRLF & @CRLF
        Next
    EndIf
    Local $msg = "Scanned folders:" & @CRLF & $results
    If $notFound <> "" Then $msg &= @CRLF & $notFound
    MsgBox(64, "Scan Complete", $msg)
EndFunc

Func _ScanFolders_DiscoverApps($scanPath, $depth, $filters)
    Local $apps = ObjCreate("Scripting.Dictionary")
    _ScanFolders_ScanFolderNative($scanPath, $apps, $depth, 0, $filters)
    Return $apps
EndFunc

Func _ScanFolders_ScanFolderNative($folder, ByRef $apps, $maxDepth, $curDepth, $filters)
    If $curDepth > $maxDepth Or $maxDepth < 0 Then Return

    Local $search = FileFindFirstFile($folder & "\*")
    If $search = -1 Then Return

    While 1
        Local $file = FileFindNextFile($search)
        If @error Then ExitLoop
        Local $fullPath = $folder & "\" & $file
        Local $attrib = FileGetAttrib($fullPath)
        ; ----------- Flexible ignore check -------------
        If _ScanFolders_IsIgnored($fullPath) Then
            ; ConsoleWrite("IGNORED: " & $fullPath & @CRLF)
            ContinueLoop
        EndIf
        ; ----------- End ignore check ------------------
        If StringInStr($attrib, "D") Then
            If $file <> "." And $file <> ".." And $curDepth < $maxDepth _
                And Not StringInStr($attrib, "H") _
                And Not StringInStr($attrib, "S") Then
                _ScanFolders_ScanFolderNative($fullPath, $apps, $maxDepth, $curDepth + 1, $filters)
            EndIf
        Else
            If Not StringInStr($attrib, "H") And Not StringInStr($attrib, "S") Then
                If StringInStr($file, ".") Then
                    Local $ext = StringLower(StringTrimLeft($file, StringInStr($file, ".", 0, -1)))
                    Local $filtersArr = StringSplit($filters, ",")
                    For $f = 1 To $filtersArr[0]
                        If $ext = StringStripWS($filtersArr[$f], 3) Then
                            Local $fileName = _GetFileName($fullPath)
                            Local $baseNameNoExt = StringRegExpReplace($fileName, "\.[^.]*$", "")
                            Local $baseName = StringReplace($baseNameNoExt, "Portable", "")
                            $apps.Item($baseName) = $fullPath
                            ExitLoop
                        EndIf
                    Next
                EndIf
            EndIf
        EndIf
    WEnd
    FileClose($search)
EndFunc

Func _ScanFolders_WriteOrUpdateCategoryIni($catIni, $apps, $catName)
    Local $existingSections = IniReadSectionNames($catIni)
    Local $existingDict = ObjCreate("Scripting.Dictionary")
    If IsArray($existingSections) Then
        For $i = 1 To $existingSections[0]
            $existingDict.Item($existingSections[$i]) = 1
        Next
    EndIf
    Local $hFile = FileOpen($catIni, 1)
    If $hFile = -1 Then
        $hFile = FileOpen($catIni, 2)
        If $hFile = -1 Then
            MsgBox(16, "Error", "Unable to create category ini file: " & $catIni)
            Return
        EndIf
    EndIf
    For $appName In $apps.Keys
        If Not $existingDict.Exists($appName) Then
            FileWrite($hFile, "[" & $appName & "]" & @CRLF)
            FileWrite($hFile, "ButtonText=" & $appName & @CRLF)
            FileWrite($hFile, "RunFile=" & _MakePortablePath($apps.Item($appName), @ScriptDir) & @CRLF)
            FileWrite($hFile, "RunAsAdmin=0" & @CRLF)
            FileWrite($hFile, "WorkDir=" & @CRLF)
            FileWrite($hFile, "Arguments=" & @CRLF)
            FileWrite($hFile, "SingleInstance=0" & @CRLF)
            FileWrite($hFile, "Sandboxie=0" & @CRLF)
            FileWrite($hFile, "SandboxName=" & @CRLF)
            FileWrite($hFile, "Category=" & $catName & @CRLF)
            FileWrite($hFile, "SymLinkCreate=" & @CRLF)
            FileWrite($hFile, "SymLink1=" & @CRLF)
            FileWrite($hFile, "Fave=0" & @CRLF)
            FileWrite($hFile, "Hide=0" & @CRLF)
            FileWrite($hFile, "SetEnv1=" & @CRLF & @CRLF)
        EndIf
    Next
    FileClose($hFile)
EndFunc

Func _ScanFolders_GetApps($categories)
    Local $apps = ObjCreate("Scripting.Dictionary")
    Local $faveArr[0]
    Local $faveCount = 0
    For $catName In $categories.Keys
        Local $catApps = $categories.Item($catName)
        Local $arr[ $catApps.Count ]
        Local $i = 0
        For $appName In $catApps.Keys
            $arr[$i] = $appName
            $i += 1
        Next
        $apps.Item($catName) = $arr
        Local $catIni = _ResolvePath("App\" & $catName & ".ini", @ScriptDir)
        If FileExists($catIni) Then
            For $appName In $catApps.Keys
                If IniRead($catIni, $appName, "Fave", "0") = "1" Then
                    ReDim $faveArr[$faveCount+1]
                    $faveArr[$faveCount] = $appName
                    $faveCount += 1
                EndIf
            Next
        EndIf
    Next
    If $faveCount > 0 Then
        $apps.Item("Fave") = $faveArr
    EndIf
    Return $apps
EndFunc

Func _ScanFolders_LoadCategoriesFromIni()
    ; Returns a dictionary of categories, each containing a dictionary of apps
    Local $categories = ObjCreate("Scripting.Dictionary")
    Local $appFolder = _ResolvePath("App", @ScriptDir)
    Local $fileList = _FileListToArray($appFolder, "*.ini", 1)
    If IsArray($fileList) Then
        For $i = 1 To $fileList[0]
            Local $file = $fileList[$i]
            Local $fileLower = StringLower($file)
            ; Filter out global config files (case-insensitive)
            If $fileLower = "settings.ini" Or $fileLower = "ignorelist.ini" Then ContinueLoop
            ; Only process category INIs
            Local $catName = StringTrimRight($file, 4) ; Remove ".ini"
            Local $catIni = $appFolder & "\" & $file
            Local $catApps = ObjCreate("Scripting.Dictionary")
            Local $sections = IniReadSectionNames($catIni)
            If IsArray($sections) Then
                For $j = 1 To $sections[0]
                    Local $appName = $sections[$j]
                    Local $runFile = IniRead($catIni, $appName, "RunFile", "")
                    If $runFile <> "" Then
                        $catApps.Item($appName) = $runFile
                    EndIf
                Next
            EndIf
            $categories.Item($catName) = $catApps
        Next
    EndIf
    Return $categories
EndFunc

Func _ScanFolders_GetAppsWithButtonText($categories)
    Local $apps = ObjCreate("Scripting.Dictionary")
    Local $faveArr[0]
    Local $faveCount = 0
    For $catName In $categories.Keys
        Local $catApps = $categories.Item($catName)
        ; Count non-fave apps for this category
        Local $nonFaveCount = 0
        Local $catIni = _ResolvePath("App\" & $catName & ".ini", @ScriptDir)
        If FileExists($catIni) Then
            ; First pass to count non-fave apps
            For $appName In $catApps.Keys
                If IniRead($catIni, $appName, "Hide", "0") = "1" Then ContinueLoop
                If IniRead($catIni, $appName, "Fave", "0") = "1" Then
                    ReDim $faveArr[$faveCount+1]
                    $faveArr[$faveCount] = $appName
                    $faveCount += 1
                    ContinueLoop ; <-- Do NOT include in category
                EndIf
                $nonFaveCount += 1
            Next
            ; Second pass to build category array
            Local $arr[$nonFaveCount][2]
            Local $i = 0
            For $appName In $catApps.Keys
                If IniRead($catIni, $appName, "Hide", "0") = "1" Then ContinueLoop
                If IniRead($catIni, $appName, "Fave", "0") = "1" Then ContinueLoop
                Local $buttonText = IniRead($catIni, $appName, "ButtonText", $appName)
                $arr[$i][0] = $appName
                $arr[$i][1] = $buttonText
                $i += 1
            Next
            $apps.Item($catName) = $arr
        EndIf
    Next
    If $faveCount > 0 Then
        $apps.Item("Fave") = $faveArr
    EndIf
    Return $apps
EndFunc
