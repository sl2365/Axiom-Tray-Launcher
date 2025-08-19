#include-once
#include <Array.au3>
#include <WinAPIFiles.au3>
#include <FileConstants.au3>
#include <WinAPI.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>

; ----------------- Utilities -----------------
Func _ResolvePath($path)
    If $path = "" Then Return ""
    If StringRegExp($path, "^[A-Za-z]:\\") Then Return $path
    Return _WinAPI_GetFullPathName(@ScriptDir & "\" & $path)
EndFunc

Func _IsDirectory($path)
    Local $attr = FileGetAttrib($path)
    If @error Or $attr = "" Then Return False
    Return StringInStr($attr, "D") > 0
EndFunc

Func _DirOf($path)
    If $path = "" Then Return @ScriptDir
    Local $p = StringReplace($path, "/", "\")
    Local $i = StringInStr($p, "\", 0, -1)
    If $i > 0 Then Return StringLeft($p, $i - 1)
    Return @ScriptDir
EndFunc

; ----------------- Settings bootstrap -----------------
Func _ResourceGetAsString($resName)
    Local $hExe = _WinAPI_GetModuleHandle(0)
    Local $hRes = DllCall("kernel32.dll", "ptr", "FindResourceW", "ptr", $hExe, "wstr", "CUSTOM", "wstr", "SETTINGSTEMPLATE.INI")
    If @error Or $hRes[0] = 0 Then Return SetError(1, 0, "")
    Local $hGlobal = DllCall("kernel32.dll", "ptr", "LoadResource", "ptr", $hExe, "ptr", $hRes[0])
    If @error Or $hGlobal[0] = 0 Then Return SetError(2, 0, "")
    Local $pData = DllCall("kernel32.dll", "ptr", "LockResource", "ptr", $hGlobal[0])
    If @error Or $pData[0] = 0 Then Return SetError(3, 0, "")
    Local $size = DllCall("kernel32.dll", "dword", "SizeofResource", "ptr", $hExe, "ptr", $hRes[0])
    If @error Or $size[0] = 0 Then Return SetError(4, 0, "")
    Local $struct = DllStructCreate("byte[" & $size[0] & "]", $pData[0])
    Return BinaryToString(DllStructGetData($struct, 1))
EndFunc

;~ Func EnsureIniExists()
;~     If Not FileExists($INI_FILE) Then
;~         Local $template = _ResourceGetAsString("CUSTOM")
;~         If @error Or $template = "" Then
;~             MsgBox(16, "Error", "Embedded SettingsTemplate.ini not found.")
;~             Exit
;~         EndIf
;~         Local $hFile = FileOpen($INI_FILE, $FO_OVERWRITE + $FO_CREATEPATH)
;~         If $hFile = -1 Then
;~             MsgBox(16, "Error", "Failed to create " & $INI_FILE)
;~             Exit
;~         EndIf
;~         FileWrite($hFile, $template)
;~         FileClose($hFile)
;~     EndIf
;~ EndFunc

;~ Func EnsureIniDefaults()
;~     Local $TEMPLATE_INI = @ScriptDir & "\App\SettingsTemplate.ini"
;~     Local $sections = IniReadSectionNames($TEMPLATE_INI)
;~     If @error Or Not IsArray($sections) Then Return
;~     For $i = 1 To $sections[0]
;~         Local $templateKeys = IniReadSection($TEMPLATE_INI, $sections[$i])
;~         If @error Or Not IsArray($templateKeys) Then ContinueLoop
;~         For $j = 1 To $templateKeys[0][0]
;~             Local $key = $templateKeys[$j][0]
;~             Local $defval = $templateKeys[$j][1]
;~             Local $existing = IniRead($INI_FILE, $sections[$i], $key, "")
;~             If $existing = "" Then IniWrite($INI_FILE, $sections[$i], $key, $defval)
;~         Next
;~     Next
;~ EndFunc

Func _GetTemplateContent()
    If @Compiled Then
        Local $template = _ResourceGetAsString("CUSTOM")
        If @error Or $template = "" Then
            MsgBox(16, "Error", "Embedded SettingsTemplate.ini not found.")
            Exit
        EndIf
        Return $template
    Else
        Local $template = FileRead(@ScriptDir & "\SettingsTemplate.ini")
        If @error Or $template = "" Then
            MsgBox(16, "Error", "SettingsTemplate.ini file not found in script folder.")
            Exit
        EndIf
        Return $template
    EndIf
EndFunc

Func EnsureIniExists()
    If Not FileExists($INI_FILE) Then
        Local $template = _GetTemplateContent()
        Local $hFile = FileOpen($INI_FILE, $FO_OVERWRITE + $FO_CREATEPATH)
        If $hFile = -1 Then
            MsgBox(16, "Error", "Failed to create " & $INI_FILE)
            Exit
        EndIf
        FileWrite($hFile, $template)
        FileClose($hFile)
    EndIf
EndFunc

Func ReadSettings()
    $CloseMenuOnClick = Number(IniRead($INI_FILE, "GLOBAL", "CloseMenuOnClick", "0"))
    $SandboxieConfigured = False
    $SandboxiePath = IniRead($INI_FILE, "GLOBAL", "SandboxiePath", "")
    Local $sbxResolved = _ResolvePath($SandboxiePath)
    If $SandboxiePath <> "" And (FileExists($sbxResolved) Or _IsDirectory($sbxResolved)) Then $SandboxieConfigured = True

    $userVars.RemoveAll()
    Local $globalKeys = IniReadSection($INI_FILE, "GLOBAL")
    If @error Or Not IsArray($globalKeys) Then Return
    For $i = 1 To $globalKeys[0][0]
        If StringLower($globalKeys[$i][0]) <> "setpath" Then
            $userVars($globalKeys[$i][0]) = $globalKeys[$i][1]
        Else
            Local $parts = StringSplit($globalKeys[$i][1], "|")
            If $parts[0] = 2 Then $userVars($parts[1]) = $parts[2]
        EndIf
    Next
EndFunc

Func ExpandUserVars($str)
    Local $result = $str
    For $key In $userVars.Keys
        $result = StringReplace($result, "%" & $key & "%", $userVars($key))
    Next
    Return $result
EndFunc

Func ReadButtonSettings($iniFile, $section)
    Local $dict = ObjCreate("Scripting.Dictionary")
    Local $keys = IniReadSection($iniFile, $section)
    If @error Or Not IsArray($keys) Or UBound($keys) < 2 Then Return $dict
    For $i = 1 To $keys[0][0]
        $dict($keys[$i][0]) = $keys[$i][1]
    Next
    Return $dict
EndFunc

Func ReadApps()
    Local $sections = IniReadSectionNames($INI_FILE)
    Local $btns[0]
    For $i = 1 To $sections[0]
        If StringRegExp($sections[$i], "(?i)^button\s*\d+\b") Then _ArrayAdd($btns, $sections[$i])
    Next

    If UBound($btns) = 0 Then
        ReDim $appsData[1][12]
        ReDim $appsSections[1]
        $appsData[0][$APP_NAME] = ""
        $appsSections[0] = ""
        Return
    EndIf

    ReDim $appsData[UBound($btns)][12]
    ReDim $appsSections[UBound($btns)]

    For $idx = 0 To UBound($btns) - 1
        Local $section = $btns[$idx]
        Local $dict = ReadButtonSettings($INI_FILE, $section)
        $appsSections[$idx] = $section
        Local $idStr = StringRegExpReplace($section, "[^\d]", "")
        Local $idNum = (StringIsInt($idStr) ? Number($idStr) : 0)

        $appsData[$idx][$APP_NAME]     = $dict.Exists("ButtonText") ? $dict.Item("ButtonText") : ("Button " & $idNum)
        $appsData[$idx][$APP_CAT]      = $dict.Exists("Category") ? StringStripWS($dict.Item("Category"), 3) : ""
        $appsData[$idx][$APP_PATH]     = $dict.Exists("RunFile") ? ExpandUserVars($dict.Item("RunFile")) : ""
        $appsData[$idx][$APP_ARGS]     = $dict.Exists("Arguments") ? ExpandUserVars($dict.Item("Arguments")) : ""
        $appsData[$idx][$APP_SINGLE]   = $dict.Exists("SingleInstance") ? (StringStripWS($dict.Item("SingleInstance"), 3) = "1") : False
        $appsData[$idx][$APP_ADMIN]    = $dict.Exists("RunAsAdmin") ? Number($dict.Item("RunAsAdmin")) : 0
        $appsData[$idx][$APP_NET]      = $dict.Exists("NetAccess") ? Number($dict.Item("NetAccess")) : 1
        $appsData[$idx][$APP_SBX_PATH] = $dict.Exists("Sandboxie") ? ExpandUserVars($dict.Item("Sandboxie")) : ""
        $appsData[$idx][$APP_SBX_NAME] = $dict.Exists("SandboxName") ? ExpandUserVars($dict.Item("SandboxName")) : ""
        $appsData[$idx][$APP_ASSOC]    = $dict.Exists("FileAssocApp") ? ExpandUserVars($dict.Item("FileAssocApp")) : ""
        $appsData[$idx][$APP_HIDE]     = "0"
        $appsData[$idx][$APP_ID]       = $idNum
    Next
EndFunc

Func SaveAppsList($appsArray)
    If FileExists($PORTABLE_LIST_FILE) Then FileDelete($PORTABLE_LIST_FILE)
    For $i = 0 To UBound($appsArray) - 1
        IniWrite($PORTABLE_LIST_FILE, $appsArray[$i][$APP_NAME], "Category", $appsArray[$i][$APP_CAT])
        IniWrite($PORTABLE_LIST_FILE, $appsArray[$i][$APP_NAME], "Path", $appsArray[$i][$APP_PATH])
        IniWrite($PORTABLE_LIST_FILE, $appsArray[$i][$APP_NAME], "Arguments", $appsArray[$i][$APP_ARGS])
        Local $hideVal = ($appsArray[$i][$APP_HIDE] == "") ? "0" : $appsArray[$i][$APP_HIDE]
        IniWrite($PORTABLE_LIST_FILE, $appsArray[$i][$APP_NAME], "Hide", $hideVal)
    Next
EndFunc

Func LoadPortableAppsList()
    If Not FileExists($PORTABLE_LIST_FILE) Then Return
    Local $sections = IniReadSectionNames($PORTABLE_LIST_FILE)
    If @error Or Not IsArray($sections) Then Return
    For $i = 1 To $sections[0]
        Local $appName = $sections[$i]
        Local $path = IniRead($PORTABLE_LIST_FILE, $appName, "Path", "")
        Local $args = IniRead($PORTABLE_LIST_FILE, $appName, "Arguments", "")
        Local $hide = IniRead($PORTABLE_LIST_FILE, $appName, "Hide", "0")
        Local $singleInstanceStr = IniRead($PORTABLE_LIST_FILE, $appName, "SingleInstance", "0")
        Local $singleInstance = (StringStripWS($singleInstanceStr, 3) = "1")
        If $hide <> "0" Then ContinueLoop

        Local $alreadyPresent = False
        For $j = 0 To UBound($appsData) - 1
            If StringLower($appsData[$j][$APP_PATH]) = StringLower($path) Then
                $alreadyPresent = True
                ExitLoop
            EndIf
        Next

        If Not $alreadyPresent And $path <> "" Then
            Local $prevCount = UBound($appsData)
            ReDim $appsData[$prevCount + 1][12]
            $appsData[$prevCount][$APP_NAME]     = $appName
            $appsData[$prevCount][$APP_CAT]      = "PortableApps"
            $appsData[$prevCount][$APP_PATH]     = $path
            $appsData[$prevCount][$APP_ARGS]     = $args
            $appsData[$prevCount][$APP_SINGLE]   = $singleInstance
            $appsData[$prevCount][$APP_ADMIN]    = 0
            $appsData[$prevCount][$APP_NET]      = 1
            $appsData[$prevCount][$APP_SBX_PATH] = ""
            $appsData[$prevCount][$APP_SBX_NAME] = ""
            $appsData[$prevCount][$APP_ASSOC]    = ""
            $appsData[$prevCount][$APP_HIDE]     = $hide
            $appsData[$prevCount][$APP_ID]       = 0
            ReDim $appsSections[$prevCount + 1]
            $appsSections[$prevCount] = ""
        EndIf
    Next
EndFunc

Func LoadScannedApps()
    Local $search = FileFindFirstFile(@ScriptDir & "\App\*.ini")
    If $search = -1 Then Return
    While 1
        Local $file = FileFindNextFile($search)
        If @error Then ExitLoop
        If $file = "Settings.ini" Or $file = "PortableAppsList.ini" Then ContinueLoop
        Local $iniPath = @ScriptDir & "\App\" & $file
        Local $sections = IniReadSectionNames($iniPath)
        If @error Or Not IsArray($sections) Then ContinueLoop
        For $i = 1 To $sections[0]
            Local $appName = $sections[$i]
            Local $path = IniRead($iniPath, $appName, "Path", "")
            Local $args = IniRead($iniPath, $appName, "Arguments", "")
            Local $hide = IniRead($iniPath, $appName, "Hide", "0")
            Local $category = IniRead($iniPath, $appName, "Category", StringRegExpReplace($file, "\.ini$", ""))
            If $hide <> "0" Or $path = "" Then ContinueLoop

            ; === PATCH START: Check for duplicate by path ===
            Local $alreadyPresent = False
            For $j = 0 To UBound($appsData) - 1
                If StringLower($appsData[$j][$APP_PATH]) = StringLower($path) Then
                    $alreadyPresent = True
                    ExitLoop
                EndIf
            Next
            If $alreadyPresent Then ContinueLoop ; skip duplicate
            ; === PATCH END ===

            Local $prevCount = UBound($appsData)
            ReDim $appsData[$prevCount + 1][12]
            $appsData[$prevCount][$APP_NAME]     = $appName
            $appsData[$prevCount][$APP_CAT]      = $category
            $appsData[$prevCount][$APP_PATH]     = $path
            $appsData[$prevCount][$APP_ARGS]     = $args
            $appsData[$prevCount][$APP_SINGLE]   = False
            $appsData[$prevCount][$APP_ADMIN]    = 0
            $appsData[$prevCount][$APP_NET]      = 1
            $appsData[$prevCount][$APP_SBX_PATH] = ""
            $appsData[$prevCount][$APP_SBX_NAME] = ""
            $appsData[$prevCount][$APP_ASSOC]    = ""
            $appsData[$prevCount][$APP_HIDE]     = $hide
            $appsData[$prevCount][$APP_ID]       = 0
            ReDim $appsSections[$prevCount + 1]
            $appsSections[$prevCount] = ""
        Next
    WEnd
    FileClose($search)
EndFunc

; ----------------- Updated Scan Logic -----------------
Func ScanAppsFolders()
    Local $index = 1
    Local $scanSummary = "Scanned folders:" & @CRLF
    Local $foundAppsTotal = 0
    Local $missingFolders = "" ; Collect missing folders

    While True
        Local $scanKey = "Scan" & $index
        Local $depthKey = "Scan" & $index & "Depth"
        Local $scanPath = IniRead($INI_FILE, "GLOBAL", $scanKey, "")
        If $scanPath = "" Then ExitLoop
        Local $scanDepth = IniRead($INI_FILE, "GLOBAL", $depthKey, "1")
        $scanPath = _ResolvePath($scanPath)
        If Not FileExists($scanPath) Or Not _IsDirectory($scanPath) Then
            $missingFolders &= "Path " & $index & " - " & $scanPath & @CRLF
            $index += 1
            ContinueLoop
        EndIf

        Local $folderName = StringRegExpReplace($scanPath, "^.*\\([^\\]+)\\?$", "$1")
        Local $outputIni = @ScriptDir & "\App\" & $folderName & ".ini"

        Local $appsFound = _ScanFolderWriteIni($scanPath, Number($scanDepth), $folderName, $outputIni)
        $scanSummary &= "Path " & $index & " - " & $scanPath & @CRLF
        $scanSummary &= "           - " & $appsFound & " new apps found." & @CRLF & @CRLF
        $foundAppsTotal += $appsFound

        $index += 1
    WEnd

    ; Build the summary message
    Local $msg = ""
    $msg &= $scanSummary & "Total new apps found: " & $foundAppsTotal & @CRLF
    If $missingFolders <> "" Then
        $msg &= @CRLF & "Folders not found:" & @CRLF & $missingFolders
    EndIf

    MsgBox(64, "Scan Complete", $msg)
EndFunc

Func _ScanFolderWriteIni($folder, $maxDepth, $defaultCategory, $outputIni, $currentDepth = 1)
    Local $appsFound = 0
    Local $search = FileFindFirstFile($folder & "\*")
    If $search = -1 Then Return 0
    While 1
        Local $item = FileFindNextFile($search)
        If @error Then ExitLoop
        If $item = "." Or $item = ".." Then ContinueLoop
        Local $fullPath = $folder & "\" & $item
        If _IsDirectory($fullPath) Then
            If $currentDepth < $maxDepth Then
                $appsFound += _ScanFolderWriteIni($fullPath, $maxDepth, $defaultCategory, $outputIni, $currentDepth + 1)
            EndIf
        ElseIf StringRight($item, 4) = ".exe" Then
            Local $appName = StringRegExpReplace($item, "(?i)Portable", "")
            $appName = StringRegExpReplace($appName, "\.exe$", "")
            $appName = StringStripWS($appName, 3)
            ; Check ALL sections for duplicate Path
            Local $sections = IniReadSectionNames($outputIni)
            Local $alreadyExists = False
            If IsArray($sections) Then
                For $i = 1 To $sections[0]
                    Local $existingPath = IniRead($outputIni, $sections[$i], "Path", "")
                    If StringLower($existingPath) = StringLower($fullPath) Then
                        $alreadyExists = True
                        ExitLoop
                    EndIf
                Next
            EndIf
            If $alreadyExists Then ContinueLoop ; Skip duplicates by Path
            ; Add new app
            IniWrite($outputIni, $appName, "Category", $defaultCategory)
            IniWrite($outputIni, $appName, "Path", $fullPath)
            IniWrite($outputIni, $appName, "Arguments", "")
            IniWrite($outputIni, $appName, "Hide", "0")
            $appsFound += 1
        EndIf
    WEnd
    FileClose($search)
    Return $appsFound
	TrayUI_Destroy()
    TrayUI_BuildTrayMenu()
EndFunc

; ----------------- IPC (no-file IPC via registered message) -----------------
Func _IPC_AllowMessageFromLowerIntegrity($hWnd, $msg)
    Local Const $MSGFLT_ALLOW = 1
    Local $filter = DllStructCreate("dword cbSize; dword ExtStatus")
    DllStructSetData($filter, 1, DllStructGetSize($filter))
    DllCall("user32.dll", "bool", "ChangeWindowMessageFilterEx", _
        "hwnd", $hWnd, "uint", $msg, "dword", $MSGFLT_ALLOW, "ptr", DllStructGetPtr($filter))
EndFunc

Func IPC_RunBtn_Handler($hWnd, $uMsg, $wParam, $lParam)
;~ 	MsgBox(0, "Handler", "Message received: " & $uMsg & @CRLF & "g_WM_RUNBTN = " & $g_WM_RUNBTN)
    If $uMsg = $g_WM_RUNBTN Then
        Local $id = Number($wParam)
        ; DEBUG: Confirm message receipt
;~         MsgBox(0, "IPC", "Received IPC for ID: " & $id)
        If $id > 0 Then
            For $i = 0 To UBound($appsData) - 1
                If $appsData[$i][$APP_ID] = $id Then
                    RunButtonById($id)
                    ExitLoop
                EndIf
            Next
        EndIf
        Return 0
    EndIf
    Return 0
EndFunc

Func IPC_Init()
;~ 	MsgBox(0, "IPC_Init", "IPC Window Created")
    If $g_WM_RUNBTN = 0 Then
        $g_WM_RUNBTN = _WinAPI_RegisterWindowMessage("AxiomTrayLauncher.RunButton")
    EndIf
    Local $hGui = GUICreate($IPC_WINDOW_TITLE, 10, 10, -1, -1, $WS_POPUP, $WS_EX_TOOLWINDOW)
    _IPC_AllowMessageFromLowerIntegrity($hGui, $g_WM_RUNBTN)
    GUIRegisterMsg($g_WM_RUNBTN, "IPC_RunBtn_Handler")
    GUISetState(@SW_HIDE, $hGui)
;~     If $DEBUG_IPC Then MsgBox(64, "IPC", "IPC ready. WM=" & $g_WM_RUNBTN)
    Return $hGui
EndFunc

Func IPC_PostRunBtn($id)
    If $id <= 0 Then Return SetError(1, 0, False)
    Local $hWnd = WinGetHandle($IPC_WINDOW_TITLE)
    If $hWnd = 0 Then
        WinWait($IPC_WINDOW_TITLE, "", 2)
        $hWnd = WinGetHandle($IPC_WINDOW_TITLE)
    EndIf
    If $hWnd = 0 Then Return SetError(2, 0, False)
;~ 	MsgBox(0, "Sender", "hWnd = " & $hWnd & @CRLF & "g_WM_RUNBTN = " & $g_WM_RUNBTN & @CRLF & "id = " & $id)
    Local $ret = DllCall("user32.dll", "bool", "PostMessageA", _
                         "hwnd", $hWnd, "uint", $g_WM_RUNBTN, "wparam", $id, "lparam", 0)
    Return (Not @error) And IsArray($ret) And ($ret[0] <> 0)
EndFunc

; ----------------- Launcher + Sandboxie + Firewall -----------------
Func SetFirewallNetAccess($exePath, $enable)
    If $exePath = "" Then Return
    Local $ruleName = StringRegExpReplace($exePath, "^.*\\", "MenuApp_")
    If $enable Then
        RunWait('netsh advfirewall firewall delete rule name="' & $ruleName & '"', "", @SW_HIDE)
    Else
        RunWait('netsh advfirewall firewall add rule name="' & $ruleName & '" dir=out action=block program="' & $exePath & '" enable=yes', "", @SW_HIDE)
        RunWait('netsh advfirewall firewall add rule name="' & $ruleName & '" dir=in action=block program="' & $exePath & '" enable=yes', "", @SW_HIDE)
    EndIf
EndFunc

Func _IsServiceRunning($svcName)
    Local $pid = Run(@ComSpec & " /c sc query " & $svcName, "", @SW_HIDE, $STDOUT_CHILD)
    Local $out = ""
    While 1
        $out &= StdoutRead($pid)
        If @error Then ExitLoop
    WEnd
    Return StringInStr($out, "RUNNING") > 0
EndFunc

Func _EnsureSandboxieService()
    If Not _IsServiceRunning($SandboxieServiceName) Then
        RunWait(@ComSpec & " /c net start " & $SandboxieServiceName, "", @SW_HIDE)
    EndIf
EndFunc

Func _StopSandboxieService()
    If _IsServiceRunning($SandboxieServiceName) Then
        Local $answer = MsgBox(65, "Sandboxie", "Stopping Sandboxie services..." & @CRLF & @CRLF & "Press OK to proceed or Cancel to abort.")
        If $answer = 1 Then ; OK pressed
            RunWait(@ComSpec & " /c net stop " & $SandboxieServiceName, "", @SW_HIDE)
            ShellExecute(@ComSpec, ' /c net stop ' & $SandboxieServiceName, '', 'runas')
        Else
            ; Cancel pressed: do nothing
            Return
        EndIf
    EndIf
    If ProcessExists("SandMan.exe") Then ProcessClose("SandMan.exe")
EndFunc

Func LaunchApp($index)
    Local $appName      = $appsData[$index][$APP_NAME]
    Local $appPath      = $appsData[$index][$APP_PATH]
    Local $arguments    = $appsData[$index][$APP_ARGS]
    Local $singleInstance = $appsData[$index][$APP_SINGLE]
    Local $runAsAdmin   = $appsData[$index][$APP_ADMIN]
    Local $netAccess    = $appsData[$index][$APP_NET]
    Local $sandboxiePath = $appsData[$index][$APP_SBX_PATH]
    Local $sandboxName  = $appsData[$index][$APP_SBX_NAME]
    Local $fileAssocApp = $appsData[$index][$APP_ASSOC]

    ; Resolve relative paths to absolute
    If Not StringRegExp($sandboxiePath, "^[A-Za-z]:\\") And $sandboxiePath <> "" Then $sandboxiePath = @ScriptDir & "\" & $sandboxiePath
    If Not StringRegExp($appPath, "^[A-Za-z]:\\") And $appPath <> "" Then $appPath = @ScriptDir & "\" & $appPath

    If $appPath = "" Or Not FileExists($appPath) Then
        MsgBox(48, $appName, "No app found at this location:" & @CRLF & $appPath)
        Return
    EndIf

    ; Compute the correct working directory for the target app
    Local $workdir = _DirOf($appPath)

    ; Enforce single instance if requested
    If $singleInstance Then
        Local $exeName = StringRegExpReplace($appPath, "^.*\\", "")
        If ProcessExists($exeName) Then
            MsgBox(48, $appName, "This app is already running:" & @CRLF & $exeName)
            Return
        EndIf
    EndIf

    ; Apply firewall rules as needed
    SetFirewallNetAccess($appPath, $netAccess)

    ; Launch via file association wrapper if configured
    If $fileAssocApp <> "" Then
        If $runAsAdmin Then
            ShellExecute($fileAssocApp, '"' & $appPath & '"', $workdir, "runas")
        Else
            ShellExecute($fileAssocApp, '"' & $appPath & '"', $workdir)
        EndIf
        Return
    EndIf

    ; Launch via Sandboxie if configured
    If $sandboxiePath <> "" Then
        If $sandboxiePath = "%SandboxiePath%" Then $sandboxiePath = $SandboxiePath
        If $sandboxName = "" Then $sandboxName = "DefaultBox"
        Local $cmd = '"' & $sandboxiePath & '" /box:' & $sandboxName & ' "' & $appPath & '"'
        If StringStripWS($arguments, 3) <> "" Then $cmd &= " " & $arguments
        _EnsureSandboxieService()
        ; IMPORTANT: pass the target app's folder as the working dir
        Run($cmd, $workdir, @SW_SHOW)
        Return
    EndIf

    ; Direct launch (with correct working directory)
    If $runAsAdmin Then
        ShellExecute($appPath, $arguments, $workdir, "runas")
    Else
        ShellExecute($appPath, $arguments, $workdir)
    EndIf
EndFunc

Func _IsMainButton($i)
    Local $cat = StringLower(StringStripWS($appsData[$i][$APP_CAT], 3))
    If $cat = "" Then $cat = "uncategorised"
    If $cat = "portableapps" Then Return False
    If $cat = "fave" Then Return False
    Return True
EndFunc

Func RunButtonById($id)
    For $i = 0 To UBound($appsData) - 1
        If $appsData[$i][$APP_ID] = $id Then
            LaunchApp($i)
            Local $closeSetting = IniRead(@ScriptDir & "\App\Settings.ini", "GLOBAL", "CloseMenuOnClick", "0")
            If $closeSetting == "1" Then Exit
            Return True
        EndIf
    Next
    MsgBox(48, "AxiomTrayLauncher", "No button with ID: BUTTON" & $id)
    Return False
EndFunc

Func RestartAsAdmin()
    ShellExecute(@ScriptFullPath, "/restart-admin", "", "runas")
    Exit
EndFunc
