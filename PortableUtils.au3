; --- PortableUtils.au3 ---
; Advanced portable app launching utility functions for INI settings

Func SetPortableEnvVars($settings)
    For $key In $settings.Keys()
        If StringLeft($key, 4) = "Env_" Then
            Local $envPath = $settings.Item($key)
            If StringLeft($envPath, 2) = ".\" Then
                DirCreate($envPath)
            EndIf
            EnvSet(StringTrimLeft($key, 4), $envPath)
        EndIf
    Next
EndFunc

Func SetupSymLinks($settings)
    For $key In $settings.Keys()
        If $key = "SymLink" Or StringRegExp($key, "^SymLink\d+$") Then
            Local $pair = StringSplit($settings.Item($key), "|")
            If $pair[0] = 2 Then
                Local $source = $pair[1]
                Local $dest = $pair[2]
                If FileExists($dest) Then
                    ; Check if $dest is a directory
                    If StringInStr(FileGetAttrib($dest), "D") Then
                        RunWait(@ComSpec & ' /c rmdir "' & $dest & '"', "", @SW_HIDE)
                    Else
                        FileDelete($dest)
                    EndIf
                EndIf
                ; Create parent directory for $dest
                DirCreate(StringLeft($dest, StringInStr($dest, "\", 0, -1) - 1))
                ; Check if $source is a directory
                If FileExists($source) And StringInStr(FileGetAttrib($source), "D") Then
                    RunWait(@ComSpec & ' /c mklink /J "' & $dest & '" "' & $source & '"', "", @SW_HIDE)
                ; Check if $source is a file
                ElseIf FileExists($source) Then
                    RunWait(@ComSpec & ' /c mklink /H "' & $dest & '" "' & $source & '"', "", @SW_HIDE)
                EndIf
            EndIf
        EndIf
    Next
EndFunc

Func RemoveSymLinks($settings)
    For $key In $settings.Keys()
        If $key = "SymLink" Or StringRegExp($key, "^SymLink\d+$") Then
            Local $pair = StringSplit($settings.Item($key), "|")
            If $pair[0] = 2 Then
                Local $dest = $pair[2]
                If FileExists($dest) Then
                    ; Check if $dest is a directory
                    If StringInStr(FileGetAttrib($dest), "D") Then
                        RunWait(@ComSpec & ' /c rmdir "' & $dest & '"', "", @SW_HIDE)
                    Else
                        FileDelete($dest)
                    EndIf
                EndIf
            EndIf
        EndIf
    Next
EndFunc

Func HandleRegistry($settings, $mode)
    If $settings.Exists("Reg_Key") Then
        Local $regKey = $settings.Item("Reg_Key")
        Local $backupFile = @ScriptDir & "\Data\RegBackup_" & StringReplace($regKey, "\", "_") & ".reg"
        If $mode = "restore" And FileExists($backupFile) Then
            RunWait(@ComSpec & ' /c reg import "' & $backupFile & '"', "", @SW_HIDE)
        ElseIf $mode = "backup" Then
            RunWait(@ComSpec & ' /c reg export "' & $regKey & '" "' & $backupFile & '" /y', "", @SW_HIDE)
        EndIf
    EndIf
EndFunc

Func DoCleanup($settings)
    For $key In $settings.Keys()
        If StringLeft($key, 7) = "Cleanup" Then
            Local $pattern = $settings.Item($key)
            If StringRegExp($pattern, "\*") Then
                FileDelete($pattern)
            ElseIf FileExists($pattern) And StringInStr(FileGetAttrib($pattern), "D") Then
                DirRemove($pattern, 1)
            EndIf
        EndIf
    Next
EndFunc

Func ShowSplash($settings)
    If $settings.Exists("Splash") Then
        Local $img = $settings.Item("Splash")
        Local $wait = 1500
        If $settings.Exists("SplashWait") Then $wait = Number($settings.Item("SplashWait"))
        SplashImageOn("Launching...", $img, 400, 200)
        Sleep($wait)
        SplashOff()
    EndIf
EndFunc

Func LaunchSandboxed($exe, $args, $workdir, $settings)
    If $settings.Exists("Sandboxie") Then
        Local $sandboxiePath = $settings.Item("Sandboxie")
        Local $cmd = '"' & $sandboxiePath & '" /box:DefaultBox "' & $exe & '" ' & $args
        Run($cmd, $workdir)
        Return True
    EndIf
    Return False
EndFunc

Func MapDriveLetter($settings)
    If $settings.Exists("DriveMap") Then
        Local $parts = StringSplit($settings.Item("DriveMap"), " ")
        If $parts[0] = 2 Then
            Local $driveLetter = $parts[1]
            Local $targetPath = $parts[2]
            RunWait(@ComSpec & ' /c subst ' & $driveLetter & ' "' & $targetPath & '"', "", @SW_HIDE)
        EndIf
    EndIf
EndFunc

Func IsSingleInstance($settings)
    If $settings.Exists("SingleInstance") And StringLower($settings.Item("SingleInstance")) = "true" Then
        Local $exe = $settings.Item("Exe")
        If ProcessExists($exe) Then Return True
    EndIf
    Return False
EndFunc

Func LaunchPortableApp($settings)
    If IsSingleInstance($settings) Then
        MsgBox(48, "Already Running", "This app is already running.")
        Return
    EndIf

    ShowSplash($settings)
    MapDriveLetter($settings)
    SetPortableEnvVars($settings)
    SetupSymLinks($settings)
    HandleRegistry($settings, "restore")

    If Not $settings.Exists("Exe") Then
        MsgBox(16, "Error", "Missing 'Exe' setting in button section.")
        RemoveSymLinks($settings)
        Return
    EndIf
    Local $exe = $settings.Item("Exe")
    Local $workdir = @ScriptDir
    If $settings.Exists("WorkDir") Then $workdir = $settings.Item("WorkDir")
    Local $args = ""
    If $settings.Exists("Arguments") Then $args = $settings.Item("Arguments")

    If LaunchSandboxed($exe, $args, $workdir, $settings) Then
        HandleRegistry($settings, "backup")
        RemoveSymLinks($settings)
        DoCleanup($settings)
        Return
    EndIf

    Local $cmd = '"' & $exe & '"'
    If $args <> "" Then $cmd &= " " & $args
    Local $pid = Run($cmd, $workdir)
    If $pid Then
        While ProcessExists($pid)
            Sleep(500)
        WEnd
    EndIf

    HandleRegistry($settings, "backup")
    RemoveSymLinks($settings)
    DoCleanup($settings)
EndFunc
