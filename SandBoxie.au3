; SandBoxie.au3

#include-once

Func _RunWithSandboxie($exe, $args, $workDir, $sandboxName, $settingsIni)
    Local $sandboxieRaw = IniRead($settingsIni, "GLOBAL", "SandboxiePath", "")
    If $sandboxieRaw = "" Then
        MsgBox(16, "Sandboxie Error", "SandboxiePath not set in Settings.ini")
        Return SetError(1, 0, 0)
    EndIf
    Local $sandboxiePath = _ResolvePath($sandboxieRaw, @ScriptDir)
    If Not FileExists($sandboxiePath) Then
        MsgBox(16, "Sandboxie Error", "Sandboxie executable not found: " & $sandboxiePath)
        Return SetError(2, 0, 0)
    EndIf

    Local $cmd = '"' & $sandboxiePath & '" /box:' & $sandboxName & ' "' & $exe & '"'
    If $args <> "" Then $cmd &= " " & $args

    Run($cmd, $workDir)

    ; Wait for the app to launch (up to 10 seconds)
    Local $appExeName = StringTrimLeft($exe, StringInStr($exe, "\", 0, -1))
    Local $sandboxedPid = 0
    Local $start = TimerInit()
    While TimerDiff($start) < 10000
        $sandboxedPid = ProcessExists($appExeName)
        If $sandboxedPid <> 0 Then ExitLoop
        Sleep(250)
    WEnd

    If $sandboxedPid = 0 Then
        ; Could not find app, but let Sandboxie run anyway (no error, just info)
        MsgBox(48, "Sandboxie Warning", "Could not find sandboxed app process: " & $appExeName & @CRLF & "Sandboxie may still be running the app.")
    Else
        ; Wait for it to exit
        While ProcessExists($sandboxedPid)
            Sleep(500)
        WEnd
    EndIf

    ; Close SandMan.exe if running
    If ProcessExists("SandMan.exe") Then
        ProcessClose("SandMan.exe")
    EndIf

    ; Stop Sandboxie service, request elevation only if needed
    _StopSandboxieService()

    Return 1
EndFunc

Func _StopSandboxieService()
    Local $SandboxieServiceName = "SbieSvc"
    If _IsServiceRunning($SandboxieServiceName) Then
        RunWait(@ComSpec & " /c net stop " & $SandboxieServiceName, "", @SW_HIDE)
        Sleep(1000)
        If _IsServiceRunning($SandboxieServiceName) Then
            ShellExecute(@ComSpec, " /c net stop " & $SandboxieServiceName, "", "runas")
            Sleep(2000)
        EndIf
    EndIf

    ; Final check: forcibly kill SbieSvc.exe if still running
    If ProcessExists("SbieSvc.exe") Then
        ; Try normal kill
        ProcessClose("SbieSvc.exe")
        Sleep(500)
        ; If STILL running, force kill with taskkill
        If ProcessExists("SbieSvc.exe") Then
            ShellExecute(@ComSpec, ' /c taskkill /F /IM SbieSvc.exe', '', 'runas')
            Sleep(1000)
        EndIf
    EndIf

    ; Debugging messages
    If ProcessExists("SbieSvc.exe") Then
        MsgBox(16, "Sandboxie Service", "SbieSvc.exe is STILL running even after forced kill!")
    Else
        MsgBox(64, "Sandboxie Service", "SbieSvc.exe process is now closed.")
    EndIf
EndFunc

Func _IsServiceRunning($serviceName)
    Local $output = ""
    Local $cmd = 'sc query "' & $serviceName & '"'
    Local $pid = Run(@ComSpec & " /c " & $cmd, "", @SW_HIDE, $STDOUT_CHILD)
    While ProcessExists($pid)
        $output &= StdoutRead($pid)
        Sleep(50)
    WEnd
    Return StringInStr($output, "RUNNING") > 0
EndFunc
