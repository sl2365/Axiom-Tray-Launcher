; SandBoxie.au3:

#include-once
#include "TrayMenu.au3"

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

    ; Use Start.exe for Sandboxie-Plus, not SandMan.exe
    ; If settings only gives SandMan.exe, fix your config to point to Start.exe
    Local $cmd = '"' & $sandboxiePath & '" /box:' & $sandboxName & ' "' & $exe & '"'
    If $args <> "" Then $cmd &= " " & $args

    Local $pid = Run($cmd, $workDir)
    If $pid = 0 Then
        MsgBox(16, "Sandboxie Error", "Failed to launch Sandboxie!")
        Return SetError(3, 0, 0)
    EndIf

    ; Attempt to get the real process name (works for .exe only)
    Local $appExeName = $exe
    If StringInStr($appExeName, "\") Then
        $appExeName = StringTrimLeft($appExeName, StringInStr($appExeName, "\", 0, -1))
    EndIf

    ; Try to get the PID of the sandboxed app (optional, for monitoring)
    Local $sandboxedPid = 0
    Local $start = TimerInit()
    While TimerDiff($start) < 5000
        $sandboxedPid = ProcessExists($appExeName)
        If $sandboxedPid <> 0 Then ExitLoop
        Sleep(250)
    WEnd

    ; Return the launched PID (Sandboxie launcher), and sandboxed app PID if found
    ; You should add these to $g_MonitoredApps for monitoring/cleanup in your main loop!
    Local $result[2]
	$result[0] = $pid
	$result[1] = $sandboxedPid
	Return $result
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
        ProcessClose("SbieSvc.exe")
        Sleep(500)
        If ProcessExists("SbieSvc.exe") Then
            ShellExecute(@ComSpec, ' /c taskkill /F /IM SbieSvc.exe', '', 'runas')
            Sleep(1000)
        EndIf
    EndIf

    If ProcessExists("SbieSvc.exe") Then
        MsgBox(16, "Sandboxie Service", "SbieSvc.exe is STILL running even after forced kill!")
    Else
;~         MsgBox(64, "Sandboxie Service", "SbieSvc.exe process is now closed.")
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
