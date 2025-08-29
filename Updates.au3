#include-once

; --------- Global Variables ---------
Global $UPDATE_VERSION = _Updates_GetFileVersion() ; Current app version, auto-read from EXE if possible
Global $UPDATE_URL = "https://raw.githubusercontent.com/sl2365/Axiom-Tray-Launcher/main/latest.txt" ; Raw link to latest.txt
Global $UPDATE_ZIP_URL_PREFIX = "https://github.com/sl2365/Axiom-Tray-Launcher/releases/download/"
Global $UPDATE_ZIP_FILENAME = "AxiomTrayLauncher.zip"

Global $SETTINGS_INI = @ScriptDir & "\App\Settings.ini"
Global $UPDATE_SECTION = "GLOBAL"
Global $UPDATE_ON_START_KEY = "UpdateOnStart"
Global $UPDATE_LAST_CHECK_KEY = "LastUpdateCheck"

; --------- Check for Updates ---------
Func Updates_Check($manualCheck = False)
    Local $checkOnStart = IniRead($SETTINGS_INI, $UPDATE_SECTION, $UPDATE_ON_START_KEY, "0")
    Local $lastCheck = IniRead($SETTINGS_INI, $UPDATE_SECTION, $UPDATE_LAST_CHECK_KEY, "")
    Local $today = @YEAR & @MON & @MDAY

    If Not $manualCheck And $checkOnStart <> "1" Then Return
    If Not $manualCheck And $lastCheck = $today Then Return

    ; Download latest.txt and get latest version (only version number, e.g. 1.0.0.215)
    Local $latestInfo = InetRead($UPDATE_URL, 1)
    If @error Then
        If $manualCheck Then MsgBox(16, "Update Check", "Unable to check for updates (network error).")
        Return
    EndIf

    Local $latestVersion = StringStripWS(BinaryToString($latestInfo), 3)
    If StringLen($latestVersion) < 5 Then
        If $manualCheck Then MsgBox(16, "Update Check", "Update info format error (latest.txt should contain only the version number).")
        Return
    EndIf

    Local $downloadUrl = $UPDATE_ZIP_URL_PREFIX & $latestVersion & "/" & $UPDATE_ZIP_FILENAME

    If _Updates_VersionCompare($latestVersion, $UPDATE_VERSION) <= 0 Then
        If $manualCheck Then MsgBox(64, "Update Check", "No new updates available." & @CRLF & "Current version: " & $UPDATE_VERSION)
        IniWrite($SETTINGS_INI, $UPDATE_SECTION, $UPDATE_LAST_CHECK_KEY, $today)
        Return
    EndIf

    ; ----------- Confirmation message before proceeding with update ----------
    Local $warnMsg = "A new update (" & $latestVersion & ") is available." & @CRLF & _
        "Current version: " & $UPDATE_VERSION & @CRLF & @CRLF & _
        "Updating will close AxiomTrayLauncher" & @CRLF & "and replace it with the new version." & @CRLF & @CRLF & _
        "Continue with update?" ; <- Clear warning
    Local $response = MsgBox(33, "Update Confirmation", $warnMsg)
    If $response <> 1 Then
        ; User cancelled, do not proceed
        IniWrite($SETTINGS_INI, $UPDATE_SECTION, $UPDATE_LAST_CHECK_KEY, $today)
        Return
    EndIf

    ; ------------------- Proceed with update as before ----------------------
    Local $zipPath = @TempDir & "\" & $UPDATE_ZIP_FILENAME
    Local $psCmd = 'powershell -Command "Invoke-WebRequest -Uri ''' & $downloadUrl & ''' -OutFile ''' & $zipPath & '''"'
    RunWait(@ComSpec & " /c " & $psCmd, "", @SW_HIDE)
    If @error Or Not FileExists($zipPath) Then
        MsgBox(16, "Update Error", "Download failed. Please try again later.")
        Return
    EndIf

    ; Unzip using PowerShell (no 7z.exe needed)
    Local $extractDir = @ScriptDir & "\App\_updater"
    DirCreate($extractDir)
    _Updates_UnzipWithPowerShell($zipPath, $extractDir)

    ; DO NOT copy files or delete _updater folder here!
    LaunchUpdateHelper($extractDir)

    MsgBox(64, "Update Complete", "Click OK to restart and apply the update.")
    IniWrite($SETTINGS_INI, $UPDATE_SECTION, $UPDATE_LAST_CHECK_KEY, $today)

    ; Exit so the batch script can replace the EXE
    Exit
EndFunc

Func LaunchUpdateHelper($extractDir)
    Local $batFile = @TempDir & "\update_launcher.bat"
    Local $srcExe = $extractDir & "\AxiomTrayLauncher.exe"
    Local $destExe = @ScriptDir & "\AxiomTrayLauncher.exe"

    Local $batchContent = ""
    $batchContent &= "@echo off" & @CRLF
    $batchContent &= ":loop" & @CRLF
    $batchContent &= 'tasklist | find /i "AxiomTrayLauncher.exe" >nul' & @CRLF
    $batchContent &= "if not errorlevel 1 (" & @CRLF
    $batchContent &= "    timeout /t 1 >nul" & @CRLF
    $batchContent &= "    goto loop" & @CRLF
    $batchContent &= ")" & @CRLF
    $batchContent &= 'copy /y "' & $srcExe & '" "' & $destExe & '"' & @CRLF
    $batchContent &= 'start "" /b "' & $destExe & '"' & @CRLF ; <-- /b hides window!
    $batchContent &= 'del "%~f0"' & @CRLF
	$batchContent &= 'exit' & @CRLF

    FileWrite($batFile, $batchContent)
    ShellExecute($batFile, "", @TempDir, "", @SW_HIDE)
EndFunc

Func _Updates_UnzipWithPowerShell($zipPath, $extractDir)
    Local $psCmd = 'powershell -Command "Expand-Archive -Force -Path ''' & $zipPath & ''' -DestinationPath ''' & $extractDir & '''"'
    RunWait(@ComSpec & " /c " & $psCmd, "", @SW_HIDE)
EndFunc

Func _Updates_VersionCompare($v1, $v2)
    Local $a1 = StringSplit($v1, ".", 2)
    Local $a2 = StringSplit($v2, ".", 2)
    Local $maxLen = (UBound($a1) > UBound($a2)) ? UBound($a1) : UBound($a2)
    For $i = 0 To $maxLen - 1
        Local $n1 = ($i < UBound($a1)) ? Number($a1[$i]) : 0
        Local $n2 = ($i < UBound($a2)) ? Number($a2[$i]) : 0
        If $n1 < $n2 Then Return -1
        If $n1 > $n2 Then Return 1
    Next
    Return 0
EndFunc

Func _Updates_CopyFiles($srcDir, $destDir)
    Local $search = FileFindFirstFile($srcDir & "\*")
    If $search = -1 Then Return
    While 1
        Local $file = FileFindNextFile($search)
        If @error Then ExitLoop
        If $file = "." Or $file = ".." Then ContinueLoop
        Local $srcPath = $srcDir & "\" & $file
        Local $destPath = $destDir & "\" & $file
        If StringInStr(FileGetAttrib($srcPath), "D") Then
            DirCreate($destPath)
            _Updates_CopyFiles($srcPath, $destPath)
        Else
            FileCopy($srcPath, $destPath, 1)
        EndIf
    WEnd
    FileClose($search)
EndFunc

; --------- Optional: Auto-Read File Version from EXE ---------
Func _Updates_GetFileVersion()
    Local $ver = FileGetVersion(@ScriptFullPath, $FV_FILEVERSION)
    If $ver = "" Or @error Then
        Return "0.0.0.0"
    Else
        Return $ver
    EndIf
EndFunc

