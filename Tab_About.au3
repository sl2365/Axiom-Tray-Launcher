; Tab_About.au3

#include-once

Global $g_AboutLinkCtrl, $g_CheckUpdatesCtrl, $g_OpenSettingsFolderCtrl

; Tab_About_Create updated to create all three buttons and return their handles
Func Tab_About_Create($hGUI, $hTab)
    GUICtrlCreateTabItem("About")
    Local $startY = 50
    GUICtrlCreatePic(@ScriptDir & "\App\AxiomTrayJPG.jpg", 20, $startY + 25, 64, 64)
    Local $lblTitle = GUICtrlCreateLabel("Axiom Tray Launcher", 110, $startY, 250, 30)
    GUICtrlSetFont($lblTitle, 16, 800, 0, "Segoe UI")

    Local $aAboutLabels[5]
    $aAboutLabels[0] = GUICtrlCreateLabel("Version: " & $UPDATE_VERSION, 110, $startY + 35, 250, 20)
    $aAboutLabels[1] = GUICtrlCreateLabel("Author: sl23", 110, $startY + 60, 250, 20)
    $aAboutLabels[2] = GUICtrlCreateLabel("Build date: 2025-08-17", 110, $startY + 85, 250, 20)
    $aAboutLabels[3] = GUICtrlCreateLabel("A tray menu to launch your portable apps.", 110, $startY + 110, 250, 20)
    $aAboutLabels[4] = GUICtrlCreateLabel("Website: https://github.com/sl2365/Axiom-Tray-Launcher", 110, $startY + 135, 550, 20)

    For $i = 0 To UBound($aAboutLabels) - 1
        GUICtrlSetFont($aAboutLabels[$i], 10)
    Next

    ; Visit Website button
    $g_AboutLinkCtrl = GUICtrlCreateButton("Visit Website", 150, $startY + 190, 180, 28)
    GUICtrlSetFont($g_AboutLinkCtrl, 10, 700)

    ; Check for Updates button
    $g_CheckUpdatesCtrl = GUICtrlCreateButton("Check for Updates", 150, $startY + 225, 180, 28)
    GUICtrlSetFont($g_CheckUpdatesCtrl, 10, 700)

    ; Open Settings Folder button
    $g_OpenSettingsFolderCtrl = GUICtrlCreateButton("Open Settings Folder", 150, $startY + 260, 180, 28)
    GUICtrlSetFont($g_OpenSettingsFolderCtrl, 10, 700)

    Return $lblTitle
EndFunc
