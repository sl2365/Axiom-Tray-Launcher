# Axiom-Tray-Launcher
Tray app for launching apps.

Full relative path support throughout the app.
Supports launching apps via Sandboxie-Plus, just add the path to SandMan.exe and enable Sandboxie in the settings.
Folder scanning for instant listing of all apps within a folder, includes recursive scanning. You can also set a folder to scan for certain file types, default is EXE.
Organise apps by specifying categories. The default is the scanned folder name for your categories.
Set apps as "Favourite" so it shows below categories.
Hide unwanted apps.
Generate shortcuts for all apps. Doesn't create shortcuts for hidden apps.

Set Global variables for paths. Can be used everywhere.
Generate symlinks globally or per app. Global symlink can be set to be created/removed on startup/exit of the tray menu. Or you can set them just on startup so they can be persistent.
Per app symlinks are always created/removed on starting/exiting the app. So you can choose to use symlinks here or make them global and persistent.
