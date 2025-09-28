# Axiom-Tray-Launcher
Axiom Tray Launcher is a powerful and flexible Windows utility that allows you to launch applications, manage shortcuts, create and clean up symbolic links, and organize your workflow from a single tray menu. Designed for power users and admins, it streamlines the launching and management of apps, folders, and system tools, all while offering advanced features for automation and environment customization.

---

## Uses:
- Combine all portable menu apps into a single menu.
- Launch specific apps via Sandboxie.
- Launch specific apps with Junctions/SymLinks.
- Create a complete set up of all your Junctions/SymLinks. Add and remove with a single click.
- Fully portable.

---

## Features:
- **Customizable Tray Menu -** Organize and launch your applications, files, and scripts directly from the system tray.
- **Portable -** Full relative path support throughout the app.
- **Sandboxie support -** Supports launching apps via Sandboxie-Plus, just add the path to SandMan.exe, specify the sandbox name and enable Sandboxie in the settings.
- **Specify scan locations -** Folder scanning for instant listing of all apps/files within a folder, includes recursive scanning. You can also set a filter to scan for certain file types, default is EXE.
- **Blacklist -** Set up a list to exclude apps/files during scanning.
- **Categories -** Organise apps into categories. The default category is the scanned folder name. You can change this on a per-app basis.
- **Favourites -** For easier access, you can set apps as "Favourite" so it shows below categories.
- **Hide -** Hide unwanted apps. You may not want to ignore but temporarily hide some apps.
- **Shortcut Generation -** Automatically generate shortcuts for apps and files, with easy configuration. Doesn't create shortcuts for hidden apps.
- **Environment Variable Management -** View, copy, and use Windows environment variables in paths and shortcuts, or create your own.
- **Symlink Management -** Create, monitor, and clean up symbolic links and junctions for global or per-app use.
- **Automatic Cleanup -** Cleans up symlinks and resources on exit or when monitored apps close.
- **Update Checker -** Built-in updater checks for new releases.
- **Settings GUI -** Intuitive settings window for managing categories, environment variables, ignored items, and more.
- **Per-App Configuration -** Customize symlinks and settings for individual apps.

---

## Getting Started

### Installation
1. Download the latest release from [GitHub Releases](https://github.com/sl2365/Axiom-Tray-Launcher/releases).
2. Extract the files to a folder of your choice.
3. Run `AxiomTrayLauncher.exe`. The launcher will appear as an icon in your Windows system tray.
4. Click the icon and hover over the title to open a sub-menu, where you'll find the Settings.

---

## Usage Guide

### Tray Menu

- **Left-click the tray icon** to open the menu.
- Launch any configured application, file, or shortcut with a single click.
- Use the "Settings" option to customize categories, add or remove apps, and configure symlinks.

### Settings GUI

- Open the Settings from the tray menu.
- Use the tabs to:
  - **Global -** Set user variables and main options.
  - **Scan Folders -** Add or scan folders for quick launching. (100 Max)
  - **Apps -** Manage your list of applications and their configurations.
  - **Ignore -** Exclude specific items from menus or scans. (300 Max)
  - **Env Vars -** View and copy system environment variables.
  - **About -** Check for updates and access documentation.

### Symlink and Variables Management

- **User Variables:**
  - Create your own variables that can be used throughout the app.

- **Global Symlinks:**  
  - In Settings, configure global symlinks under the appropriate tab.
  - Enable or disable global symlink creation/removal. For example, enable on startup, disable removal on exit for persistent links.
  - The launcher will create or clean up these links automatically as needed.

- **Per-App Symlinks:**  
  - In the Apps tab, configure symlinks for specific applications.
  - When launching or closing an app, Axiom Tray Launcher will create/remove symlinks as required.

### Shortcut Generation

- Apps and files added to the launcher can have shortcuts generated for them.
- Configure shortcut settings in the Apps or Scan Folders tab.
- Shortcuts are created in the same folder structure as their physical locations, but within the Axiom folder: Shortcuts.

### Sandboxie & App Monitoring

- Launch apps in Sandboxie directly from the tray.
- The launcher monitors Sandboxie and symlinked apps asynchronously to avoid freezing the menu.
- Resources and links are cleaned up after monitored apps close.

### Environment Variables

- Access the **Env Vars** tab in Settings to view, copy, and use Windows environment variables.
- Variables can be used in symlink and shortcut paths for flexible configuration.

---

## Tips

- **Symlink Cleanup:**  
  The launcher will prompt or automatically clean up symlinks on exit if enabled.
- **Manual Actions:**  
  You can manually trigger symlink creation or removal from the tray menu or Settings.
- **Advanced Configuration:**  
  Edit `.ini` files in the `App` directory for fine-tuned control over categories, apps, and symlinks. This is the same as the Settings GUI, so either way works.

---

## License

Axiom Tray Launcher is released under the GPL-3.0 License.
