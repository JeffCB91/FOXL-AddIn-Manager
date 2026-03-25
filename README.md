# FOXL Add-In Manager

A lightweight, standalone desktop utility built in Python to help developers and users manage their Front Office Excel Add-In environments, configuration paths, and local test builds.

## Features

* **Dual Interfaces:** Separate build targets for Developers (full config viewing, network scanning) and Users (streamlined environment toggling and zip installation).
* **Environment Management:** Easily toggle your local `env` file between `Prod`, `UAT`, `Dev`, and `local`.
* **Zip Build Installer:** Users can browse for a local test `.zip` file, provide a folder name, and the tool will automatically extract the contents to `C:\ExcelAddIn` and repoint the Excel registry to the new build.
* **Registry Management:** Safely swaps Excel's `OPEN` registry keys to point to test builds, with a one-click "Revert to Standard Loader" safety net.
* **Excel Process Controls:** Built-in buttons to launch Excel, request a safe close, or forcefully kill hung Excel processes natively.
* **Version Tracking:** Asynchronously checks network drives for available .Net 8.0 and .Net 6.0 add-in versions.
* **Config Viewer (Admin):** Dynamically parses and displays the JSON configuration variables for any local or network version.

## Project Structure

The codebase is modularized to separate the user interfaces from the business logic:

```text
foxl-manager/
├── config.py                # Stores all hardcoded paths, URLs, and constants
├── main.py                  # Entry point for the Developer/Admin dashboard
├── main_user.py             # Entry point for the streamlined User dashboard
│
├── core/                    # Core business logic
│   ├── env_manager.py       # Reads and writes to the local .env file
│   ├── scanner.py           # Asynchronously scans network and local directories
│   ├── registry_ops.py      # Handles Windows Registry scanning and updating
│   ├── system_ops.py        # Helper functions (Explorer, Excel launchers/killers)
│   └── zip_ops.py           # Handles zip extraction and file moving
│
└── ui/                      # Presentation layer (Tkinter)
    ├── main_window.py       # The Developer dashboard UI
    ├── user_window.py       # The User configurator UI
    └── config_viewer.py     # The dynamic JSON configuration viewer
```

## Prerequisites

To run the script from the source code or build a new executable, you will need:
* **Python 3.x** installed on your Windows machine.
* **Windows OS** (This script utilizes `winreg`, `subprocess`, and `os.startfile`).

## Installation & Setup

1. Clone or download this repository to your local machine.
2. Open a terminal in the project directory.
3. Install the build requirements:
   ```bash
   pip install -r requirements.txt
   ```

## Building the Executables (.exe)

This project supports building two separate executables depending on the target audience.

**1. Building the Admin/Developer Version:**
Run this command to build the full tool with network scanning and config viewing:
```bash
pyinstaller --onefile --windowed --name "foxl_addin_admin" main.py
```

**2. Building the Standard User Version:**
Run this command to build the restricted version focused purely on environment switching and local zip installations:
```bash
pyinstaller --onefile --windowed --name "foxl_addin_user" main_user.py
```

Once the build finishes, your `.exe` files will be located in the `dist/` folder.

### Cleaning Up Build Files
PyInstaller generates temporary folders during the build process. To clean up your workspace after grabbing your `.exe`:

**Command Prompt (cmd):**
```cmd
rmdir /s /q build dist
del /q *.spec
```

**PowerShell:**
```powershell
Remove-Item -Recurse -Force build, dist, *.spec -ErrorAction SilentlyContinue
```

## Command Line Usage (Quick Environment Toggling)

If you prefer to change your environment from the terminal without opening the UI, you can use these one-liners (replace `Dev` with `Prod`, `UAT`, or `local`):

**PowerShell:**
```powershell
$f = "$env:APPDATA\NinetyOne - FrontOfficeExcelAddIn\front-office-excel-addin-env"; (Get-Content $f) -replace '^ENV=.*', 'ENV=Dev' | Set-Content $f
```

**Command Prompt (cmd):**
```cmd
set "f=%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\front-office-excel-addin-env" & type "%f%" | findstr /v "^ENV=" > "%f%.tmp" & echo ENV=Dev>> "%f%.tmp" & move /y "%f%.tmp" "%f%"
```