# FOXL Add-In Manager

A lightweight, standalone desktop utility built in Python to help developers and users manage their Front Office Excel Add-In environments and configuration paths.

## Features

* **Environment Management:** Easily toggle your local `.env` file between `Prod`, `UAT`, `Dev`, and `local` without manually editing text files.
* **Version Tracking:** Asynchronously checks the network drives for available `.Net 8.0` and `.Net 6.0` add-in versions.
* **Registry Diagnostics:** Scans the Windows Registry (`HKCU\Software\Microsoft\Office\16.0\excel\options`) for any mentions of 'NinetyOne' to ensure clean installations or pinpoint issues.
* **Quick Access Directories:** One-click buttons to open Windows Explorer directly to your local `.env` folder, loader paths, and local test paths.
* **Registry Editor Shortcut:** Instantly opens `regedit.exe` navigated exactly to the Excel options key.

## Prerequisites

To run the script from the source code or build a new executable, you will need:
* **Python 3.x** installed on your Windows machine.
* **Windows OS** (This script utilizes `winreg` and `os.startfile`, which are Windows-specific).

## Installation & Setup

1. Clone or download this repository to your local machine.
2. Open a terminal in the project directory.
3. Install the build requirements (only if you plan to build to executable):
```bash
  pip install -r requirements.txt
```
## Running from Source
To run the application directly via Python:
```bash
  python addin_manager.py
```
## Building the Executable (.exe)
To share this tool with team members who do not have Python installed, you can compile it into a single standalone `.exe` file.

Run the following command in your terminal:

```bash
    pyinstaller --onefile --windowed foxl_addin_manager.py
```
- `--onefile`: Packages everything into a single `.exe`.
- `--windowed`: Prevents the black command prompt console from appearing behind the UI.

Once the build finishes, you will find addin_manager.exe inside the newly created dist/ folder. You can safely share this file, pin it to your taskbar, or delete the generated build/ folder and `.spec` file.

## Usage Notes
- **Network Checks**: The app checks the network drives for versions asynchronously. If you are offline or not connected to the VPN, the app will still launch instantly and simply display "Offline / Access Denied" in the version dropdowns.
- **Registry Editor**: Using the "Open Regedit Here" button temporarily updates the Registry Editor's `LastKey` memory. You may need to run the app as Administrator if you encounter permission errors using this specific button.