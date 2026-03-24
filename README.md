# FOXL Add-In Manager

A lightweight, standalone desktop utility built in Python to help developers and users manage their Front Office Excel Add-In environments and configuration paths.

## Project Structure

The codebase is modularized to separate the user interface from the business logic:

```text
foxl-manager/
├── main.py                  # The entry point that launches the app
├── config.py                # Stores all hardcoded paths, URLs, and constants
│
├── core/                    # Core business logic
│   ├── env_manager.py       # Reads and writes to the local .env file
│   ├── scanner.py           # Asynchronously scans network and local directories
│   ├── registry_ops.py      # Handles Windows Registry scanning
│   └── system_ops.py        # Helper functions (Explorer, Excel launchers)
│
└── ui/                      # Presentation layer (Tkinter)
    ├── main_window.py       # The primary dashboard
    └── config_viewer.py     # The dynamic JSON configuration viewer
```

## Prerequisites

To run the script from the source code or build a new executable, you will need:
* **Python 3.x** installed on your Windows machine.
* **Windows OS** (This script utilizes `winreg` and `os.startfile`, which are Windows-specific).

## Installation & Setup

1. Clone or download this repository to your local machine.
2. Open a terminal in the project directory.
3. Install the build requirements (PyInstaller):
```bash
  pip install -r requirements.txt
```

## Running from Source

To run the application directly via Python during development:
```bash
  python main.py
```

## Building the Executable (.exe)

To share this tool with team members who do not have Python installed, you can compile it into a single standalone `.exe` file. 

Run the following command in your terminal from the root `foxl-manager` directory:
```bash
  pyinstaller --onefile --windowed --name "foxl_addin_manager" main.py
```

* `--onefile`: Packages everything into a single `.exe`.
* `--windowed`: Prevents the black command prompt console from appearing behind the UI.
* `--name`: Forces PyInstaller to name the output file `foxl_addin_manager.exe` instead of `main.exe`.

Once the build finishes, you will find `foxl_addin_manager.exe` inside the newly created `dist/` folder. 

### Cleaning Up Build Files

PyInstaller generates several temporary folders and files during the build process. To clean up your workspace after grabbing your `.exe` from the `dist/` folder, you can run the following commands in your Windows terminal:

**Command Prompt (cmd):**
```cmd
rmdir /s /q build dist
del /q *.spec
```

**PowerShell:**
```powershell
Remove-Item -Recurse -Force build, dist, *.spec -ErrorAction SilentlyContinue
```

## Usage Notes
* **Network Checks:** The app checks the network drives for versions asynchronously. If you are offline or not connected to the VPN, the app will launch instantly and simply display "Offline / Access Denied".
* **Config Viewer:** You can inspect the parsed JSON parameters for any environment or version (both local and network) by clicking the "View Configs..." button.
* **Registry Editor:** Using the "Open Regedit Here" button temporarily updates the Registry Editor's `LastKey` memory. You may need to run the app as Administrator if you encounter permission errors using this specific button.