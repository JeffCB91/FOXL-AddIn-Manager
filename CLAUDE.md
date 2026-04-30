# FOXL Dev Support — Claude Context

## What this project is
A Windows desktop utility for managing the Ninety One **Front Office Excel Add-In (FOXL)**.  
Two separate UIs: an **Admin** view (full config, network scanning, deployment) and a **User** view (env toggle, zip install).  
Packaged to `.exe` via PyInstaller. UI framework: plain **tkinter + ttk** (no CustomTkinter).  
Dependencies: `requests`, `pyinstaller` — everything else is stdlib.

## Project structure
```
FOXL-Dev-Support/
├── config.py                  # All hardcoded paths and constants
├── main.py                    # Admin UI entry point
├── main_user.py               # User UI entry point
├── requirements.txt           # requests, pyinstaller
├── PAT_GUIDE.md               # How to generate an ADO Personal Access Token
├── README.md
├── foxl_addin_admin.spec      # PyInstaller spec → foxl_addin_admin.exe
├── foxl_addin_user.spec       # PyInstaller spec → foxl_addin_user.exe
├── core/                      # Business logic (no UI imports)
│   ├── env_manager.py
│   ├── log_ops.py
│   ├── registry_ops.py
│   ├── scanner.py
│   ├── system_ops.py
│   ├── zip_ops.py
│   ├── ado_client.py
│   └── deploy_ops.py
├── ui/                        # tkinter windows and panels
│   ├── main_window.py
│   ├── user_window.py
│   ├── config_viewer.py
│   ├── log_viewer.py
│   └── deploy_window.py
└── tests/                     # 108 pytest tests (none existed on main before this branch)
    ├── test_env_manager.py
    ├── test_log_ops.py
    ├── test_registry_ops.py
    ├── test_scanner.py
    ├── test_system_ops.py
    └── test_zip_ops.py
```

## Configuration constants (`config.py`)
All paths and external service settings live here — never hardcode elsewhere.

| Constant | Value |
|---|---|
| `ENV_FILE_PATH` | `%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\front-office-excel-addin-env` |
| `PAT_FILE_PATH` | `%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\.env` |
| `BASE_LOCAL_PATH` | `C:\ExcelAddIn` |
| `LOG_DIR_PATH` | `C:\ProgramData\NinetyOne.ExcelAddIn\Logs` |
| `REG_PATH` | `Software\Microsoft\Office\16.0\excel\options` |
| `LOADER_PATH` | `C:\Program Files\Microsoft Office\root\Office16\Library\InvestmentTechExcelAddin` |
| `FOXL_LOADER_PATH` | `C:\Program Files\Ninety One\FOXL\v8\NinetyOne.ExcelAddIn.Loader-AddIn64.xll` |
| `ADD_IN_PATH` | `C:\Program Files\Ninety One\FOXL\v8\_91ExcelAddin` |
| `LOCAL_TEST_PATH` | `C:\ExcelAddIn\_91ExcelAddIn` |
| `NETWORK_PATH_8` | `\\iamldnfs1\GDrive\...\FrontOfficeExcelAddIn\DotNet8\InvestmentTechExcelAddIn` |
| `NETWORK_PATH_6` | `\\iamldnfs1\GDrive\...\FrontOfficeExcelAddIn\InvestmentTechExcelAddIn` |
| `TEMPLATES_PATH` | `\\uranus\FMC\Data\Internal Data\FOXL Templates` |
| `CONFIG_FILE_NAME` | `NinetyOne.ExcelAddIn.config.json` |
| `ADO_ORG` | `"Ninety-One"` |
| `ADO_PROJECT` | `"Ninety-One"` |
| `ADO_PIPELINE_ID` | `556` |
| `ADO_ARTIFACT_NAME` | `"Binaries"` |
| `ADO_ARTIFACT_FILE` | `"NetworkShareFiles.zip"` |

## Core modules (`core/`)

### `env_manager.py`
Read/write `ENV` and `VERSION` in the flat `KEY=value` env file.
- `read_env()` → `(env_str, version_str)` — defaults `"Not set"` if file missing
- `update_env_param(prefix, value)` → `(bool, error_msg)` — rejects empty / sentinel values like `"Checking..."`, `"No versions"`, `"N/A"`; creates directory if needed

### `log_ops.py`
Parse, aggregate, and export add-in log files from `LOG_DIR_PATH`.
- `extract_service(filename)` → first dot-segment of filename (e.g. `"AddIn"`)
- `format_service_name(raw)` → human display name via rename map (`LoaderLog`→`Loader`, `RiskAnalyticDataAccess`→`Risk Analytics`) and suffix stripping (`ExcelAddInService`, `ExcelAddIn`, `Service`)
- `get_services(log_files)` → sorted `[(raw, display), ...]`
- `get_log_files(log_dir)` → `.log`/`.txt` files sorted by modification time
- `generate_unified_timeline(log_files, reverse=False)` → `(entries, errors)` where `entries = [(timestamp_str, text_str), ...]`; handles continuation/stack-trace lines; timestamp format `YYYY-MM-DD HH:MM:SS`
- `export_logs_to_zip(log_files, destination_zip)` → zip archive

### `registry_ops.py`
Windows registry access via `winreg` (Windows-only).
- `scan_registry_for_ninetyone()` → `(bool, [matching_values] | error_msg)` — scans `HKCU\...\excel\options` for `"ninetyone"` (case-insensitive)
- `open_regedit_at_path()` → sets `LastKey` then launches `regedit.exe` (requires admin)
- `point_excel_to_addin(new_xll_path)` → updates `OPEN`/`OPEN1`/`OPEN2` key to `/R "<path>"`, preferring existing NinetyOne/InvestmentTechExcelAddIn entry, else next free slot

### `scanner.py`
- `scan_path_sync(path)` → `(bool, [subdirs_reverse_sorted] | error_msg)` — designed to run in background threads

### `system_ops.py`
Thin subprocess/OS wrappers (Windows-only).
- `open_in_explorer(path)` → `bool` via `os.startfile()`
- `launch_excel()` → `(bool, error_msg)` via `start excel`
- `close_excel()` → `(bool, error_msg)` via `taskkill /im excel.exe`
- `kill_excel()` → `(bool, error_msg)` via `taskkill /f /t /im excel.exe` + `wmic`

### `zip_ops.py`
Extract add-in zip archives to local path.
- `find_xll_in_folder(folder_path)` → `(bool, xll_path | error_msg)` — recursive search
- `extract_and_install_zip(zip_path, target_subfolder_name)` → `(bool, status_msg, xll_path | None)` — extracts to `C:\ExcelAddIn\<name>\`, copies JSON config if present

### `ado_client.py`
Azure DevOps REST API client. Auth: Basic with empty user + PAT. SSL verify=False (corporate proxy).
- `fetch_builds(pat, top=20)` → `(bool, [build_dicts] | error_msg)` — pipeline 556 builds
- `download_artifact_zip(pat, build_id, progress_cb=None)` → `(bool, tmp_filepath | error_msg)` — streams `NetworkShareFiles.zip`; callback: `progress_cb(downloaded, total)`

### `deploy_ops.py`
Versioned releases on network share. Version folders are `vX.Y.Z` regex.
- `get_existing_versions(network_path)` → `(bool, [version_strs] | error_msg)` — reverse-sorted
- `get_next_version(network_path)` → `(bool, version_str | error_msg)` — increments patch
- `deploy_zip_to_network(zip_path, version, network_path)` → `(bool, dest_path | error_msg)` — fails if folder exists
- `rollback_to_network(source_version, network_path)` → `(bool, (new_version, dest_path) | error_msg)` — copies existing version folder as new release

## UI modules (`ui/`)

### `main_window.py` — Admin Dashboard
- ENV/VERSION switching (Combobox + background scan)
- Network Tabs: .Net 8.0 and .Net 6.0 version lists (async `scanner.scan_path_sync`)
- Quick-access Explorer buttons: Logs, Loader, Add-in, Test path
- Registry panel: scan for NinetyOne entries, open Regedit
- Excel controls: Open / Close / Kill
- Opens `config_viewer.py` and `deploy_window.py` as modal dialogs

### `user_window.py` — User Dashboard
- ENV selection (Combobox)
- Install from Zip: browse zip → enter folder name → extract + point Excel to .xll
- Install from Folder: browse pre-extracted folder → find .xll → point Excel
- Revert to Standard: resets registry to FOXL loader path
- Excel controls: Open / Close / Kill
- Opens `log_viewer.py` as modal

### `config_viewer.py` — Config Variables Viewer
- Source dropdown: Network .Net 8 / Network .Net 6 / Local Test Path
- Version dropdown: dynamically populated
- Environment dropdown: from JSON, defaults to current ENV
- Treeview table: Parameter / Value columns
- JSON format: `[{"environment": "Dev", "parameters": {"Key": "Value"}}]`

### `log_viewer.py` — Unified Log Timeline Viewer
- Unified Timeline tab (all services merged) + File View tab (single file)
- Service filter dropdown, Oldest/Newest toggle, Ctrl+F search with highlight
- Keyword colouring: ERROR (red), WARN (orange), INFO (blue)
- Toolbar: Copy Selected, Copy All, Prettify JSON, Refresh, Compress & Save for Support
- Status bar: entry count, timestamp range, search match count
- Background thread builds timeline to prevent UI freeze

### `deploy_window.py` — ADO Build Deployment
- PAT field + "Remember PAT" checkbox (persisted to `PAT_FILE_PATH`)
- Load Builds → Treeview: Build #, Result, Finished UTC, Branch, Requested By
- Deploy selected build: resolves next version, confirms, downloads + extracts to network
- Progress bar (download %) + ScrolledText log output
- Rollback: select existing version → copy as new release
- PAT Guide button (opens `PAT_GUIDE.md`)
- Excel controls: Open / Close / Kill
- All downloads/deploys run in background threads

## Tests (`tests/`)
108 tests, no UI tests (tkinter not testable headlessly).

| File | Classes | Count |
|---|---|---|
| `test_env_manager.py` | `TestReadEnv`, `TestUpdateEnvParam` | 27 |
| `test_log_ops.py` | `TestExtractService`, `TestFormatServiceName`, `TestGetServices`, `TestGetLogFiles`, `TestGenerateUnifiedTimeline`, `TestExportLogsToZip` | 47 |
| `test_registry_ops.py` | `TestScanRegistryForNinetyone`, `TestOpenRegeditAtPath`, `TestPointExcelToAddin` | 27 |
| `test_scanner.py` | `TestScanPathSync` | 6 |
| `test_system_ops.py` | `TestOpenInExplorer`, `TestLaunchExcel`, `TestCloseExcel`, `TestKillExcel` | 8 |
| `test_zip_ops.py` | `TestFindXllInFolder`, `TestExtractAndInstallZip` | 12 |

### Test conventions
- `pytest` + `tmp_path` for all file I/O
- Constants patched via `patch("core.<module>.<CONSTANT>", tmp_path / ...)`
- `winreg` stubbed at module level in `test_registry_ops.py` so tests run cross-platform
- No real registry writes, no network, no real processes
- Philosophy: behaviour over coverage — don't test exact subprocess flag strings or stdlib wrappers

## Windows-only considerations
- `winreg`, `os.startfile`, `taskkill`, `wmic` — stub/mock in tests, never call on non-Windows
- App is only ever run and packaged on Windows
- PyInstaller builds: `foxl_addin_admin.spec` → `foxl_addin_admin.exe`, `foxl_addin_user.spec` → `foxl_addin_user.exe`
- `verify=False` on all `requests` calls — corporate proxy intercepts TLS

## Building
```bash
pyinstaller foxl_addin_admin.spec   # Admin exe
pyinstaller foxl_addin_user.spec    # User exe
```
Outputs land in `dist/`. No `-onefile` needed — specs handle everything.
