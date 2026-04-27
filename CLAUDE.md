# FOXL Dev Support — Claude Context

## What this project is
A Windows desktop utility for managing the Ninety One **Front Office Excel Add-In (FOXL)**.  
Two separate UIs: an **Admin** view (full config, network scanning) and a **User** view (env toggle, zip install).  
Packaged to `.exe` via PyInstaller — no pip install, just `pyinstaller`.

## Entry points
| File | Purpose |
|---|---|
| `main.py` | Admin UI entry point |
| `main_user.py` | User UI entry point |
| `config.py` | All hardcoded paths and constants (registry key, network shares, local dirs) |

## Core modules (`core/`)
| Module | What it does |
|---|---|
| `env_manager.py` | Read/write `ENV` and `VERSION` to `%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\front-office-excel-addin-env`. Simple `KEY=value` flat file. |
| `log_ops.py` | Parse `.log`/`.txt` files into a unified chronological timeline. Handles continuation lines (stack traces), service name formatting, and zip export. |
| `registry_ops.py` | Scan/update `HKCU\Software\Microsoft\Office\16.0\excel\options` for Excel add-in `OPEN`/`OPEN1`/`OPEN2` keys. Windows-only (`winreg`). |
| `scanner.py` | Scan a directory and return subdirectories (version folders) in reverse-sorted order. 11 lines. |
| `system_ops.py` | Thin OS wrappers: open Explorer, launch/close/kill Excel via subprocess. |
| `zip_ops.py` | Extract a zip to `C:\ExcelAddIn\<version>\`, find the `.xll`, copy `NinetyOne.ExcelAddIn.config.json` if present. |
| `ado_client.py` | Fetch builds and download artifacts from Azure DevOps REST API using a PAT. `fetch_builds` lists recent pipeline runs; `download_artifact_zip` streams `NetworkShareFiles.zip` to a temp file. Uses `verify=False` (corporate proxy). |
| `deploy_ops.py` | Network share deployment logic. `get_next_version` / `get_existing_versions` scan `vX.Y.Z` folders; `deploy_zip_to_network` extracts a zip into the next version slot; `rollback_to_network` copies an existing version folder as the new release. |

## UI modules (`ui/`)
- `main_window.py` — Admin dashboard
- `user_window.py` — User interface
- `config_viewer.py` — Config/path viewer panel
- `log_viewer.py` — Unified log timeline viewer
- `deploy_window.py` — ADO build deploy dialog: load builds, deploy/rollback to network share, PAT save/load, Excel open/close/kill buttons

## Configuration constants (`config.py`)
Key constants future tasks will likely reference:
- `ENV_FILE_PATH` — where env/version is stored
- `BASE_LOCAL_PATH` = `C:\ExcelAddIn`
- `REG_PATH` = `Software\Microsoft\Office\16.0\excel\options`
- `LOG_DIR_PATH` = `C:\ProgramData\NinetyOne.ExcelAddIn\Logs`
- `NETWORK_PATH_8` / `NETWORK_PATH_6` — UNC paths to versioned add-in releases
- `CONFIG_FILE_NAME` = `NinetyOne.ExcelAddIn.config.json`
- `ADO_ORG` / `ADO_PROJECT` / `ADO_PIPELINE_ID` / `ADO_ARTIFACT_NAME` / `ADO_ARTIFACT_FILE` — Azure DevOps pipeline constants
- `PAT_FILE_PATH` = `%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\.env` — persisted PAT storage

## Tests (`tests/`)
108 tests across 6 files, all added on `copilot/analyze-test-coverage`. No tests existed on `main` before that branch.

| Test file | Class(es) | Notes |
|---|---|---|
| `test_env_manager.py` | `TestReadEnv`, `TestUpdateEnvParam` | Uses `tmp_path` + `patch("core.env_manager.ENV_FILE_PATH", ...)` pattern throughout |
| `test_log_ops.py` | `TestExtractService`, `TestFormatServiceName`, `TestGetServices`, `TestGetLogFiles`, `TestGenerateUnifiedTimeline`, `TestExportLogsToZip` | Most complex module; timeline tests cover continuation lines, reverse order, unreadable files |
| `test_registry_ops.py` | `TestScanRegistryForNinetyone`, `TestOpenRegeditAtPath`, `TestPointExcelToAddin` | Defines a `winreg` stub at top of file so tests run cross-platform |
| `test_scanner.py` | `TestScanPathSync` | 6 tests, lean |
| `test_system_ops.py` | `TestOpenInExplorer`, `TestLaunchExcel`, `TestCloseExcel`, `TestKillExcel` | Thin subprocess/os wrappers; tests cover happy path + exception only |
| `test_zip_ops.py` | `TestFindXllInFolder`, `TestExtractAndInstallZip` | Uses `_make_zip` helper; patches `core.zip_ops.BASE_LOCAL_PATH` to `tmp_path` |

### Test conventions
- All tests use `pytest` with `tmp_path` fixture for file I/O
- Module-level constants patched via `patch("core.<module>.<CONSTANT>", ...)`
- `winreg` is not available on non-Windows — `test_registry_ops.py` injects a stub module before import
- No database, no network calls, no real registry writes in any test
- Test philosophy: behaviour over coverage — avoid testing stdlib wrappers or exact subprocess flag strings

## Windows-only considerations
- `winreg`, `os.startfile`, `taskkill`, `wmic` are all Windows-only
- The app is only ever run/packaged on Windows; tests stub these out for CI
- PyInstaller specs: `foxl_addin_admin.spec`, `foxl_addin_user.spec`
