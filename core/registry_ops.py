import winreg
import subprocess
from config import REG_PATH


def scan_registry_for_ninetyone():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH)
        found, i = [], 0
        while True:
            try:
                name, value, _ = winreg.EnumValue(key, i)
                if isinstance(value, str) and "ninetyone" in value.lower():
                    found.append(f"{name}: {value}")
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
        return True, found
    except FileNotFoundError:
        return False, f"Registry key not found:\n{REG_PATH}"
    except Exception as e:
        return False, str(e)


def open_regedit_at_path():
    try:
        # Fix 1: Properly format the path with a single backslash
        target_key = f"Computer\\HKEY_CURRENT_USER\\{REG_PATH}"

        # Update Regedit's LastKey memory
        applet_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(applet_key, "LastKey", 0, winreg.REG_SZ, target_key)
        winreg.CloseKey(applet_key)

        # Fix 2: Use subprocess with the /m flag to force a new instance
        subprocess.Popen(["regedit.exe", "/m"])
        return True, ""
    except PermissionError:
        return False, "Run as Administrator required to modify Regedit's LastKey memory."
    except Exception as e:
        return False, str(e)