import winreg
import os
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


def point_excel_to_addin(new_xll_path):
    """Finds the existing FOXL/NinetyOne addin in the registry and replaces it with the new path."""
    try:
        # Open with READ and WRITE access
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_READ | winreg.KEY_WRITE)

        i = 0
        target_value_name = None
        existing_opens = []

        # Enumerate all values to find existing add-ins
        while True:
            try:
                name, value, _ = winreg.EnumValue(key, i)
                if name.startswith("OPEN"):
                    existing_opens.append(name)
                    # Look for the loader or an existing local ninetyone addin
                    if isinstance(value, str) and (
                            "ninetyone" in value.lower() or "investmenttechexceladdin" in value.lower()):
                        target_value_name = name
                i += 1
            except OSError:
                break  # Reached the end of the keys

        # Excel requires the /R flag and quotes around the path
        new_value_data = f'/R "{new_xll_path}"'

        if target_value_name:
            # Overwrite the existing FOXL add-in slot
            winreg.SetValueEx(key, target_value_name, 0, winreg.REG_SZ, new_value_data)
        else:
            # If no FOXL add-in was found, find the next available OPEN slot (e.g., OPEN, OPEN1, OPEN2)
            if "OPEN" not in existing_opens:
                target_value_name = "OPEN"
            else:
                num = 1
                while f"OPEN{num}" in existing_opens:
                    num += 1
                target_value_name = f"OPEN{num}"

            winreg.SetValueEx(key, target_value_name, 0, winreg.REG_SZ, new_value_data)

        winreg.CloseKey(key)
        return True, f"Registry updated. Excel will load:\n{new_xll_path}"

    except PermissionError:
        return False, "Run as Administrator required to modify Excel Registry."
    except Exception as e:
        return False, f"Failed to update registry:\n{str(e)}"
