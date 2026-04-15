import os
import subprocess


def open_in_explorer(path):
    if os.path.exists(path):
        os.startfile(path)
        return True
    return False


def launch_excel():
    try:
        # Using CMD 'start' with shell=True explicitly breaks the parent-child link.
        # Excel will now launch as a completely independent process.
        subprocess.Popen("start excel", shell=True)
        return True, ""
    except Exception as e:
        return False, str(e)


def close_excel():
    """Attempts to close Excel gracefully."""
    try:
        subprocess.run(["taskkill", "/im", "excel.exe"], capture_output=True, text=True)
        return True, ""
    except Exception as e:
        return False, str(e)


def kill_excel():
    """Nuclear option to instantly terminate Excel and all related process trees."""
    try:
        # 1. Force kill the process AND its entire tree (/t)
        subprocess.run(["taskkill", "/f", "/t", "/im", "excel.exe"], capture_output=True, text=True)

        # 2. Fallback: Windows Management Instrumentation (WMI)
        # This bypasses standard task management and kills the process at the system level
        subprocess.run(["wmic", "process", "where", "name='excel.exe'", "call", "terminate"], capture_output=True,
                       text=True)

        return True, ""
    except Exception as e:
        return False, str(e)
