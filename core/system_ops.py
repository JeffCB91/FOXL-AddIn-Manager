import os
import subprocess

def open_in_explorer(path):
    if os.path.exists(path):
        os.startfile(path)
        return True
    return False

def launch_excel():
    try:
        os.startfile("excel.exe")
        return True, ""
    except Exception as e:
        return False, str(e)

def close_excel():
    """Attempts to close Excel gracefully (will prompt to save documents)."""
    try:
        # Calling taskkill without /f asks the application to close nicely
        subprocess.run(["taskkill", "/im", "excel.exe"], capture_output=True, text=True)
        return True, ""
    except Exception as e:
        return False, str(e)

def kill_excel():
    """Forces Excel to terminate immediately (useful for hangs)."""
    try:
        # The /f flag forces the process to end instantly
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"], capture_output=True, text=True)
        return True, ""
    except Exception as e:
        return False, str(e)