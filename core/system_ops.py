import os

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
