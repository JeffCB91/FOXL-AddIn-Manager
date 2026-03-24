import os

def scan_path_sync(path):
    """Synchronous scan, meant to be called from a background thread."""
    try:
        if os.path.exists(path):
            versions = sorted([d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))], reverse=True)
            return True, versions
        return False, "Path not found"
    except Exception:
        return False, "Error / Offline"
