import os
import re
import zipfile

from config import NETWORK_PATH_8


def get_next_version(network_path=NETWORK_PATH_8):
    try:
        entries = os.listdir(network_path)
    except OSError as e:
        return False, str(e)

    pattern = re.compile(r"^v(\d+)\.(\d+)\.(\d+)$", re.IGNORECASE)
    versions = []
    for entry in entries:
        m = pattern.match(entry)
        if m:
            versions.append((int(m.group(1)), int(m.group(2)), int(m.group(3))))

    if not versions:
        return False, "No versioned folders (vX.Y.Z) found in network path"

    a, b, c = max(versions)
    return True, f"v{a}.{b}.{c + 1}"


def deploy_zip_to_network(zip_path, version, network_path=NETWORK_PATH_8):
    dest = os.path.join(network_path, version)
    try:
        os.makedirs(dest, exist_ok=False)
    except FileExistsError:
        return False, f"Version folder already exists: {dest}"
    except OSError as e:
        return False, f"Could not create folder: {e}"

    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(dest)
        return True, dest
    except zipfile.BadZipFile:
        try:
            os.rmdir(dest)
        except OSError:
            pass
        return False, "Downloaded file is not a valid zip archive"
    except Exception as e:
        return False, str(e)
