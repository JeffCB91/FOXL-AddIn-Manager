import os
import re
import shutil
import zipfile

from config import NETWORK_PATH_8

_VERSION_RE = re.compile(r"^v(\d+)\.(\d+)\.(\d+)$", re.IGNORECASE)


def _parse_versions(entries):
    versions = []
    for entry in entries:
        m = _VERSION_RE.match(entry)
        if m:
            versions.append((int(m.group(1)), int(m.group(2)), int(m.group(3))))
    return versions


def get_existing_versions(network_path=NETWORK_PATH_8):
    try:
        entries = os.listdir(network_path)
    except OSError as e:
        return False, str(e)
    versions = sorted(_parse_versions(entries), reverse=True)
    return True, [f"v{a}.{b}.{c}" for a, b, c in versions]


def get_next_version(network_path=NETWORK_PATH_8):
    try:
        entries = os.listdir(network_path)
    except OSError as e:
        return False, str(e)

    versions = _parse_versions(entries)
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


def rollback_to_network(source_version, network_path=NETWORK_PATH_8):
    ok, next_ver = get_next_version(network_path)
    if not ok:
        return False, next_ver

    source_path = os.path.join(network_path, source_version)
    dest_path = os.path.join(network_path, next_ver)

    if not os.path.isdir(source_path):
        return False, f"Source version folder not found: {source_path}"

    try:
        shutil.copytree(source_path, dest_path)
        return True, (next_ver, dest_path)
    except Exception as e:
        return False, str(e)
