import os
from config import ENV_FILE_PATH


def read_env():
    env, version = "Not set", "Not set"
    if os.path.exists(ENV_FILE_PATH):
        try:
            with open(ENV_FILE_PATH, 'r') as f:
                for line in f:
                    if line.startswith("ENV="):
                        env = line.strip().split("=")[1]
                    elif line.startswith("VERSION="):
                        version = line.strip().split("=")[1]
        except Exception as e:
            pass
    return env, version


def update_env_param(prefix, new_value):
    if not new_value or any(x in new_value for x in ["Checking", "No versions", "N/A"]):
        return False, "Invalid value selected."

    lines, found = [], False
    if os.path.exists(ENV_FILE_PATH):
        with open(ENV_FILE_PATH, 'r') as f: lines = f.readlines()

    for i, line in enumerate(lines):
        if line.startswith(f"{prefix}="):
            lines[i] = f"{prefix}={new_value}\n"
            found = True
            break

    if not found:
        lines.insert(0 if prefix == "ENV" else 1, f"{prefix}={new_value}\n")

    try:
        os.makedirs(os.path.dirname(ENV_FILE_PATH), exist_ok=True)
        with open(ENV_FILE_PATH, 'w') as f:
            f.writelines(lines)
        return True, ""
    except Exception as e:
        return False, str(e)