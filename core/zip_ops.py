import os
import zipfile
import tempfile
import shutil
from config import BASE_LOCAL_PATH


def find_xll_in_folder(folder_path):
    """Searches a folder (and subdirectories) for a .xll file.

    Returns:
        tuple: (success, xll_path_or_error_msg)
    """
    if not os.path.isdir(folder_path):
        return False, "The selected path is not a valid folder."
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.xll'):
                return True, os.path.join(root, file)
    return False, "No .xll file found in the selected folder."


def extract_and_install_zip(zip_path, target_subfolder_name):
    """Extracts zip, finds .xll and .json, and moves them to the local test path."""

    final_dir = os.path.join(BASE_LOCAL_PATH, target_subfolder_name)

    try:
        os.makedirs(final_dir, exist_ok=True)

        with tempfile.TemporaryDirectory() as temp_dir:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            xll_file = None
            json_file = None

            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if file.endswith('.xll'):
                        xll_file = os.path.join(root, file)
                    elif file.endswith('.json') and "config" in file.lower():
                        json_file = os.path.join(root, file)

            if not xll_file:
                return False, "Could not find a .xll file inside the provided zip.", None

            final_xll_path = os.path.join(final_dir, os.path.basename(xll_file))
            shutil.copy2(xll_file, final_xll_path)

            if json_file:
                shutil.copy2(json_file, os.path.join(final_dir, os.path.basename(json_file)))

            return True, "Files successfully extracted and installed.", final_xll_path

    except zipfile.BadZipFile:
        return False, "The selected file is not a valid ZIP archive.", None
    except Exception as e:
        return False, f"An error occurred during extraction:\n{str(e)}", None
