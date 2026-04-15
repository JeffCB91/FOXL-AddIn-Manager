import os
import zipfile
import re

# Explicit renames for service identifiers that need specific display names
_RENAME_MAP = {
    'LoaderLog': 'Loader',
    'RiskAnalyticDataAccess': 'Risk Analytics',
}

# Suffixes stripped to shorten service names when no explicit rename exists
_STRIP_SUFFIXES = ('ExcelAddInService', 'ExcelAddIn', 'Service')


def extract_service(filename):
    """Returns the first dot-delimited segment from a log filename.
    e.g. 'AddIn' from 'AddIn.EXCEL.12345.log'
    """
    return filename.split('.')[0]


def format_service_name(raw):
    """Returns a short display name for a raw service identifier."""
    if raw in _RENAME_MAP:
        return _RENAME_MAP[raw]
    for suffix in _STRIP_SUFFIXES:
        if raw.endswith(suffix):
            return raw[:-len(suffix)]
    return raw


def get_services(log_files):
    """Returns a sorted list of (raw, display) service name tuples derived from the given log files."""
    raw_names = sorted({extract_service(os.path.basename(f)) for f in log_files})
    return [(raw, format_service_name(raw)) for raw in raw_names]


def get_log_files(log_dir):
    """Returns a list of all .log or .txt files in the log directory, sorted by modification time."""
    if not os.path.exists(log_dir):
        return []
    files = [os.path.join(log_dir, f) for f in os.listdir(log_dir) if f.endswith(('.log', '.txt'))]
    return sorted(files, key=os.path.getmtime)


def generate_unified_timeline(log_files, reverse=False):
    """Reads all logs, parses their timestamps, and sorts them chronologically.

    Args:
        log_files: List of log file paths to process.
        reverse:   If True, return entries newest-first.

    Returns:
        tuple: (entries, errors) where entries is a list of (timestamp_str, text_str) tuples
               and errors is a list of (filename, error_str) tuples for files that failed to read.
    """
    all_entries = []
    errors = []
    time_pattern = re.compile(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})")

    for filepath in log_files:
        filename = os.path.basename(filepath)
        try:
            with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                current_time = None
                current_block = []

                for line in f:
                    match = time_pattern.search(line)
                    if match:
                        if current_block and current_time is not None:
                            all_entries.append((current_time, "".join(current_block)))
                        current_time = match.group(1)
                        current_block = [f"[{filename}] {line}"]
                    elif current_time is not None:
                        # Continuation line (e.g. stack trace) — attach to current block
                        current_block.append(line)
                    # Lines before the first timestamp are discarded

                if current_block and current_time is not None:
                    all_entries.append((current_time, "".join(current_block)))

        except Exception as e:
            errors.append((filename, str(e)))

    all_entries.sort(key=lambda x: x[0], reverse=reverse)
    return all_entries, errors


def export_logs_to_zip(log_files, destination_zip):
    """Compresses the provided log files into a single zip archive."""
    try:
        with zipfile.ZipFile(destination_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in log_files:
                zipf.write(file, os.path.basename(file))
        return True, ""
    except Exception as e:
        return False, str(e)
