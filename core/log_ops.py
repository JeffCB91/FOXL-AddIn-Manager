import os
import zipfile
import re


def get_log_files(log_dir):
    """Returns a list of all .log or .txt files in the log directory, sorted by modification time."""
    if not os.path.exists(log_dir):
        return []
    files = [os.path.join(log_dir, f) for f in os.listdir(log_dir) if f.endswith(('.log', '.txt'))]
    return sorted(files, key=os.path.getmtime)


def generate_unified_timeline(log_files):
    """Reads all logs, parses their timestamps, and sorts them chronologically.

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

    all_entries.sort(key=lambda x: x[0])
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
