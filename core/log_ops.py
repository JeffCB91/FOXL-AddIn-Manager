import os
import zipfile
import re


def get_log_files(log_dir):
    """Returns a list of all .log or .txt files in the log directory."""
    if not os.path.exists(log_dir):
        return []
    return [os.path.join(log_dir, f) for f in os.listdir(log_dir) if f.endswith(('.log', '.txt'))]


def generate_unified_timeline(log_files):
    """Reads all logs, parses their timestamps, and sorts them chronologically."""
    all_entries = []
    # Assumes standard log format starting with YYYY-MM-DD HH:MM:SS
    time_pattern = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})")

    for filepath in log_files:
        filename = os.path.basename(filepath)
        try:
            with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                current_time = "1970-01-01 00:00:00"
                current_block = []

                for line in f:
                    match = time_pattern.search(line)
                    if match:
                        # Save the previous block
                        if current_block:
                            all_entries.append((current_time, "".join(current_block)))
                        # Start a new block
                        current_time = match.group(1)
                        current_block = [f"[{filename}] {line}"]
                    else:
                        # Append to current block (e.g., stack traces)
                        current_block.append(line)

                # Catch the last block in the file
                if current_block:
                    all_entries.append((current_time, "".join(current_block)))
        except Exception:
            pass

    # Sort all entries across all files by their parsed timestamp
    all_entries.sort(key=lambda x: x[0])

    # Return as a single massive string
    return "".join(
        [entry[1] for entry in all_entries]) if all_entries else "No logs found or unable to parse timestamps."


def export_logs_to_zip(log_files, destination_zip):
    """Compresses the provided log files into a single zip archive."""
    try:
        with zipfile.ZipFile(destination_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in log_files:
                zipf.write(file, os.path.basename(file))
        return True, ""
    except Exception as e:
        return False, str(e)
