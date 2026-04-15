import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import json
import os
import threading

from config import LOG_DIR_PATH
from core.log_ops import get_log_files, generate_unified_timeline, export_logs_to_zip


class LogViewer(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("FOXL Support Log Viewer")
        self.geometry("900x600")

        self.log_files = get_log_files(LOG_DIR_PATH)
        self.create_widgets()
        self.load_data()

    def create_widgets(self):
        # --- Top Toolbar ---
        toolbar = ttk.Frame(self, padding=(10, 5))
        toolbar.pack(fill="x")

        ttk.Button(toolbar, text="Copy Selected", command=self.copy_selection).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Copy All", command=self.copy_all).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Prettify Selected JSON", command=self.prettify_json).pack(side="left", padx=(10, 2))
        ttk.Button(toolbar, text="Refresh", command=self.refresh).pack(side="left", padx=(10, 2))
        ttk.Button(toolbar, text="Compress & Save Logs for Support", command=self.export_zip).pack(side="right", padx=2)

        # --- Tabs ---
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=5)

        # Tab 1: Unified Timeline
        self.tab_unified = ttk.Frame(self.nb)
        self.nb.add(self.tab_unified, text="Unified Timeline")
        self.txt_unified = scrolledtext.ScrolledText(self.tab_unified, wrap=tk.WORD, font=("Consolas", 9))
        self.txt_unified.pack(fill="both", expand=True, padx=5, pady=5)

        # Tab 2: File View
        self.tab_file = ttk.Frame(self.nb)
        self.nb.add(self.tab_file, text="File View")

        file_ctrl = ttk.Frame(self.tab_file)
        file_ctrl.pack(fill="x", padx=5, pady=5)
        ttk.Label(file_ctrl, text="Select Log File:").pack(side="left", padx=5)

        self.file_var = tk.StringVar()
        self.file_dd = ttk.Combobox(file_ctrl, textvariable=self.file_var, state="readonly", width=50)
        self.file_dd.pack(side="left", padx=5)
        self.file_dd.bind("<<ComboboxSelected>>", self.on_file_select)

        self.txt_file = scrolledtext.ScrolledText(self.tab_file, wrap=tk.WORD, font=("Consolas", 9))
        self.txt_file.pack(fill="both", expand=True, padx=5, pady=5)

    def load_data(self):
        self.txt_unified.delete(1.0, tk.END)

        if not self.log_files:
            self.txt_unified.insert(tk.END, f"No logs found in:\n{LOG_DIR_PATH}")
            return

        self.txt_unified.insert(tk.END, "Compiling unified timeline...\n")

        # Populate file dropdown immediately (fast)
        filenames = [os.path.basename(f) for f in self.log_files]
        self.file_dd['values'] = filenames
        if filenames:
            self.file_dd.set(filenames[0])
            self.on_file_select()

        # Parse and merge logs in background to avoid freezing the UI
        threading.Thread(target=self._load_unified_bg, daemon=True).start()

    def _load_unified_bg(self):
        entries, errors = generate_unified_timeline(self.log_files)

        if entries:
            data = "".join(e[1] for e in entries)
        else:
            data = "No logs found or unable to parse timestamps."

        if errors:
            error_lines = "\n".join(f"  {name}: {msg}" for name, msg in errors)
            data += f"\n\n--- Files with parse errors ---\n{error_lines}\n"

        def _update():
            self.txt_unified.delete(1.0, tk.END)
            self.txt_unified.insert(tk.END, data)

        self.after(0, _update)

    def refresh(self):
        self.log_files = get_log_files(LOG_DIR_PATH)
        self.txt_file.delete(1.0, tk.END)
        self.file_dd['values'] = []
        self.file_var.set("")
        self.load_data()

    def on_file_select(self, event=None):
        selected_name = self.file_var.get()
        target_path = next((f for f in self.log_files if os.path.basename(f) == selected_name), None)

        self.txt_file.delete(1.0, tk.END)
        if target_path and os.path.exists(target_path):
            try:
                with open(target_path, 'r', encoding='utf-8', errors='replace') as f:
                    self.txt_file.insert(tk.END, f.read())
            except OSError as e:
                self.txt_file.insert(tk.END, f"Could not read file:\n{e}")

    # --- Toolbar Actions ---
    def get_active_textbox(self):
        """Returns whichever text box is currently visible based on the active tab."""
        return self.txt_unified if self.nb.index("current") == 0 else self.txt_file

    def copy_selection(self):
        txt = self.get_active_textbox()
        try:
            selected = txt.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.clipboard_clear()
            self.clipboard_append(selected)
        except tk.TclError:
            messagebox.showwarning("Warning", "No text selected.")

    def copy_all(self):
        txt = self.get_active_textbox()
        self.clipboard_clear()
        self.clipboard_append(txt.get(1.0, tk.END))

    def prettify_json(self):
        txt = self.get_active_textbox()
        try:
            start_idx = txt.index(tk.SEL_FIRST)
            end_idx = txt.index(tk.SEL_LAST)
            selected_text = txt.get(start_idx, end_idx)

            parsed = json.loads(selected_text)
            pretty_json = json.dumps(parsed, indent=4)

            txt.delete(start_idx, end_idx)
            txt.insert(start_idx, pretty_json)

        except tk.TclError:
            messagebox.showwarning("Warning", "Please highlight the JSON text you want to format first.")
        except json.JSONDecodeError:
            messagebox.showerror("Error", "The selected text is not valid JSON.")

    def export_zip(self):
        if not self.log_files:
            messagebox.showwarning("Warning", "No log files to compress.")
            return

        dest = filedialog.asksaveasfilename(
            title="Save Support Logs",
            defaultextension=".zip",
            initialfile="FOXL_Support_Logs.zip",
            filetypes=[("Zip files", "*.zip")]
        )

        if dest:
            success, msg = export_logs_to_zip(self.log_files, dest)
            if success:
                messagebox.showinfo("Success", f"Logs successfully compressed to:\n{dest}")
            else:
                messagebox.showerror("Error", f"Failed to compress logs:\n{msg}")
