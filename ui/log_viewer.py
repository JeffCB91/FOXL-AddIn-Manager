import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import json
import os
import threading

from config import LOG_DIR_PATH
from core.log_ops import (
    get_log_files, get_services, extract_service,
    generate_unified_timeline, export_logs_to_zip,
)

# Keyword → (tag name, foreground colour)
_HL_KEYWORDS = {
    'ERROR': ('hl_error', '#e05252'),
    'WARN':  ('hl_warn',  '#c8960c'),
    'INFO':  ('hl_info',  '#5b9bd5'),
}


class LogViewer(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("FOXL Support Log Viewer")
        self.geometry("900x640")

        self.log_files = get_log_files(LOG_DIR_PATH)
        self._service_filter = None   # None = all services; otherwise the raw service name
        self._service_map = {}        # display name → raw name (populated in load_data)
        self._last_entries = []       # most-recently-loaded (timestamp, text) list
        self._search_positions = []   # list of Tkinter index strings for current matches
        self._search_idx = -1
        self._search_query = ""

        self.create_widgets()
        self.load_data()

    # ------------------------------------------------------------------
    # Widget construction
    # ------------------------------------------------------------------

    def create_widgets(self):
        # --- Toolbar ---
        toolbar = ttk.Frame(self, padding=(10, 5))
        toolbar.pack(fill="x")

        ttk.Button(toolbar, text="Copy Selected", command=self.copy_selection).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Copy All", command=self.copy_all).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Prettify Selected JSON", command=self.prettify_json).pack(side="left", padx=(10, 2))
        ttk.Button(toolbar, text="Refresh", command=self.refresh).pack(side="left", padx=(10, 2))
        ttk.Button(toolbar, text="Compress & Save Logs for Support", command=self.export_zip).pack(side="right", padx=2)

        # --- Filter Bar ---
        filter_bar = ttk.Frame(self, padding=(10, 3))
        filter_bar.pack(fill="x")

        ttk.Label(filter_bar, text="Service:").pack(side="left")
        self.service_var = tk.StringVar(value="All Services")
        self.service_dd = ttk.Combobox(filter_bar, textvariable=self.service_var, state="readonly", width=16)
        self.service_dd.pack(side="left", padx=(3, 12))
        self.service_dd.bind("<<ComboboxSelected>>", self._on_service_change)

        ttk.Label(filter_bar, text="Order:").pack(side="left")
        self.order_var = tk.StringVar(value="Oldest First")
        order_dd = ttk.Combobox(filter_bar, textvariable=self.order_var,
                                values=["Oldest First", "Newest First"],
                                state="readonly", width=12)
        order_dd.pack(side="left", padx=(3, 12))
        order_dd.bind("<<ComboboxSelected>>", self._on_order_change)

        ttk.Label(filter_bar, text="Find:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(filter_bar, textvariable=self.search_var, width=24)
        self.search_entry.pack(side="left", padx=(3, 2))
        self.search_entry.bind("<Return>", self._search)
        ttk.Button(filter_bar, text="▲", width=2, command=self._find_prev).pack(side="left", padx=1)
        ttk.Button(filter_bar, text="▼", width=2, command=self._find_next).pack(side="left", padx=1)

        self.bind("<Control-f>", self._focus_search)

        # --- Tabs ---
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=5)
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_change)

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

        # --- Status Bar ---
        self.status_var = tk.StringVar(value="Loading...")
        ttk.Label(self, textvariable=self.status_var, relief="sunken",
                  anchor="w", padding=(6, 2)).pack(fill="x", padx=10, pady=(0, 5))

    # ------------------------------------------------------------------
    # Data loading
    # ------------------------------------------------------------------

    def load_data(self):
        """Full reload: rebuilds service list, file dropdown, and unified timeline."""
        services = get_services(self.log_files)
        self._service_map = {"All Services": None}
        self._service_map.update({display: raw for raw, display in services})
        self.service_dd['values'] = list(self._service_map.keys())
        self.service_var.set("All Services")
        self._service_filter = None

        self._reload_file_dropdown()
        self._reload_unified()

    def _get_filtered_files(self):
        """Returns log_files restricted to the current service filter."""
        if self._service_filter is None:
            return self.log_files
        return [f for f in self.log_files
                if extract_service(os.path.basename(f)) == self._service_filter]

    def _reload_unified(self):
        """Kicks off a background thread to rebuild the unified timeline."""
        self.txt_unified.delete(1.0, tk.END)
        self.txt_unified.insert(tk.END, "Compiling unified timeline...\n")
        self.status_var.set("Loading...")
        threading.Thread(target=self._load_unified_bg, daemon=True).start()

    def _load_unified_bg(self):
        """Background worker — must not touch Tkinter widgets directly."""
        filtered_files = self._get_filtered_files()
        reverse = (self.order_var.get() == "Newest First")
        entries, errors = generate_unified_timeline(filtered_files, reverse=reverse)
        self._last_entries = entries

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
            self._apply_highlights(self.txt_unified)
            self._update_status()

        self.after(0, _update)

    def _reload_file_dropdown(self):
        """Repopulates the File View dropdown based on the current service filter."""
        filtered = self._get_filtered_files()
        filenames = [os.path.basename(f) for f in filtered]
        self.file_dd['values'] = filenames
        if filenames:
            self.file_dd.set(filenames[0])
            self.on_file_select()
        else:
            self.file_dd.set("")
            self.txt_file.delete(1.0, tk.END)

    def refresh(self):
        """Re-discovers log files on disk and reloads everything."""
        self.log_files = get_log_files(LOG_DIR_PATH)
        self.txt_file.delete(1.0, tk.END)
        self._clear_search()
        self.load_data()

    def on_file_select(self, event=None):
        selected_name = self.file_var.get()
        filtered = self._get_filtered_files()
        target_path = next((f for f in filtered if os.path.basename(f) == selected_name), None)

        self.txt_file.delete(1.0, tk.END)
        if target_path and os.path.exists(target_path):
            try:
                with open(target_path, 'r', encoding='utf-8', errors='replace') as f:
                    self.txt_file.insert(tk.END, f.read())
                self._apply_highlights(self.txt_file)
                self._update_status()
            except OSError as e:
                self.txt_file.insert(tk.END, f"Could not read file:\n{e}")

    # ------------------------------------------------------------------
    # Filter event handlers
    # ------------------------------------------------------------------

    def _on_service_change(self, event=None):
        selected = self.service_var.get()
        self._service_filter = self._service_map.get(selected)
        self._clear_search()
        self._reload_unified()
        self._reload_file_dropdown()

    def _on_order_change(self, event=None):
        self._reload_unified()

    def _on_tab_change(self, event=None):
        self._clear_search()
        self._update_status()

    # ------------------------------------------------------------------
    # Keyword highlighting
    # ------------------------------------------------------------------

    def _apply_highlights(self, txt):
        """Apply ERROR / WARN / INFO color tags to a text widget."""
        for keyword, (tag, color) in _HL_KEYWORDS.items():
            txt.tag_config(tag, foreground=color)
            txt.tag_remove(tag, '1.0', tk.END)
            start = '1.0'
            while True:
                pos = txt.search(keyword, start, stopindex=tk.END)
                if not pos:
                    break
                end = f"{pos}+{len(keyword)}c"
                txt.tag_add(tag, pos, end)
                start = end

    # ------------------------------------------------------------------
    # Search / Find
    # ------------------------------------------------------------------

    def _focus_search(self, event=None):
        self.search_entry.focus_set()
        return "break"

    def _search(self, event=None):
        self._clear_search_tags()
        query = self.search_var.get().strip()
        self._search_query = query
        if not query:
            self._update_status()
            return

        txt = self.get_active_textbox()
        self._search_positions = []
        start = '1.0'
        while True:
            pos = txt.search(query, start, stopindex=tk.END, nocase=True)
            if not pos:
                break
            end = f"{pos}+{len(query)}c"
            self._search_positions.append(pos)
            start = end

        txt.tag_config('hl_search', background='#ffff88', foreground='black')
        for pos in self._search_positions:
            txt.tag_add('hl_search', pos, f"{pos}+{len(query)}c")

        self._search_idx = 0 if self._search_positions else -1
        if self._search_positions:
            self._scroll_to_current(txt)

        self._update_status()

    def _find_next(self):
        if not self._search_positions:
            self._search()
            return
        self._search_idx = (self._search_idx + 1) % len(self._search_positions)
        self._scroll_to_current(self.get_active_textbox())
        self._update_status()

    def _find_prev(self):
        if not self._search_positions:
            self._search()
            return
        self._search_idx = (self._search_idx - 1) % len(self._search_positions)
        self._scroll_to_current(self.get_active_textbox())
        self._update_status()

    def _scroll_to_current(self, txt):
        if not self._search_positions or self._search_idx < 0:
            return
        pos = self._search_positions[self._search_idx]
        end = f"{pos}+{len(self._search_query)}c"
        txt.tag_remove('hl_search_cur', '1.0', tk.END)
        txt.tag_config('hl_search_cur', background='#ff9900', foreground='black')
        txt.tag_add('hl_search_cur', pos, end)
        txt.see(pos)

    def _clear_search_tags(self):
        for txt in (self.txt_unified, self.txt_file):
            for tag in ('hl_search', 'hl_search_cur'):
                txt.tag_remove(tag, '1.0', tk.END)

    def _clear_search(self):
        self._clear_search_tags()
        self._search_positions = []
        self._search_idx = -1
        self._search_query = ""

    # ------------------------------------------------------------------
    # Status bar
    # ------------------------------------------------------------------

    def _update_status(self):
        query = self._search_query
        if query and self._search_positions:
            i = self._search_idx + 1
            n = len(self._search_positions)
            self.status_var.set(f"{i} of {n} matches for '{query}'")
            return
        if query and not self._search_positions:
            self.status_var.set(f"No matches for '{query}'")
            return

        if self.nb.index('current') == 0:
            entries = self._last_entries
            if entries:
                timestamps = [e[0] for e in entries]
                self.status_var.set(
                    f"{len(entries)} entries  │  {min(timestamps)}  →  {max(timestamps)}"
                )
            else:
                self.status_var.set("No entries")
        else:
            filename = self.file_var.get()
            if filename:
                line_count = int(self.txt_file.index('end-1c').split('.')[0])
                self.status_var.set(f"{filename}  ({line_count} lines)")
            else:
                self.status_var.set("")

    # ------------------------------------------------------------------
    # Toolbar actions
    # ------------------------------------------------------------------

    def get_active_textbox(self):
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
