import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import os

from config import ENV_FILE_PATH, LOADER_PATH, ADD_IN_PATH, LOCAL_TEST_PATH, BASE_LOCAL_PATH, NETWORK_PATH_8, \
    NETWORK_PATH_6, TEMPLATES_PATH, LOG_DIR_PATH
from core.env_manager import read_env, update_env_param
from core.scanner import scan_path_sync
from core.registry_ops import scan_registry_for_ninetyone, open_regedit_at_path
from core.system_ops import open_in_explorer, launch_excel, close_excel, kill_excel
from ui.config_viewer import ConfigViewer
from ui.user_window import UserWindow
from ui.log_viewer import LogViewer


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("   FOXL Helper")
        self.root.geometry("520x590")
        self.root.resizable(False, False)

        # ttk.Style().theme_use('clam')
        self.env_var = tk.StringVar()
        self.current_version_var = tk.StringVar(value="Not set")

        self.versions = {"v8": [], "v6": [], "local": []}

        self.create_widgets()
        self.sync_ui_with_env()
        self.start_background_scans()

    def create_widgets(self):
        # --- Environment Section ---
        env_frame = ttk.LabelFrame(self.root, text="Environment Settings", padding=(10, 10))
        env_frame.pack(fill="x", padx=10, pady=5)
        env_frame.columnconfigure(2, weight=1)
        env_frame.columnconfigure(3, weight=1)

        ttk.Label(env_frame, text="Current ENV:").grid(row=0, column=0, sticky="w", pady=5)
        self.env_dd = ttk.Combobox(env_frame, textvariable=self.env_var, values=["Prod", "UAT", "Dev", "local"],
                                   state="readonly", width=18)
        self.env_dd.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(env_frame, text="Update ENV",
                   command=lambda: self.save_env_param("ENV", self.env_var.get())).grid(row=0, column=2, padx=5, sticky="ew")
        ttk.Button(env_frame, text="Open Env Folder",
                   command=lambda: self.do_explore(os.path.dirname(ENV_FILE_PATH))).grid(row=1, column=2, padx=5, sticky="ew")
        ttk.Label(env_frame, text="Current VERSION:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Label(env_frame, textvariable=self.current_version_var,
                  font=("TkDefaultFont", 9, "bold")).grid(row=1, column=1, sticky="w", padx=10, pady=5)
        ttk.Button(env_frame, text="View Configs...",
                   command=self.open_viewer).grid(row=0, column=3, padx=5, sticky="ew")
        ttk.Button(env_frame, text="User View",
                   command=self.open_user_view).grid(row=1, column=3, padx=5, pady=2, sticky="ew")

        # --- Tabs Section ---
        nb_f = ttk.Frame(self.root)
        nb_f.pack(fill="x", padx=10, pady=5)
        self.nb = ttk.Notebook(nb_f)
        self.nb.pack(fill="x", expand=True)

        self.v8_var, self.v8_dd = self.create_tab(".Net 8.0", NETWORK_PATH_8)
        self.v6_var, self.v6_dd = self.create_tab(".Net 6.0", NETWORK_PATH_6)

        # --- Paths Section ---
        path_f = ttk.LabelFrame(self.root, text="Local Directories", padding=(10, 10))
        path_f.pack(fill="x", padx=10, pady=5)
        path_f.columnconfigure(0, weight=1)
        path_f.columnconfigure(1, weight=1)

        # Row 0: Logs Buttons (Note: Your original had two identical logs buttons)
        ttk.Button(path_f, text="Open Logs Folder",
                   command=lambda: self.do_explore(os.path.dirname(LOG_DIR_PATH))).grid(row=0, column=0, sticky="ew",
                                                                                     padx=(0, 2), pady=2)
        ttk.Button(path_f, text="Support Logs", command=self.open_log_viewer).grid(row=0, column=1, sticky="ew",
                                                                                     padx=(2, 0), pady=2)

        # Row 1: Loader and v8 Add-in
        ttk.Button(path_f, text="Open Loader Path",
                   command=lambda: self.do_explore(LOADER_PATH)).grid(row=1, column=0, sticky="ew", padx=(0, 2), pady=2)
        ttk.Button(path_f, text="Open v8 Add-in Path",
                   command=lambda: self.do_explore(ADD_IN_PATH)).grid(row=1, column=1, sticky="ew", padx=(2, 0), pady=2)

        # Row 2: Test and v6 Add-in
        ttk.Button(path_f, text="Open Test Path",
                   command=lambda: self.do_explore(BASE_LOCAL_PATH)).grid(row=2, column=0, sticky="ew", padx=(0, 2),
                                                                          pady=2)
        ttk.Button(path_f, text="Open v6 Add-in Path",
                   command=lambda: self.do_explore(LOCAL_TEST_PATH)).grid(row=2, column=1, sticky="ew", padx=(2, 0),
                                                                          pady=2)

        # --- Registry Section ---
        reg_f = ttk.LabelFrame(self.root, text="Registry Check", padding=(10, 10))
        reg_f.pack(fill="x", padx=10, pady=5)
        btn_f = ttk.Frame(reg_f)
        btn_f.pack(fill="x", pady=(0, 5))
        ttk.Button(btn_f, text="Check 'NinetyOne'",
                   command=self.do_reg_check).pack(side="left", fill="x", expand=True, padx=(0, 2))
        ttk.Button(btn_f, text="Open Regedit Here",
                   command=self.do_reg_open).pack(side="right", fill="x", expand=True, padx=(2, 0))
        self.reg_out = scrolledtext.ScrolledText(reg_f, height=6, width=50, wrap=tk.WORD)
        self.reg_out.pack(fill="x")

        # --- 2. Update Excel Controls Section ---
        excel_frame = ttk.LabelFrame(self.root, text="Excel Controls", padding=(10, 10))
        excel_frame.pack(fill="x", padx=10, pady=5)

        btn_container = ttk.Frame(excel_frame)
        btn_container.pack(fill="x")

        ttk.Button(btn_container, text="Open Excel",
                   command=self.do_launch_excel).pack(side="left", expand=True, fill="x", padx=(0, 2))
        ttk.Button(btn_container, text="FOXL Templates",
                   command=lambda: self.do_explore(TEMPLATES_PATH)).pack(side="left", expand=True, fill="x", padx=(0, 2))
        ttk.Button(btn_container, text="Close Safely",
                   command=self.do_close_excel).pack(side="left", expand=True, fill="x", padx=2)
        ttk.Button(btn_container, text="Force Kill",
                   command=self.do_kill_excel).pack(side="right", expand=True, fill="x", padx=(2, 0))

    def create_tab(self, name, path):
        t = ttk.Frame(self.nb, padding=(10, 10))
        self.nb.add(t, text=name)
        ttk.Label(t, text="Available:").grid(row=0, column=0, sticky="w", pady=5)
        v = tk.StringVar(value="Checking...")
        dd = ttk.Combobox(t, textvariable=v, state="readonly", width=22)
        dd.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(t, text="Save to Env",
                   command=lambda: self.save_env_param("VERSION", v.get())).grid(row=0, column=2, padx=5)
        ttk.Button(t, text="Open Network Folder",
                   command=lambda: self.do_explore(path)).grid(row=0, column=3, padx=5, sticky="ew")
        return v, dd

    # --- Actions ---
    def sync_ui_with_env(self):
        e, v = read_env()
        self.env_var.set(e)
        self.current_version_var.set(v)

    def save_env_param(self, prefix, val):
        success, msg = update_env_param(prefix, val)
        if success:
            if prefix == "VERSION": self.current_version_var.set(val)
            messagebox.showinfo("Success", f"{prefix} updated to: {val}")
        else:
            messagebox.showerror("Error", msg)

    def start_background_scans(self):
        threading.Thread(target=self.bg_scan, args=(NETWORK_PATH_8, "v8", self.v8_var, self.v8_dd), daemon=True).start()
        threading.Thread(target=self.bg_scan, args=(NETWORK_PATH_6, "v6", self.v6_var, self.v6_dd), daemon=True).start()
        threading.Thread(target=self.bg_scan, args=(LOCAL_TEST_PATH, "local", None, None), daemon=True).start()

    def bg_scan(self, path, key, var, dd):
        success, result = scan_path_sync(path)
        if success:
            self.versions[key] = result
            if var and dd: self.root.after(0, lambda: self.update_dd(dd, var, result))
        else:
            if var and dd: self.root.after(0, lambda: self.update_dd(dd, var, [], result))

    def update_dd(self, dd, var, items, fallback="No versions"):
        dd['values'] = items
        var.set(items[0] if items else fallback)

    def open_viewer(self):
        ConfigViewer(self.root, self.env_var.get(), self.versions["v8"], self.versions["v6"], self.versions["local"])

    def open_user_view(self):
        child_window = tk.Toplevel(self.root)
        user_view = UserWindow(child_window)
        child_window.grab_set()

    def open_log_viewer(self):
        LogViewer(self.root).grab_set()

    def do_explore(self, path):
        if not open_in_explorer(path): messagebox.showwarning("Not Found", path)

    def do_reg_check(self):
        self.reg_out.delete(1.0, tk.END)
        success, result = scan_registry_for_ninetyone()
        if success:
            self.reg_out.insert(tk.END, ("Found NinetyOne:\n\n" + "\n".join(result)) if result else "Not found")
        else:
            self.reg_out.insert(tk.END, result)

    def do_reg_open(self):
        success, msg = open_regedit_at_path()
        if not success: messagebox.showerror("Error", msg)

    # --- 3. Update Excel Methods ---
    def do_launch_excel(self):
        success, msg = launch_excel()
        if not success: messagebox.showerror("Error", msg)

    def do_close_excel(self):
        success, msg = close_excel()
        if not success: messagebox.showerror("Error", msg)

    def do_kill_excel(self):
        if messagebox.askyesno("Confirm Force Kill",
                               "This will instantly kill Excel. Any unsaved work will be lost. Continue?"):
            success, msg = kill_excel()
            if not success: messagebox.showerror("Error", msg)
