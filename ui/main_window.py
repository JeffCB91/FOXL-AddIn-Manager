import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import os

from config import ENV_FILE_PATH, LOADER_PATH, LOCAL_TEST_PATH, NETWORK_PATH_8, NETWORK_PATH_6
from core.env_manager import read_env, update_env_param
from core.scanner import scan_path_sync
from core.registry_ops import scan_registry_for_ninetyone, open_regedit_at_path
from core.system_ops import open_in_explorer, launch_excel
from ui.config_viewer import ConfigViewer


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("FOXL Add-In Manager")
        self.root.geometry("520x650")
        self.root.resizable(False, False)

        ttk.Style().theme_use('clam')

        self.env_var = tk.StringVar()
        self.current_version_var = tk.StringVar(value="Not set")

        self.versions = {"v8": [], "v6": [], "local": []}

        self.create_widgets()
        self.sync_ui_with_env()
        self.start_background_scans()

    def create_widgets(self):
        # --- Environment Section ---
        env_f = ttk.LabelFrame(self.root, text="Environment Settings", padding=(10, 10))
        env_f.pack(fill="x", padx=10, pady=5)
        ttk.Label(env_f, text="Current ENV:").grid(row=0, column=0, sticky="w", pady=5)
        self.env_dd = ttk.Combobox(env_f, textvariable=self.env_var, values=["Prod", "UAT", "Dev", "local"],
                                   state="readonly", width=18)
        self.env_dd.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(env_f, text="Update ENV", command=lambda: self.save_env_param("ENV", self.env_var.get())).grid(row=0,
                                                                                                                  column=2,
                                                                                                                  padx=5)
        ttk.Label(env_f, text="Current VERSION:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Label(env_f, textvariable=self.current_version_var, font=("TkDefaultFont", 9, "bold")).grid(row=1, column=1,
                                                                                                        sticky="w",
                                                                                                        padx=10, pady=5)
        ttk.Button(env_f, text="View Configs...", command=self.open_viewer).grid(row=0, column=3, padx=5)

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
        ttk.Button(path_f, text="Open Env File Folder",
                   command=lambda: self.do_explore(os.path.dirname(ENV_FILE_PATH))).pack(fill="x", pady=2)
        ttk.Button(path_f, text="Open Loader Path", command=lambda: self.do_explore(LOADER_PATH)).pack(fill="x", pady=2)
        ttk.Button(path_f, text="Open Local Test Path", command=lambda: self.do_explore(LOCAL_TEST_PATH)).pack(fill="x",
                                                                                                               pady=2)

        # --- Registry Section ---
        reg_f = ttk.LabelFrame(self.root, text="Registry Check", padding=(10, 10))
        reg_f.pack(fill="both", expand=True, padx=10, pady=5)
        btn_f = ttk.Frame(reg_f)
        btn_f.pack(fill="x", pady=(0, 5))
        ttk.Button(btn_f, text="Check 'NinetyOne'", command=self.do_reg_check).pack(side="left", fill="x", expand=True,
                                                                                    padx=(0, 2))
        ttk.Button(btn_f, text="Open Regedit Here", command=self.do_reg_open).pack(side="right", fill="x", expand=True,
                                                                                   padx=(2, 0))

        # Changed height from 5 to 2
        self.reg_out = scrolledtext.ScrolledText(reg_f, height=2, width=50, wrap=tk.WORD)
        self.reg_out.pack(fill="both", expand=True)

        # --- Action Section ---
        act_f = ttk.Frame(self.root, padding=(10, 10))
        act_f.pack(fill="x", padx=10, pady=5)
        ttk.Button(act_f, text="Launch Excel", command=self.do_launch_excel).pack(fill="x", ipady=5)

    def create_tab(self, name, path):
        t = ttk.Frame(self.nb, padding=(10, 10))
        self.nb.add(t, text=name)
        ttk.Label(t, text="Available:").grid(row=0, column=0, sticky="w", pady=5)
        v = tk.StringVar(value="Checking...")
        dd = ttk.Combobox(t, textvariable=v, state="readonly", width=22)
        dd.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(t, text="Save to Env", command=lambda: self.save_env_param("VERSION", v.get())).grid(row=0, column=2,
                                                                                                        padx=5)
        ttk.Button(t, text="Open Folder", command=lambda: self.do_explore(path)).grid(row=0, column=3, padx=5)
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

    def do_launch_excel(self):
        success, msg = launch_excel()
        if not success: messagebox.showerror("Error", msg)
