import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import winreg
import threading
import json

# --- Configuration Paths ---
ENV_FILE_PATH = os.path.expandvars(r"%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\front-office-excel-addin-env")
LOADER_PATH = r"C:\Program Files\Microsoft Office\root\Office16\Library\InvestmentTechExcelAddin"
LOCAL_TEST_PATH = r"C:\ExcelAddIn\_91ExcelAddIn"
REG_PATH = r"Software\Microsoft\Office\16.0\excel\options"

# Network Paths
NETWORK_PATH_8 = r"\\iamldnfs1\GDrive\Depts\Investment IT\Investment Solutions\Software\FrontOfficeExcelAddIn\DotNet8\InvestmentTechExcelAddIn"
NETWORK_PATH_6 = r"\\iamldnfs1\GDrive\Depts\Investment IT\Investment Solutions\Software\FrontOfficeExcelAddIn\InvestmentTechExcelAddIn"
CONFIG_FILE_NAME = "NinetyOne.ExcelAddIn.config.json"


class AddinManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FOXL Add-In Manager")
        self.root.geometry("520x760")
        self.root.resizable(False, False)

        style = ttk.Style()
        style.theme_use('clam')

        self.env_var = tk.StringVar()
        self.current_version_var = tk.StringVar(value="Not set")

        # Version Caches
        self.network_versions_8 = []
        self.network_versions_6 = []
        self.local_versions = []

        self.create_widgets()
        self.load_current_env()
        self.fetch_versions_async()

    def create_widgets(self):
        # --- Environment Section ---
        env_frame = ttk.LabelFrame(self.root, text="Environment Settings", padding=(10, 10))
        env_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(env_frame, text="Current ENV:").grid(row=0, column=0, sticky="w", pady=5)
        self.env_dropdown = ttk.Combobox(env_frame, textvariable=self.env_var, values=["Prod", "UAT", "Dev", "local"],
                                         state="readonly", width=18)
        self.env_dropdown.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(env_frame, text="Update ENV", command=self.save_env).grid(row=0, column=2, padx=5)

        ttk.Label(env_frame, text="Current VERSION:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Label(env_frame, textvariable=self.current_version_var, font=("TkDefaultFont", 9, "bold")).grid(row=1,
                                                                                                            column=1,
                                                                                                            sticky="w",
                                                                                                            padx=10,
                                                                                                            pady=5)
        ttk.Button(env_frame, text="View Configs...", command=self.open_config_viewer).grid(row=1, column=2, padx=5)

        # --- Network/Local Versions Section (Tabs) ---
        notebook_frame = ttk.Frame(self.root)
        notebook_frame.pack(fill="x", padx=10, pady=5)

        self.notebook = ttk.Notebook(notebook_frame)
        self.notebook.pack(fill="x", expand=True)

        # Tab 1: .Net 8.0
        tab8 = ttk.Frame(self.notebook, padding=(10, 10))
        self.notebook.add(tab8, text=".Net 8.0")
        ttk.Label(tab8, text="Available:").grid(row=0, column=0, sticky="w", pady=5)
        self.version8_var = tk.StringVar(value="Checking network...")
        self.version8_dropdown = ttk.Combobox(tab8, textvariable=self.version8_var, state="readonly", width=22)
        self.version8_dropdown.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(tab8, text="Save to Env", command=lambda: self.save_version(self.version8_var.get())).grid(row=0,
                                                                                                              column=2,
                                                                                                              padx=5)
        ttk.Button(tab8, text="Open Folder", command=lambda: self.open_in_explorer(NETWORK_PATH_8)).grid(row=0,
                                                                                                         column=3,
                                                                                                         padx=5)

        # Tab 2: .Net 6.0
        tab6 = ttk.Frame(self.notebook, padding=(10, 10))
        self.notebook.add(tab6, text=".Net 6.0")
        ttk.Label(tab6, text="Available:").grid(row=0, column=0, sticky="w", pady=5)
        self.version6_var = tk.StringVar(value="Checking network...")
        self.version6_dropdown = ttk.Combobox(tab6, textvariable=self.version6_var, state="readonly", width=22)
        self.version6_dropdown.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(tab6, text="Save to Env", command=lambda: self.save_version(self.version6_var.get())).grid(row=0,
                                                                                                              column=2,
                                                                                                              padx=5)
        ttk.Button(tab6, text="Open Folder", command=lambda: self.open_in_explorer(NETWORK_PATH_6)).grid(row=0,
                                                                                                         column=3,
                                                                                                         padx=5)

        # --- Paths & Explorer Section ---
        paths_frame = ttk.LabelFrame(self.root, text="Local Directories", padding=(10, 10))
        paths_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(paths_frame, text="Open Env File Folder",
                   command=lambda: self.open_in_explorer(os.path.dirname(ENV_FILE_PATH))).pack(fill="x", pady=2)
        ttk.Button(paths_frame, text="Open Loader Path", command=lambda: self.open_in_explorer(LOADER_PATH)).pack(
            fill="x", pady=2)
        ttk.Button(paths_frame, text="Open Local Test Path",
                   command=lambda: self.open_in_explorer(LOCAL_TEST_PATH)).pack(fill="x", pady=2)

        # --- Registry Section ---
        reg_frame = ttk.LabelFrame(self.root, text="Registry Check", padding=(10, 10))
        reg_frame.pack(fill="both", expand=True, padx=10, pady=5)
        reg_btn_frame = ttk.Frame(reg_frame)
        reg_btn_frame.pack(fill="x", pady=(0, 5))
        ttk.Button(reg_btn_frame, text="Check 'NinetyOne'", command=self.check_registry).pack(side="left", fill="x",
                                                                                              expand=True, padx=(0, 2))
        ttk.Button(reg_btn_frame, text="Open Regedit Here", command=self.open_regedit).pack(side="right", fill="x",
                                                                                            expand=True, padx=(2, 0))
        self.reg_output = scrolledtext.ScrolledText(reg_frame, height=5, width=50, wrap=tk.WORD)
        self.reg_output.pack(fill="both", expand=True)

        # --- Actions Section ---
        action_frame = ttk.Frame(self.root, padding=(10, 10))
        action_frame.pack(fill="x", padx=10, pady=5)
        ttk.Button(action_frame, text="Launch Excel", command=self.launch_excel).pack(fill="x", ipady=5)

    # --- Core Logic ---
    def load_current_env(self):
        if not os.path.exists(ENV_FILE_PATH): return
        try:
            with open(ENV_FILE_PATH, 'r') as f:
                for line in f:
                    if line.startswith("ENV="):
                        self.env_var.set(line.strip().split("=")[1])
                    elif line.startswith("VERSION="):
                        self.current_version_var.set(line.strip().split("=")[1])
        except Exception as e:
            messagebox.showerror("Read Error", f"Could not read env file:\n{e}")

    def update_env_file_param(self, prefix, new_value):
        if not new_value or any(x in new_value for x in ["Checking", "No versions", "N/A"]): return False
        lines, found = [], False
        if os.path.exists(ENV_FILE_PATH):
            with open(ENV_FILE_PATH, 'r') as f: lines = f.readlines()
        for i, line in enumerate(lines):
            if line.startswith(f"{prefix}="):
                lines[i] = f"{prefix}={new_value}\n"
                found = True
                break
        if not found: lines.insert(0 if prefix == "ENV" else 1, f"{prefix}={new_value}\n")
        try:
            os.makedirs(os.path.dirname(ENV_FILE_PATH), exist_ok=True)
            with open(ENV_FILE_PATH, 'w') as f:
                f.writelines(lines)
            return True
        except Exception as e:
            messagebox.showerror("Write Error", str(e))
            return False

    def save_env(self):
        if self.update_env_file_param("ENV", self.env_var.get()):
            messagebox.showinfo("Success", f"ENV set to: {self.env_var.get()}")

    def save_version(self, version):
        if self.update_env_file_param("VERSION", version):
            self.current_version_var.set(version)
            messagebox.showinfo("Success", f"VERSION set to: {version}")

    # --- Async Logic ---
    def fetch_versions_async(self):
        threading.Thread(target=self._scan_path, args=(NETWORK_PATH_8, "v8"), daemon=True).start()
        threading.Thread(target=self._scan_path, args=(NETWORK_PATH_6, "v6"), daemon=True).start()
        threading.Thread(target=self._scan_path, args=(LOCAL_TEST_PATH, "local"), daemon=True).start()

    def _scan_path(self, path, key):
        try:
            if os.path.exists(path):
                v = sorted([d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))], reverse=True)
                if key == "v8":
                    self.network_versions_8 = v
                    self.root.after(0, lambda: self._update_ui_v(v, self.version8_var, self.version8_dropdown))
                elif key == "v6":
                    self.network_versions_6 = v
                    self.root.after(0, lambda: self._update_ui_v(v, self.version6_var, self.version6_dropdown))
                elif key == "local":
                    self.local_versions = v
            else:
                if key != "local": self.root.after(0, lambda: self._update_ui_err("Path not found",
                                                                                  vars(self)[f"version{key[1]}_var"],
                                                                                  vars(self)[
                                                                                      f"version{key[1]}_dropdown"]))
        except:
            if key != "local": self.root.after(0,
                                               lambda: self._update_ui_err("Error", vars(self)[f"version{key[1]}_var"],
                                                                           vars(self)[f"version{key[1]}_dropdown"]))

    def _update_ui_v(self, v, var, dd):
        if v:
            dd['values'] = v; var.set(v[0])
        else:
            dd['values'] = []; var.set("No versions found")

    def _update_ui_err(self, m, var, dd):
        dd['values'] = [];
        var.set(m)

    # --- Config Viewer ---
    def open_config_viewer(self):
        vwr = tk.Toplevel(self.root)
        vwr.title("Config Variables Viewer")
        vwr.geometry("750x480")

        ctrl = ttk.Frame(vwr, padding=(10, 10))
        ctrl.pack(fill="x")

        ttk.Label(ctrl, text="Source:").grid(row=0, column=0, padx=2)
        src_var = tk.StringVar(value="Network .Net 8")
        src_cb = ttk.Combobox(ctrl, textvariable=src_var,
                              values=["Network .Net 8", "Network .Net 6", "Local Test Path"], state="readonly",
                              width=15)
        src_cb.grid(row=0, column=1, padx=5)

        ttk.Label(ctrl, text="Version:").grid(row=0, column=2, padx=2)
        ver_var = tk.StringVar()
        ver_cb = ttk.Combobox(ctrl, textvariable=ver_var, state="readonly", width=15)
        ver_cb.grid(row=0, column=3, padx=5)

        ttk.Label(ctrl, text="Env:").grid(row=0, column=4, padx=2)
        env_cb_var = tk.StringVar()
        env_cb = ttk.Combobox(ctrl, textvariable=env_cb_var, state="readonly", width=10)
        env_cb.grid(row=0, column=5, padx=5)

        tree_f = ttk.Frame(vwr, padding=(10, 0, 10, 10))
        tree_f.pack(fill="both", expand=True)
        tree = ttk.Treeview(tree_f, columns=("P", "V"), show="headings")
        tree.heading("P", text="Parameter");
        tree.heading("V", text="Value")
        tree.column("P", width=200);
        tree.column("V", width=500)
        sb = ttk.Scrollbar(tree_f, orient="vertical", command=tree.yview);
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y");
        tree.pack(fill="both", expand=True)

        self.tmp_cfg = []

        def on_src(e=None):
            s = src_var.get()
            v = self.network_versions_8 if "8" in s else (self.network_versions_6 if "6" in s else self.local_versions)
            ver_cb['values'] = v
            if v:
                ver_cb.set(v[0]); on_ver()
            else:
                ver_cb.set("N/A"); tree.delete(*tree.get_children())

        def on_ver(e=None):
            s, vr = src_var.get(), ver_var.get()
            if vr == "N/A": return
            base = NETWORK_PATH_8 if "8" in s else (NETWORK_PATH_6 if "6" in s else LOCAL_TEST_PATH)
            p = os.path.join(base, vr, CONFIG_FILE_NAME)
            load_j(p)

        def load_j(p):
            tree.delete(*tree.get_children());
            env_cb['values'] = [];
            env_cb.set("");
            self.tmp_cfg = []
            if not os.path.exists(p): tree.insert("", "end", values=("File Not Found", p)); return
            try:
                with open(p, 'r', encoding='utf-8') as f:
                    self.tmp_cfg = json.load(f)
                en = [i.get("environment") for i in self.tmp_cfg if isinstance(i, dict)]
                env_cb['values'] = en
                if en:
                    curr = self.env_var.get()
                    env_cb.set(curr if curr in en else en[0])
                    on_env()
            except Exception as e:
                tree.insert("", "end", values=("JSON Error", str(e)))

        def on_env(e=None):
            tree.delete(*tree.get_children())
            for i in self.tmp_cfg:
                if i.get("environment") == env_cb_var.get():
                    for k, v in i.get("parameters", {}).items(): tree.insert("", "end", values=(k, v))
                    break

        src_cb.bind("<<ComboboxSelected>>", on_src);
        ver_cb.bind("<<ComboboxSelected>>", on_ver);
        env_cb.bind("<<ComboboxSelected>>", on_env)
        on_src()

    def open_in_explorer(self, p):
        if os.path.exists(p):
            os.startfile(p)
        else:
            messagebox.showwarning("Not Found", p)

    def check_registry(self):
        self.reg_output.delete(1.0, tk.END)
        try:
            k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH)
            f, i = [], 0
            while True:
                try:
                    n, v, _ = winreg.EnumValue(k, i)
                    if isinstance(v, str) and "ninetyone" in v.lower(): f.append(f"{n}: {v}")
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(k)
            self.reg_output.insert(tk.END, ("Found NinetyOne:\n\n" + "\n".join(f)) if f else "Not found")
        except Exception as e:
            self.reg_output.insert(tk.END, str(e))

    def open_regedit(self):
        try:
            tk = r"Computer\HKEY_CURRENT_USER\\" + REG_PATH
            ak = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Applets\Regedit",
                                0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(ak, "LastKey", 0, winreg.REG_SZ, tk);
            winreg.CloseKey(ak)
            os.startfile("regedit.exe")
        except:
            messagebox.showerror("Error", "Check Permissions")

    def launch_excel(self):
        try:
            os.startfile("excel.exe")
        except:
            messagebox.showerror("Error", "Excel not found")


if __name__ == "__main__":
    root = tk.Tk();
    app = AddinManagerApp(root);
    root.mainloop()