import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import winreg
import threading

# --- Configuration Paths ---
ENV_FILE_PATH = os.path.expandvars(r"%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\front-office-excel-addin-env")
LOADER_PATH = r"C:\Program Files\Microsoft Office\root\Office16\Library\InvestmentTechExcelAddin"
LOCAL_TEST_PATH = r"C:\ExcelAddIn\_91ExcelAddIn"
REG_PATH = r"Software\Microsoft\Office\16.0\excel\options"

# Network Paths
NETWORK_PATH_8 = r"\\iamldnfs1\GDrive\Depts\Investment IT\Investment Solutions\Software\FrontOfficeExcelAddIn\DotNet8\InvestmentTechExcelAddIn"
NETWORK_PATH_6 = r"\\iamldnfs1\GDrive\Depts\Investment IT\Investment Solutions\Software\FrontOfficeExcelAddIn\InvestmentTechExcelAddIn"


class AddinManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NinetyOne Add-In Manager")
        self.root.geometry("520x720")  # Increased height to fit tabs and version row
        self.root.resizable(False, False)

        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')

        # Variables
        self.env_var = tk.StringVar()
        self.current_version_var = tk.StringVar(value="Not set")

        self.create_widgets()
        self.load_current_env()

        # Kick off asynchronous network checks for both paths
        self.fetch_versions_async()

    def create_widgets(self):
        # --- Environment Section ---
        env_frame = ttk.LabelFrame(self.root, text="Environment Settings", padding=(10, 10))
        env_frame.pack(fill="x", padx=10, pady=5)

        # ENV Row
        ttk.Label(env_frame, text="Current ENV:").grid(row=0, column=0, sticky="w", pady=5)
        self.env_dropdown = ttk.Combobox(
            env_frame,
            textvariable=self.env_var,
            values=["Prod", "UAT", "Dev", "local"],
            state="readonly",
            width=20
        )
        self.env_dropdown.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(env_frame, text="Update ENV", command=self.save_env).grid(row=0, column=2, padx=5)

        # VERSION Row
        ttk.Label(env_frame, text="Current VERSION:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Label(env_frame, textvariable=self.current_version_var, font=("TkDefaultFont", 9, "bold")).grid(row=1,
                                                                                                            column=1,
                                                                                                            sticky="w",
                                                                                                            padx=10,
                                                                                                            pady=5)

        # --- Network Versions Section (Tabs) ---
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
        ttk.Button(tab8, text="Open Network Path", command=lambda: self.open_in_explorer(NETWORK_PATH_8)).grid(row=0,
                                                                                                               column=2,
                                                                                                               padx=5)

        # Tab 2: .Net 6.0
        tab6 = ttk.Frame(self.notebook, padding=(10, 10))
        self.notebook.add(tab6, text=".Net 6.0")

        ttk.Label(tab6, text="Available:").grid(row=0, column=0, sticky="w", pady=5)
        self.version6_var = tk.StringVar(value="Checking network...")
        self.version6_dropdown = ttk.Combobox(tab6, textvariable=self.version6_var, state="readonly", width=22)
        self.version6_dropdown.grid(row=0, column=1, padx=10, pady=5)
        ttk.Button(tab6, text="Open Network Path", command=lambda: self.open_in_explorer(NETWORK_PATH_6)).grid(row=0,
                                                                                                               column=2,
                                                                                                               padx=5)

        # --- Paths & Explorer Section ---
        paths_frame = ttk.LabelFrame(self.root, text="Local Directories", padding=(10, 10))
        paths_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(paths_frame, text="Open Env File Folder",
                   command=lambda: self.open_in_explorer(os.path.dirname(ENV_FILE_PATH))).pack(fill="x", pady=2)
        ttk.Button(paths_frame, text="Open Loader Path",
                   command=lambda: self.open_in_explorer(LOADER_PATH)).pack(fill="x", pady=2)
        ttk.Button(paths_frame, text="Open Local Test Path",
                   command=lambda: self.open_in_explorer(LOCAL_TEST_PATH)).pack(fill="x", pady=2)

        # --- Registry Section ---
        reg_frame = ttk.LabelFrame(self.root, text="Registry Check (excel/options)", padding=(10, 10))
        reg_frame.pack(fill="both", expand=True, padx=10, pady=5)

        reg_btn_frame = ttk.Frame(reg_frame)
        reg_btn_frame.pack(fill="x", pady=(0, 5))

        ttk.Button(reg_btn_frame, text="Check for 'NinetyOne'", command=self.check_registry).pack(side="left", fill="x",
                                                                                                  expand=True,
                                                                                                  padx=(0, 2))
        ttk.Button(reg_btn_frame, text="Open Regedit Here", command=self.open_regedit).pack(side="right", fill="x",
                                                                                            expand=True, padx=(2, 0))

        self.reg_output = scrolledtext.ScrolledText(reg_frame, height=5, width=50, wrap=tk.WORD)
        self.reg_output.pack(fill="both", expand=True)

        # --- Actions Section ---
        action_frame = ttk.Frame(self.root, padding=(10, 10))
        action_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(action_frame, text="Launch Excel", command=self.launch_excel).pack(fill="x", ipady=5)

    def load_current_env(self):
        if not os.path.exists(ENV_FILE_PATH):
            return
        try:
            with open(ENV_FILE_PATH, 'r') as f:
                for line in f:
                    if line.startswith("ENV="):
                        self.env_var.set(line.strip().split("=")[1])
                    elif line.startswith("VERSION="):
                        self.current_version_var.set(line.strip().split("=")[1])
        except Exception as e:
            messagebox.showerror("Read Error", f"Could not read env file:\n{e}")

    def save_env(self):
        selected_env = self.env_var.get()
        if not selected_env:
            messagebox.showwarning("Warning", "Please select an environment first.")
            return

        lines, env_found = [], False
        if os.path.exists(ENV_FILE_PATH):
            with open(ENV_FILE_PATH, 'r') as f:
                lines = f.readlines()

        for i, line in enumerate(lines):
            if line.startswith("ENV="):
                lines[i] = f"ENV={selected_env}\n"
                env_found = True
                break

        if not env_found:
            lines.insert(0, f"ENV={selected_env}\n")

        try:
            os.makedirs(os.path.dirname(ENV_FILE_PATH), exist_ok=True)
            with open(ENV_FILE_PATH, 'w') as f:
                f.writelines(lines)
            messagebox.showinfo("Success", f"Environment successfully updated to: {selected_env}")
        except Exception as e:
            messagebox.showerror("Write Error", f"Failed to update env file:\n{e}")

    # --- Async Network Check Logic ---
    def fetch_versions_async(self):
        """Spawns background threads to check both network paths independently."""
        threading.Thread(target=self._check_network_path,
                         args=(NETWORK_PATH_8, self.version8_var, self.version8_dropdown), daemon=True).start()
        threading.Thread(target=self._check_network_path,
                         args=(NETWORK_PATH_6, self.version6_var, self.version6_dropdown), daemon=True).start()

    def _check_network_path(self, path, var, dropdown):
        """Runs in the background thread."""
        try:
            if os.path.exists(path):
                # Grab all folders in the directory
                versions = [d for d in os.listdir(path) if os.path.isdir(os.path.join(path, d))]
                versions.sort(reverse=True)
                self.root.after(0, self._update_versions_ui, versions, var, dropdown)
            else:
                self.root.after(0, self._update_versions_ui_error, "Path not found", var, dropdown)
        except Exception:
            self.root.after(0, self._update_versions_ui_error, "Offline / Access Denied", var, dropdown)

    def _update_versions_ui(self, versions, var, dropdown):
        if versions:
            dropdown['values'] = versions
            var.set(versions[0])
        else:
            dropdown['values'] = []
            var.set("No versions found")

    def _update_versions_ui_error(self, status_msg, var, dropdown):
        dropdown['values'] = []
        var.set(status_msg)

    # --- Utilities ---
    def open_in_explorer(self, path):
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showwarning("Not Found", f"The directory does not exist or is unreachable:\n{path}")

    def check_registry(self):
        self.reg_output.delete(1.0, tk.END)
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH)
            found_mentions, i = [], 0
            while True:
                try:
                    name, value, reg_type = winreg.EnumValue(key, i)
                    if isinstance(value, str) and "ninetyone" in value.lower():
                        found_mentions.append(f"{name}: {value}")
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)

            if found_mentions:
                self.reg_output.insert(tk.END, "Found 'NinetyOne' in the following registry values:\n\n")
                for mention in found_mentions:
                    self.reg_output.insert(tk.END, f"- {mention}\n")
            else:
                self.reg_output.insert(tk.END, "No mentions of 'NinetyOne' found in the registry key.")
        except FileNotFoundError:
            self.reg_output.insert(tk.END, f"Registry key not found:\n{REG_PATH}")
        except Exception as e:
            self.reg_output.insert(tk.END, f"Error accessing registry:\n{e}")

    def open_regedit(self):
        try:
            target_key = r"Computer\HKEY_CURRENT_USER\\" + REG_PATH
            applet_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                        r"Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", 0,
                                        winreg.KEY_SET_VALUE)
            winreg.SetValueEx(applet_key, "LastKey", 0, winreg.REG_SZ, target_key)
            winreg.CloseKey(applet_key)
            os.startfile("regedit.exe")
        except PermissionError:
            messagebox.showerror("Permission Error",
                                 "Unable to modify Regedit's LastKey. You may need to run this script as Administrator.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Registry Editor:\n{e}")

    def launch_excel(self):
        try:
            os.startfile("excel.exe")
        except Exception as e:
            messagebox.showerror("Error", f"Could not launch Excel. Is it installed?\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AddinManagerApp(root)
    root.mainloop()