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
NETWORK_VERSIONS_PATH = r"\\iamldnfs1\GDrive\Depts\Investment IT\Investment Solutions\Software\FrontOfficeExcelAddIn\DotNet8\InvestmentTechExcelAddIn"


class AddinManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NinetyOne Add-In Manager")
        self.root.geometry("500x650")  # Increased height for new section
        self.root.resizable(False, False)

        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')

        self.create_widgets()
        self.load_current_env()

        # Kick off the asynchronous network check
        self.fetch_versions_async()

    def create_widgets(self):
        # --- Environment Section ---
        env_frame = ttk.LabelFrame(self.root, text="Environment Settings", padding=(10, 10))
        env_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(env_frame, text="Current ENV:").grid(row=0, column=0, sticky="w", pady=5)

        self.env_var = tk.StringVar()
        self.env_dropdown = ttk.Combobox(
            env_frame,
            textvariable=self.env_var,
            values=["Prod", "UAT", "Dev", "local"],
            state="readonly",
            width=20
        )
        self.env_dropdown.grid(row=0, column=1, padx=10, pady=5)

        ttk.Button(env_frame, text="Update ENV", command=self.save_env).grid(row=0, column=2, padx=5)

        # --- Network Versions Section ---
        versions_frame = ttk.LabelFrame(self.root, text="Network Versions", padding=(10, 10))
        versions_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(versions_frame, text="Available:").grid(row=0, column=0, sticky="w", pady=5)

        self.version_var = tk.StringVar()
        self.version_var.set("Checking network...")
        self.version_dropdown = ttk.Combobox(
            versions_frame,
            textvariable=self.version_var,
            state="readonly",
            width=20
        )
        self.version_dropdown.grid(row=0, column=1, padx=10, pady=5)

        ttk.Button(versions_frame, text="Open Network Path",
                   command=lambda: self.open_in_explorer(NETWORK_VERSIONS_PATH)).grid(row=0, column=2, padx=5)

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

        self.reg_output = scrolledtext.ScrolledText(reg_frame, height=6, width=50, wrap=tk.WORD)
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
                        break
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
        """Spawns a background thread to check the network path."""
        # daemon=True ensures the thread dies immediately if you close the app
        threading.Thread(target=self._check_network_path, daemon=True).start()

    def _check_network_path(self):
        """Runs in the background thread."""
        try:
            if os.path.exists(NETWORK_VERSIONS_PATH):
                # Grab all folders in the directory
                versions = [d for d in os.listdir(NETWORK_VERSIONS_PATH)
                            if os.path.isdir(os.path.join(NETWORK_VERSIONS_PATH, d))]

                # Sort them (assuming semantic versioning or alphabetical)
                versions.sort(reverse=True)

                # Safely update the UI back on the main thread
                self.root.after(0, self._update_versions_ui, versions)
            else:
                self.root.after(0, self._update_versions_ui_error, "Path not found")
        except Exception:
            self.root.after(0, self._update_versions_ui_error, "Offline / Access Denied")

    def _update_versions_ui(self, versions):
        """Runs on the main thread to update the Combobox."""
        if versions:
            self.version_dropdown['values'] = versions
            self.version_var.set(versions[0])  # Default to newest/top version
        else:
            self.version_dropdown['values'] = []
            self.version_var.set("No versions found")

    def _update_versions_ui_error(self, status_msg):
        """Runs on the main thread to update the Combobox on failure."""
        self.version_dropdown['values'] = []
        self.version_var.set(status_msg)

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