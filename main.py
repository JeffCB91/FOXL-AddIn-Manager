import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import winreg

# --- Configuration Paths ---
ENV_FILE_PATH = os.path.expandvars(r"%APPDATA%\NinetyOne - FrontOfficeExcelAddIn\front-office-excel-addin-env")
LOADER_PATH = r"C:\Program Files\Microsoft Office\root\Office16\Library\InvestmentTechExcelAddin"
LOCAL_TEST_PATH = r"C:\ExcelAddIn\_91ExcelAddIn"
REG_PATH = r"Software\Microsoft\Office\16.0\excel\options"


class AddinManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NinetyOne Add-In Manager")
        self.root.geometry("500x560")  # Increased height to accommodate new buttons
        self.root.resizable(False, False)

        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')

        self.create_widgets()
        self.load_current_env()

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

        # --- Paths & Explorer Section ---
        paths_frame = ttk.LabelFrame(self.root, text="Directories", padding=(10, 10))
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

        # Container for registry buttons to keep them tidy
        reg_btn_frame = ttk.Frame(reg_frame)
        reg_btn_frame.pack(fill="x", pady=(0, 5))

        ttk.Button(reg_btn_frame, text="Check for 'NinetyOne'", command=self.check_registry).pack(side="left", fill="x",
                                                                                                  expand=True,
                                                                                                  padx=(0, 2))
        ttk.Button(reg_btn_frame, text="Open Regedit Here", command=self.open_regedit).pack(side="right", fill="x",
                                                                                            expand=True, padx=(2, 0))

        self.reg_output = scrolledtext.ScrolledText(reg_frame, height=8, width=50, wrap=tk.WORD)
        self.reg_output.pack(fill="both", expand=True)

        # --- Actions Section ---
        action_frame = ttk.Frame(self.root, padding=(10, 10))
        action_frame.pack(fill="x", padx=10, pady=5)

        ttk.Button(action_frame, text="Launch Excel", command=self.launch_excel).pack(fill="x", ipady=5)

    def load_current_env(self):
        """Reads the env file and updates the dropdown."""
        if not os.path.exists(ENV_FILE_PATH):
            return

        try:
            with open(ENV_FILE_PATH, 'r') as f:
                for line in f:
                    if line.startswith("ENV="):
                        current_env = line.strip().split("=")[1]
                        self.env_var.set(current_env)
                        break
        except Exception as e:
            messagebox.showerror("Read Error", f"Could not read env file:\n{e}")

    def save_env(self):
        """Updates the ENV= line in the file, preserving other lines like VERSION."""
        selected_env = self.env_var.get()
        if not selected_env:
            messagebox.showwarning("Warning", "Please select an environment first.")
            return

        lines = []
        env_found = False

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

    def open_in_explorer(self, path):
        """Opens the given path in Windows Explorer."""
        if os.path.exists(path):
            os.startfile(path)
        else:
            messagebox.showwarning("Not Found", f"The directory does not exist yet:\n{path}")

    def check_registry(self):
        """Scans HKCU\\Software\\Microsoft\\Office\\16.0\\excel\\options for 'NinetyOne'."""
        self.reg_output.delete(1.0, tk.END)
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH)

            found_mentions = []
            i = 0
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
        """Opens Registry Editor at the specified Excel options path."""
        try:
            # Format the path exactly how Regedit expects it in the LastKey memory
            target_key = r"Computer\HKEY_CURRENT_USER\\" + REG_PATH

            # Temporarily modify Regedit's LastKey value
            applet_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                        r"Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", 0,
                                        winreg.KEY_SET_VALUE)
            winreg.SetValueEx(applet_key, "LastKey", 0, winreg.REG_SZ, target_key)
            winreg.CloseKey(applet_key)

            # Launch Regedit (it will open to the LastKey location)
            os.startfile("regedit.exe")
        except PermissionError:
            messagebox.showerror("Permission Error",
                                 "Unable to modify Regedit's LastKey. You may need to run this script as Administrator.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Registry Editor:\n{e}")

    def launch_excel(self):
        """Launches the Excel application."""
        try:
            os.startfile("excel.exe")
        except Exception as e:
            messagebox.showerror("Error", f"Could not launch Excel. Is it installed?\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = AddinManagerApp(root)
    root.mainloop()