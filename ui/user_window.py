import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from config import FOXL_LOADER_PATH  # <-- Import the new path
from core.env_manager import read_env, update_env_param
from core.zip_ops import extract_and_install_zip
from core.registry_ops import point_excel_to_addin
from core.system_ops import launch_excel, close_excel, kill_excel


class UserWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("FOXL Add-In Configurator")
        self.root.geometry("400x470")  # <-- Slightly taller to fit the Revert button
        self.root.resizable(False, False)

        ttk.Style().theme_use('clam')

        self.env_var = tk.StringVar()
        self.current_version_var = tk.StringVar(value="Not set")

        self.zip_path_var = tk.StringVar()
        self.target_name_var = tk.StringVar(value=r"_91ExcelAddIn\DEV")

        self.create_widgets()
        self.sync_ui_with_env()

    def create_widgets(self):
        # --- Environment Section ---
        env_frame = ttk.LabelFrame(self.root, text="Environment Selection", padding=(10, 10))
        env_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(env_frame, text="Current ENV:").grid(row=0, column=0, sticky="w", pady=5)
        self.env_dd = ttk.Combobox(env_frame, textvariable=self.env_var, values=["Prod", "UAT", "Dev", "local"],
                                   state="readonly", width=15)
        self.env_dd.grid(row=0, column=1, padx=10, pady=5)

        ttk.Label(env_frame, text="Current VERSION:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Label(env_frame, textvariable=self.current_version_var, font=("TkDefaultFont", 9, "bold")).grid(row=1,
                                                                                                            column=1,
                                                                                                            sticky="w",
                                                                                                            padx=10,
                                                                                                            pady=5)

        ttk.Button(env_frame, text="Update Environment", command=self.save_env).grid(row=2, column=0, columnspan=2,
                                                                                     pady=(10, 0), sticky="ew")

        # --- Local Zip Installer Section ---
        install_frame = ttk.LabelFrame(self.root, text="Install Local Build (.zip)", padding=(10, 10))
        install_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(install_frame, text="Zip File:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(install_frame, textvariable=self.zip_path_var, width=25, state="readonly").grid(row=0, column=1,
                                                                                                  padx=5, pady=5)
        ttk.Button(install_frame, text="Browse...", command=self.browse_zip, width=8).grid(row=0, column=2, padx=5,
                                                                                           pady=5)

        ttk.Label(install_frame, text="Folder Name:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(install_frame, textvariable=self.target_name_var, width=25).grid(row=1, column=1, padx=5, pady=5,
                                                                                   sticky="w")

        ttk.Button(install_frame, text="Extract & Point Excel to Build", command=self.install_build).grid(row=2,
                                                                                                          column=0,
                                                                                                          columnspan=3,
                                                                                                          pady=(10, 5),
                                                                                                          sticky="ew")

        ttk.Button(install_frame, text="Revert to Standard FOXL Loader", command=self.revert_loader).grid(row=3,
                                                                                                          column=0,
                                                                                                          columnspan=3,
                                                                                                          pady=(0, 0),
                                                                                                          sticky="ew")

        # --- Excel Controls Section ---
        excel_frame = ttk.LabelFrame(self.root, text="Excel Controls", padding=(10, 10))
        excel_frame.pack(fill="x", padx=10, pady=10)

        btn_container = ttk.Frame(excel_frame)
        btn_container.pack(fill="x")

        ttk.Button(btn_container, text="Open Excel", command=self.do_launch_excel).pack(side="left", expand=True,
                                                                                        fill="x", padx=(0, 2))
        ttk.Button(btn_container, text="Close Safely", command=self.do_close_excel).pack(side="left", expand=True,
                                                                                         fill="x", padx=2)
        ttk.Button(btn_container, text="Force Kill", command=self.do_kill_excel).pack(side="right", expand=True,
                                                                                      fill="x", padx=(2, 0))

    # --- Actions ---
    def sync_ui_with_env(self):
        e, v = read_env()
        self.env_var.set(e)
        self.current_version_var.set(v)

    def save_env(self):
        val = self.env_var.get()
        success, msg = update_env_param("ENV", val)
        if success:
            messagebox.showinfo("Success", f"Environment updated to: {val}")
        else:
            messagebox.showerror("Error", msg)

    def browse_zip(self):
        filepath = filedialog.askopenfilename(title="Select FOXL Build Zip",
                                              filetypes=[("Zip files", "*.zip"), ("All files", "*.*")])
        if filepath: self.zip_path_var.set(filepath)

    def install_build(self):
        zip_path = self.zip_path_var.get()
        target_name = self.target_name_var.get().strip()

        if not zip_path:
            messagebox.showwarning("Warning", "Please select a zip file first.")
            return
        if not target_name:
            messagebox.showwarning("Warning", "Please provide a folder name (e.g., DEV, UAT, v1.2).")
            return

        ext_success, ext_msg, final_xll_path = extract_and_install_zip(zip_path, target_name)
        if not ext_success:
            messagebox.showerror("Extraction Error", ext_msg)
            return

        reg_success, reg_msg = point_excel_to_addin(final_xll_path)
        if not reg_success:
            messagebox.showerror("Registry Error", reg_msg)
            return

        messagebox.showinfo("Success", f"Build extracted and configured successfully!\n\n{reg_msg}")

    # --- Revert Function ---
    def revert_loader(self):
        reg_success, reg_msg = point_excel_to_addin(FOXL_LOADER_PATH)
        if reg_success:
            messagebox.showinfo("Reverted",
                                f"Successfully pointed Excel back to the standard loader:\n\n{FOXL_LOADER_PATH}")
        else:
            messagebox.showerror("Registry Error", reg_msg)

    # --- Excel Operations ---
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
