import tkinter as tk
from tkinter import ttk, messagebox
from core.env_manager import read_env, update_env_param

class UserWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("FOXL Add-In")
        self.root.geometry("320x160")
        self.root.resizable(False, False)

        ttk.Style().theme_use('clam')

        self.env_var = tk.StringVar()
        self.current_version_var = tk.StringVar(value="Not set")

        self.create_widgets()
        self.sync_ui_with_env()

    def create_widgets(self):
        # Main padding frame
        frame = ttk.Frame(self.root, padding=(20, 20))
        frame.pack(fill="both", expand=True)

        # ENV Row
        ttk.Label(frame, text="Current ENV:").grid(row=0, column=0, sticky="w", pady=5)
        self.env_dd = ttk.Combobox(frame, textvariable=self.env_var, values=["Prod", "UAT", "Dev", "local"], state="readonly", width=15)
        self.env_dd.grid(row=0, column=1, padx=10, pady=5)

        # VERSION Row
        ttk.Label(frame, text="Current VERSION:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Label(frame, textvariable=self.current_version_var, font=("TkDefaultFont", 9, "bold")).grid(row=1, column=1, sticky="w", padx=10, pady=5)

        # Action Button
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="Update Environment", command=self.save_env, width=30).pack()

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