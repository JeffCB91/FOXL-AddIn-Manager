import tkinter as tk
from tkinter import ttk
import os
import json
from config import NETWORK_PATH_8, NETWORK_PATH_6, LOCAL_TEST_PATH, CONFIG_FILE_NAME


class ConfigViewer(tk.Toplevel):
    def __init__(self, parent, current_env, net8_versions, net6_versions, local_versions):
        super().__init__(parent)
        self.title("Config Variables Viewer")
        self.geometry("750x480")

        self.current_env = current_env
        self.versions = {"8": net8_versions, "6": net6_versions, "local": local_versions}

        self.create_widgets()

    def create_widgets(self):
        ctrl = ttk.Frame(self, padding=(10, 10))
        ctrl.pack(fill="x")

        ttk.Label(ctrl, text="Source:").grid(row=0, column=0, padx=2)
        self.src_var = tk.StringVar(value="Network .Net 8")
        self.src_cb = ttk.Combobox(ctrl, textvariable=self.src_var,
                                   values=["Network .Net 8", "Network .Net 6", "Local Test Path"], state="readonly",
                                   width=15)
        self.src_cb.grid(row=0, column=1, padx=5)

        ttk.Label(ctrl, text="Version:").grid(row=0, column=2, padx=2)
        self.ver_var = tk.StringVar()
        self.ver_cb = ttk.Combobox(ctrl, textvariable=self.ver_var, state="readonly", width=15)
        self.ver_cb.grid(row=0, column=3, padx=5)

        ttk.Label(ctrl, text="Env:").grid(row=0, column=4, padx=2)
        self.env_cb_var = tk.StringVar()
        self.env_cb = ttk.Combobox(ctrl, textvariable=self.env_cb_var, state="readonly", width=10)
        self.env_cb.grid(row=0, column=5, padx=5)

        tree_f = ttk.Frame(self, padding=(10, 0, 10, 10))
        tree_f.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_f, columns=("P", "V"), show="headings")
        self.tree.heading("P", text="Parameter")
        self.tree.heading("V", text="Value")
        self.tree.column("P", width=200)
        self.tree.column("V", width=500)

        sb = ttk.Scrollbar(tree_f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        self.tmp_cfg = []

        self.src_cb.bind("<<ComboboxSelected>>", self.on_src)
        self.ver_cb.bind("<<ComboboxSelected>>", self.on_ver)
        self.env_cb.bind("<<ComboboxSelected>>", self.on_env)

        self.on_src()  # Trigger initial load

    def on_src(self, event=None):
        s = self.src_var.get()
        v = self.versions["8"] if "8" in s else (self.versions["6"] if "6" in s else self.versions["local"])
        self.ver_cb['values'] = v
        if v:
            self.ver_cb.set(v[0])
            self.on_ver()
        else:
            self.ver_cb.set("N/A")
            self.tree.delete(*self.tree.get_children())

    def on_ver(self, event=None):
        s, vr = self.src_var.get(), self.ver_var.get()
        if vr == "N/A": return
        base = NETWORK_PATH_8 if "8" in s else (NETWORK_PATH_6 if "6" in s else LOCAL_TEST_PATH)
        self.load_json(os.path.join(base, vr, CONFIG_FILE_NAME))

    def load_json(self, path):
        self.tree.delete(*self.tree.get_children())
        self.env_cb['values'] = []
        self.env_cb.set("")
        self.tmp_cfg = []

        if not os.path.exists(path):
            self.tree.insert("", "end", values=("File Not Found", path))
            return

        try:
            with open(path, 'r', encoding='utf-8') as f:
                self.tmp_cfg = json.load(f)
            en = [i.get("environment") for i in self.tmp_cfg if isinstance(i, dict)]
            self.env_cb['values'] = en
            if en:
                self.env_cb.set(self.current_env if self.current_env in en else en[0])
                self.on_env()
        except Exception as e:
            self.tree.insert("", "end", values=("JSON Error", str(e)))

    def on_env(self, event=None):
        self.tree.delete(*self.tree.get_children())
        for i in self.tmp_cfg:
            if i.get("environment") == self.env_cb_var.get():
                for k, v in i.get("parameters", {}).items():
                    self.tree.insert("", "end", values=(k, v))
                break
