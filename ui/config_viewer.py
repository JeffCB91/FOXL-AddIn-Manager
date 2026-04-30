import tkinter as tk
from tkinter import ttk
import os
import json

from config import NETWORK_PATH_8, NETWORK_PATH_6, LOCAL_TEST_PATH, CONFIG_FILE_NAME
import ui.theme as T


class ConfigViewer(tk.Toplevel):
    def __init__(self, parent, current_env, net8_versions, net6_versions, local_versions):
        super().__init__(parent)
        self.title("FOXL  Config Variables Viewer")
        self.geometry("820x520")
        T.apply(self)

        self.current_env = current_env
        self.versions    = {"8": net8_versions, "6": net6_versions, "local": local_versions}
        self._data: list = []

        self._build()

    def _build(self):
        self.configure(bg=T.BG_MAIN)

        # Topbar
        topbar = tk.Frame(self, bg=T.BG_DARK, height=48)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)
        logo_f = tk.Frame(topbar, bg=T.BG_DARK)
        logo_f.pack(side="left", padx=16, pady=10)
        tk.Label(logo_f, text="★", fg=T.ACCENT, bg=T.BG_DARK,
                 font=("Segoe UI", 12, "bold")).pack(side="left")
        tk.Label(logo_f, text="  Config Variables Viewer", fg=T.TEXT_PRI, bg=T.BG_DARK,
                 font=("Segoe UI", 10, "bold")).pack(side="left")

        # Controls card
        ctrl_c = tk.Frame(self, bg=T.BG_CARD)
        ctrl_c.pack(fill="x", padx=16, pady=(12, 0))

        ctrl_inner = tk.Frame(ctrl_c, bg=T.BG_CARD)
        ctrl_inner.pack(fill="x", padx=14, pady=10)

        # Source
        src_f = tk.Frame(ctrl_inner, bg=T.BG_CARD)
        src_f.pack(side="left", padx=(0, 20))
        tk.Label(src_f, text="SOURCE", fg=T.TEXT_MUTED, bg=T.BG_CARD,
                 font=T.F_TINY).pack(anchor="w")
        self.src_var = tk.StringVar(value="Network .Net 8")
        self.src_cb  = ttk.Combobox(src_f, textvariable=self.src_var,
                                    values=["Network .Net 8", "Network .Net 6", "Local Test Path"],
                                    state="readonly", width=18)
        self.src_cb.pack(pady=(3, 0))
        self.src_cb.bind("<<ComboboxSelected>>", self.on_src)

        # Version
        ver_f = tk.Frame(ctrl_inner, bg=T.BG_CARD)
        ver_f.pack(side="left", padx=(0, 20))
        tk.Label(ver_f, text="VERSION", fg=T.TEXT_MUTED, bg=T.BG_CARD,
                 font=T.F_TINY).pack(anchor="w")
        self.ver_var = tk.StringVar()
        self.ver_cb  = ttk.Combobox(ver_f, textvariable=self.ver_var,
                                    state="readonly", width=14)
        self.ver_cb.pack(pady=(3, 0))
        self.ver_cb.bind("<<ComboboxSelected>>", self.on_ver)

        # Environment
        env_f = tk.Frame(ctrl_inner, bg=T.BG_CARD)
        env_f.pack(side="left", padx=(0, 20))
        tk.Label(env_f, text="ENVIRONMENT", fg=T.TEXT_MUTED, bg=T.BG_CARD,
                 font=T.F_TINY).pack(anchor="w")
        self.env_var = tk.StringVar()
        self.env_cb  = ttk.Combobox(env_f, textvariable=self.env_var,
                                    state="readonly", width=12)
        self.env_cb.pack(pady=(3, 0))
        self.env_cb.bind("<<ComboboxSelected>>", self.on_env)

        # Table card
        tbl_c = tk.Frame(self, bg=T.BG_CARD)
        tbl_c.pack(fill="both", expand=True, padx=16, pady=(6, 16))
        tbl_c.rowconfigure(1, weight=1)
        tbl_c.columnconfigure(0, weight=1)

        self._title_lbl = tk.Label(tbl_c, text="", fg=T.TEXT_PRI, bg=T.BG_CARD,
                                   font=T.F_BOLD)
        self._title_lbl.pack(anchor="w", padx=14, pady=(10, 4))

        tree_f = tk.Frame(tbl_c, bg=T.BG_CARD)
        tree_f.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        tree_f.rowconfigure(0, weight=1)
        tree_f.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_f, columns=("P", "V"), show="headings")
        self.tree.heading("P", text="PARAMETER")
        self.tree.heading("V", text="VALUE")
        self.tree.column("P", width=220, minwidth=140)
        self.tree.column("V", width=500, stretch=True)
        sb = ttk.Scrollbar(tree_f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        self.on_src()

    def on_src(self, _=None):
        s = self.src_var.get()
        v = (self.versions["8"] if "8" in s else
             self.versions["6"] if "6" in s else
             self.versions["local"])
        self.ver_cb["values"] = v
        if v:
            self.ver_cb.set(v[0])
            self.on_ver()
        else:
            self.ver_cb.set("N/A")
            self.tree.delete(*self.tree.get_children())

    def on_ver(self, _=None):
        s, v = self.src_var.get(), self.ver_var.get()
        if v == "N/A": return
        base = NETWORK_PATH_8 if "8" in s else (NETWORK_PATH_6 if "6" in s else LOCAL_TEST_PATH)
        self._load(os.path.join(base, v, CONFIG_FILE_NAME))
        self._title_lbl.config(text=f"{v}  ·  {self.env_var.get()}")

    def _load(self, path):
        self.tree.delete(*self.tree.get_children())
        self.env_cb["values"] = []
        self.env_cb.set("")
        self._data = []
        if not os.path.exists(path):
            self.tree.insert("", "end", values=("File not found", path))
            return
        try:
            with open(path, encoding="utf-8") as f:
                self._data = json.load(f)
            envs = [i.get("environment") for i in self._data if isinstance(i, dict)]
            self.env_cb["values"] = envs
            self.env_cb.set(self.current_env if self.current_env in envs
                            else (envs[0] if envs else ""))
            self.on_env()
        except Exception as e:
            self.tree.insert("", "end", values=("JSON Error", str(e)))

    def on_env(self, _=None):
        self.tree.delete(*self.tree.get_children())
        env = self.env_var.get()
        for item in self._data:
            if item.get("environment") == env:
                for k, v in item.get("parameters", {}).items():
                    self.tree.insert("", "end", values=(k, v))
                break
        if self.ver_var.get():
            self._title_lbl.config(text=f"{self.ver_var.get()}  ·  {env}")
