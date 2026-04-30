import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from config import FOXL_LOADER_PATH
from core.env_manager import read_env, update_env_param
from core.zip_ops import extract_and_install_zip, find_xll_in_folder
from core.registry_ops import point_excel_to_addin
from core.system_ops import launch_excel, close_excel, kill_excel
from ui.log_viewer import LogViewer
import ui.theme as T


def _lbl(parent, text="", muted=False, bold=False, tiny=False, **kw):
    bg   = kw.pop("bg", T.BG_CARD)
    fg   = kw.pop("fg", T.TEXT_MUTED if tiny else T.TEXT_SEC if muted else T.TEXT_PRI)
    font = kw.pop("font", T.F_TINY if tiny else T.F_BOLD if bold else T.F_UI)
    return tk.Label(parent, text=text, bg=bg, fg=fg, font=font, **kw)


def _btn(parent, text, cmd, style="normal", **kw):
    _s = {
        "normal":  (T.BG_INPUT, T.TEXT_PRI, T.BORDER),
        "accent":  (T.ACCENT,   "#000000",  T.ACCENT_DIM),
        "danger":  ("#5a1f1f",  "#ff8080",  "#7a2f2f"),
        "ghost":   (T.BG_CARD,  T.TEXT_SEC, T.BG_INPUT),
    }
    bg, fg, abg = _s.get(style, _s["normal"])
    padx = kw.pop("padx", 10)
    pady = kw.pop("pady", 6)
    return tk.Button(parent, text=text, command=cmd,
                     bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
                     relief="flat", bd=0, padx=padx, pady=pady,
                     cursor="hand2", font=T.F_UI, **kw)


def _card(parent, title=""):
    f = tk.Frame(parent, bg=T.BG_CARD)
    if title:
        tk.Label(f, text=title, fg=T.TEXT_SEC, bg=T.BG_CARD,
                 font=T.F_SMALL).pack(anchor="w", padx=14, pady=(10, 2))
    return f


class UserWindow:
    def __init__(self, root: tk.Toplevel):
        self.root = root
        self.root.title("FOXL  AddIn Manager  —  User")
        self.root.geometry("460x600")
        self.root.resizable(False, False)
        T.apply(root)

        self.env_var         = tk.StringVar()
        self.current_ver_var = tk.StringVar(value="Not set")
        self.zip_path_var    = tk.StringVar()
        self.target_name_var = tk.StringVar(value=r"_91ExcelAddIn\DEV")
        self.folder_path_var = tk.StringVar()

        self._build()
        self._sync_env()

    def _build(self):
        self.root.configure(bg=T.BG_MAIN)

        # ── Topbar ────────────────────────────────────────────────────────────
        topbar = tk.Frame(self.root, bg=T.BG_DARK, height=48)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)
        logo_f = tk.Frame(topbar, bg=T.BG_DARK)
        logo_f.pack(side="left", padx=16, pady=10)
        tk.Label(logo_f, text="★", fg=T.ACCENT, bg=T.BG_DARK,
                 font=("Segoe UI", 12, "bold")).pack(side="left")
        name_f = tk.Frame(logo_f, bg=T.BG_DARK)
        name_f.pack(side="left", padx=(6, 0))
        tk.Label(name_f, text="FOXL", fg=T.TEXT_PRI, bg=T.BG_DARK,
                 font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(name_f, text="USER VIEW", fg=T.TEXT_SEC, bg=T.BG_DARK,
                 font=("Segoe UI", 6)).pack(anchor="w")

        # Excel quick buttons
        xl_f = tk.Frame(topbar, bg=T.BG_DARK)
        xl_f.pack(side="right", padx=12)
        for lbl, cmd, danger in [
            ("Open",  self.do_launch_excel, False),
            ("Close", self.do_close_excel,  False),
            ("Kill",  self.do_kill_excel,   True),
        ]:
            tk.Button(xl_f, text=lbl,
                      bg=T.BG_INPUT, fg=T.ERROR if danger else T.TEXT_PRI,
                      activebackground=T.BORDER, relief="flat", bd=0,
                      padx=10, pady=4, cursor="hand2", font=T.F_UI,
                      command=cmd).pack(side="left", padx=2)

        # ── Body ──────────────────────────────────────────────────────────────
        body = tk.Frame(self.root, bg=T.BG_MAIN)
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # Environment card
        env_c = _card(body, "Environment")
        env_c.pack(fill="x", pady=(0, 8))

        _lbl(env_c, "CURRENT ENV", tiny=True).pack(anchor="w", padx=14, pady=(4, 3))
        dd_row = tk.Frame(env_c, bg=T.BG_CARD)
        dd_row.pack(fill="x", padx=14, pady=(0, 8))
        self.env_dd = ttk.Combobox(dd_row, textvariable=self.env_var,
                                   values=["Prod", "UAT", "Dev", "local"],
                                   state="readonly", width=18)
        self.env_dd.pack(side="left", padx=(0, 10))
        _btn(dd_row, "⚡  Apply", self.save_env, style="accent").pack(side="left")

        _lbl(env_c, "CURRENT VERSION", tiny=True).pack(anchor="w", padx=14, pady=(0, 3))
        _lbl(env_c, bg=T.BG_CARD, textvariable=self.current_ver_var,
             bold=True).pack(anchor="w", padx=14, pady=(0, 10))

        # Install card
        inst_c = _card(body, "Install Local Build")
        inst_c.pack(fill="x", pady=(0, 8))

        # From ZIP
        _lbl(inst_c, "FROM ZIP", tiny=True).pack(anchor="w", padx=14, pady=(4, 3))

        zip_row = tk.Frame(inst_c, bg=T.BG_CARD)
        zip_row.pack(fill="x", padx=14, pady=(0, 4))
        zip_row.columnconfigure(0, weight=1)
        ttk.Entry(zip_row, textvariable=self.zip_path_var, state="readonly").grid(
            row=0, column=0, sticky="ew", padx=(0, 6))
        _btn(zip_row, "Browse…", self.browse_zip, pady=4).grid(row=0, column=1)

        name_row = tk.Frame(inst_c, bg=T.BG_CARD)
        name_row.pack(fill="x", padx=14, pady=(0, 6))
        name_row.columnconfigure(0, weight=1)
        _lbl(name_row, "Folder name:", muted=True, bg=T.BG_CARD).grid(
            row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(name_row, textvariable=self.target_name_var, width=24).grid(
            row=0, column=1, sticky="ew")

        _btn(inst_c, "Extract & Point Excel to Build", self.install_from_zip,
             style="accent").pack(fill="x", padx=14, pady=(0, 10))

        # Divider
        div = tk.Frame(inst_c, bg=T.BG_CARD)
        div.pack(fill="x", padx=14, pady=(0, 8))
        tk.Frame(div, bg=T.BORDER, height=1).pack(side="left", fill="x", expand=True)
        tk.Label(div, text="  or use an extracted folder  ",
                 fg=T.TEXT_MUTED, bg=T.BG_CARD, font=T.F_SMALL).pack(side="left")
        tk.Frame(div, bg=T.BORDER, height=1).pack(side="left", fill="x", expand=True)

        # From Folder
        fld_row = tk.Frame(inst_c, bg=T.BG_CARD)
        fld_row.pack(fill="x", padx=14, pady=(0, 6))
        fld_row.columnconfigure(0, weight=1)
        ttk.Entry(fld_row, textvariable=self.folder_path_var, state="readonly").grid(
            row=0, column=0, sticky="ew", padx=(0, 6))
        _btn(fld_row, "Browse…", self.browse_folder, pady=4).grid(row=0, column=1)

        _btn(inst_c, "Point Excel to Extracted Build", self.install_from_folder).pack(
            fill="x", padx=14, pady=(0, 8))

        # Revert
        tk.Frame(inst_c, bg=T.BORDER, height=1).pack(fill="x", padx=14)
        _btn(inst_c, "Revert to Standard FOXL Loader", self.revert_loader,
             style="ghost").pack(fill="x", padx=14, pady=(6, 12))

        # Support Logs button
        _btn(body, "⊕  View Support Logs", self.open_log_viewer).pack(
            fill="x", pady=(0, 4))

    # ── Actions ───────────────────────────────────────────────────────────────

    def _sync_env(self):
        e, v = read_env()
        self.env_var.set(e)
        self.current_ver_var.set(v)

    def save_env(self):
        val = self.env_var.get()
        ok, msg = update_env_param("ENV", val)
        if ok:
            messagebox.showinfo("Updated", f"Environment updated to: {val}")
        else:
            messagebox.showerror("Error", msg)

    def browse_zip(self):
        p = filedialog.askopenfilename(title="Select FOXL Build Zip",
                                       filetypes=[("Zip files", "*.zip"), ("All files", "*.*")])
        if p:
            self.zip_path_var.set(p)

    def browse_folder(self):
        p = filedialog.askdirectory(title="Select Extracted Build Folder")
        if p:
            self.folder_path_var.set(p)

    def install_from_zip(self):
        zip_path = self.zip_path_var.get()
        target   = self.target_name_var.get().strip()
        if not zip_path:
            messagebox.showwarning("Warning", "Select a zip file first.")
            return
        if not target:
            messagebox.showwarning("Warning", "Enter a folder name (e.g. DEV, v1.2).")
            return
        ok, msg, xll = extract_and_install_zip(zip_path, target)
        if not ok:
            messagebox.showerror("Extraction Error", msg)
            return
        ok2, msg2 = point_excel_to_addin(xll)
        if not ok2:
            messagebox.showerror("Registry Error", msg2)
            return
        messagebox.showinfo("Done", f"Build extracted and configured.\n\n{msg2}")

    def install_from_folder(self):
        folder = self.folder_path_var.get()
        if not folder:
            messagebox.showwarning("Warning", "Select a folder first.")
            return
        ok, result = find_xll_in_folder(folder)
        if not ok:
            messagebox.showerror("Error", result)
            return
        ok2, msg2 = point_excel_to_addin(result)
        if not ok2:
            messagebox.showerror("Registry Error", msg2)
            return
        messagebox.showinfo("Done", f"Excel pointed to build.\n\n{msg2}")

    def revert_loader(self):
        ok, msg = point_excel_to_addin(FOXL_LOADER_PATH)
        if ok:
            messagebox.showinfo("Reverted",
                                f"Pointed Excel back to standard loader:\n\n{FOXL_LOADER_PATH}")
        else:
            messagebox.showerror("Registry Error", msg)

    def open_log_viewer(self):
        LogViewer(self.root).grab_set()

    def do_launch_excel(self):
        ok, err = launch_excel()
        if not ok: messagebox.showerror("Error", err)

    def do_close_excel(self):
        ok, err = close_excel()
        if not ok: messagebox.showerror("Error", err)

    def do_kill_excel(self):
        if messagebox.askyesno("Force Kill", "Kill Excel? Unsaved work will be lost."):
            ok, err = kill_excel()
            if not ok: messagebox.showerror("Error", err)
