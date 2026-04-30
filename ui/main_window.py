import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import sys
import json

from config import (
    ENV_FILE_PATH, LOADER_PATH, ADD_IN_PATH, LOCAL_TEST_PATH,
    BASE_LOCAL_PATH, NETWORK_PATH_8, NETWORK_PATH_6, TEMPLATES_PATH,
    LOG_DIR_PATH, FOXL_LOADER_PATH, PAT_FILE_PATH, ADO_PIPELINE_ID,
)
from core.env_manager import read_env, update_env_param
from core.scanner import scan_path_sync
from core.registry_ops import scan_registry_for_ninetyone, open_regedit_at_path
from core.system_ops import open_in_explorer, launch_excel, close_excel, kill_excel
from core.log_ops import (
    get_log_files, get_services, extract_service,
    generate_unified_timeline, export_logs_to_zip,
)
from core.ado_client import fetch_builds, download_artifact_zip
from core.deploy_ops import (
    get_next_version, get_existing_versions,
    deploy_zip_to_network, rollback_to_network,
)
import ui.theme as T

_PAT_KEY  = "ADO_PAT"
_BASE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PAT_GUIDE = os.path.join(_BASE_DIR, "PAT_GUIDE.md")


def _read_pat() -> str:
    try:
        with open(PAT_FILE_PATH) as f:
            for line in f:
                if line.strip().startswith(f"{_PAT_KEY}="):
                    return line.strip()[len(_PAT_KEY) + 1:]
    except OSError:
        pass
    return ""


def _write_pat(pat: str) -> None:
    os.makedirs(os.path.dirname(PAT_FILE_PATH), exist_ok=True)
    lines, replaced = [], False
    try:
        with open(PAT_FILE_PATH) as f:
            for line in f:
                if line.strip().startswith(f"{_PAT_KEY}="):
                    lines.append(f"{_PAT_KEY}={pat}\n"); replaced = True
                else:
                    lines.append(line)
    except OSError:
        pass
    if not replaced:
        lines.append(f"{_PAT_KEY}={pat}\n")
    with open(PAT_FILE_PATH, "w") as f:
        f.writelines(lines)


# ── Tiny widget helpers ────────────────────────────────────────────────────────

def _label(parent, text="", muted=False, bold=False, mono=False, tiny=False, accent=False, **kw):
    bg = kw.pop("bg", T.BG_CARD)
    fg = kw.pop("fg",
                T.ACCENT if accent else
                T.TEXT_MUTED if tiny else
                T.TEXT_SEC if muted else
                T.TEXT_PRI)
    font = kw.pop("font",
                  T.F_MONO if mono else
                  T.F_TINY if tiny else
                  T.F_BOLD if bold else
                  T.F_UI)
    return tk.Label(parent, text=text, bg=bg, fg=fg, font=font, **kw)


def _btn(parent, text, cmd, style="normal", **kw):
    _styles = {
        "normal": (T.BG_INPUT,   T.TEXT_PRI,  T.BORDER),
        "accent": (T.ACCENT,     "#000000",   T.ACCENT_DIM),
        "danger": ("#5a1f1f",    "#ff8080",   "#7a2f2f"),
        "ghost":  (T.BG_MAIN,    T.TEXT_SEC,  T.BG_CARD),
        "ghost_card": (T.BG_CARD, T.TEXT_SEC, T.BG_INPUT),
    }
    bg, fg, abg = _styles.get(style, _styles["normal"])
    padx = kw.pop("padx", 10)
    pady = kw.pop("pady", 6)
    return tk.Button(parent, text=text, command=cmd,
                     bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
                     relief="flat", bd=0, padx=padx, pady=pady,
                     cursor="hand2", font=T.F_UI, **kw)


def _sep(parent, bg=T.BG_MAIN):
    tk.Frame(parent, bg=T.BORDER, height=1).pack(fill="x")


class MainWindow:
    _NAV = [
        ("Dashboard",    "⊞"),
        ("Deploy",       "↑"),
        ("Configs",      "≡"),
        ("Support Logs", "⊕"),
        ("Registry",     "⋯"),
    ]
    _SECTIONS = {
        "Dashboard":    "OVERVIEW",
        "Deploy":       "ADMIN",
        "Configs":      "ADMIN",
        "Support Logs": "DIAGNOSTICS",
        "Registry":     "ADMIN",
    }

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("FOXL  AddIn Manager")
        self.root.geometry("1200x820")
        self.root.minsize(980, 680)
        T.apply(root)

        # Shared state
        self._env_var    = tk.StringVar()
        self._ver_var    = tk.StringVar(value="—")
        self._versions   = {"v8": [], "v6": [], "local": []}
        self._panels: dict[str, tk.Frame] = {}
        self._nav_btns: dict[str, tk.Button] = {}

        self._build_shell()
        self._build_all_panels()
        self._nav_to("Dashboard")
        self._sync_env()
        self._start_scans()

    # ── Shell ──────────────────────────────────────────────────────────────────

    def _build_shell(self):
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Sidebar
        sidebar = tk.Frame(self.root, bg=T.BG_DARK, width=208)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        self._build_sidebar(sidebar)

        # Main column: topbar + content
        main_col = tk.Frame(self.root, bg=T.BG_MAIN)
        main_col.grid(row=0, column=1, sticky="nsew")
        main_col.columnconfigure(0, weight=1)
        main_col.rowconfigure(1, weight=1)

        self._build_topbar(main_col)

        self._content = tk.Frame(main_col, bg=T.BG_MAIN)
        self._content.grid(row=1, column=0, sticky="nsew")
        self._content.columnconfigure(0, weight=1)
        self._content.rowconfigure(0, weight=1)

    def _build_sidebar(self, parent):
        # Logo
        logo = tk.Frame(parent, bg=T.BG_DARK)
        logo.pack(fill="x", padx=16, pady=(20, 20))
        tk.Label(logo, text="★", fg=T.ACCENT, bg=T.BG_DARK, font=("Segoe UI", 16, "bold")).pack(side="left")
        name_f = tk.Frame(logo, bg=T.BG_DARK)
        name_f.pack(side="left", padx=(8, 0))
        tk.Label(name_f, text="FOXL", fg=T.TEXT_PRI, bg=T.BG_DARK,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(name_f, text="ADDIN MANAGER", fg=T.TEXT_SEC, bg=T.BG_DARK,
                 font=("Segoe UI", 6)).pack(anchor="w")

        # Section heading
        tk.Label(parent, text="ADMIN WORKSPACE", fg=T.TEXT_MUTED, bg=T.BG_DARK,
                 font=T.F_TINY).pack(fill="x", padx=16, pady=(0, 4))

        # Nav buttons
        for name, icon in self._NAV:
            btn = tk.Button(
                parent,
                text=f"  {icon}   {name}",
                anchor="w",
                bg=T.BG_DARK, fg=T.TEXT_SEC,
                activebackground=T.BG_CARD, activeforeground=T.TEXT_PRI,
                relief="flat", bd=0, cursor="hand2",
                font=T.F_UI, padx=16, pady=9,
                command=lambda n=name: self._nav_to(n),
            )
            btn.pack(fill="x")
            self._nav_btns[name] = btn

        # Status bar at bottom of sidebar
        sb_f = tk.Frame(parent, bg=T.BG_DARK)
        sb_f.pack(side="bottom", fill="x", padx=16, pady=12)
        tk.Frame(parent, bg=T.BORDER, height=1).pack(side="bottom", fill="x")

        row1 = tk.Frame(sb_f, bg=T.BG_DARK)
        row1.pack(fill="x", pady=(0, 2))
        tk.Label(row1, text="ENV", fg=T.TEXT_MUTED, bg=T.BG_DARK, font=T.F_TINY, width=8, anchor="w").pack(side="left")
        self._sb_env = tk.Label(row1, textvariable=self._env_var, fg=T.ACCENT, bg=T.BG_DARK, font=T.F_SMALL)
        self._sb_env.pack(side="left")

        row2 = tk.Frame(sb_f, bg=T.BG_DARK)
        row2.pack(fill="x")
        tk.Label(row2, text="VERSION", fg=T.TEXT_MUTED, bg=T.BG_DARK, font=T.F_TINY, width=8, anchor="w").pack(side="left")
        tk.Label(row2, textvariable=self._ver_var, fg=T.TEXT_SEC, bg=T.BG_DARK, font=T.F_SMALL).pack(side="left")

    def _build_topbar(self, parent):
        bar = tk.Frame(parent, bg=T.BG_DARK, height=52)
        bar.grid(row=0, column=0, sticky="ew")
        bar.grid_propagate(False)

        # Breadcrumb
        bc = tk.Frame(bar, bg=T.BG_DARK)
        bc.pack(side="left", padx=(24, 0), pady=12)
        self._bc_section = tk.Label(bc, text="OVERVIEW", fg=T.TEXT_MUTED, bg=T.BG_DARK, font=T.F_TINY)
        self._bc_section.pack(anchor="w")
        self._bc_title = tk.Label(bc, text="Dashboard", fg=T.TEXT_PRI, bg=T.BG_DARK,
                                  font=("Segoe UI", 11, "bold"))
        self._bc_title.pack(anchor="w")

        # Right controls
        right = tk.Frame(bar, bg=T.BG_DARK)
        right.pack(side="right", padx=16, pady=10)

        # Admin/User toggle
        self._user_btn = tk.Button(right, text="User", bg=T.BG_INPUT, fg=T.TEXT_SEC,
                                   font=T.F_UI, relief="flat", bd=0, padx=12, pady=4,
                                   cursor="hand2", command=self._open_user_view)
        self._user_btn.pack(side="right", padx=(1, 0))
        tk.Button(right, text="Admin", bg=T.ACCENT, fg="#000000",
                  font=T.F_BOLD, relief="flat", bd=0, padx=12, pady=4).pack(side="right", padx=1)

        # Divider
        tk.Frame(right, bg=T.BORDER, width=1).pack(side="right", fill="y", padx=10)

        # Excel buttons
        kill_btn = tk.Button(right, text="Kill", bg=T.BG_INPUT, fg=T.ERROR,
                             activebackground=T.BORDER, activeforeground=T.ERROR,
                             relief="flat", bd=0, padx=10, pady=4,
                             cursor="hand2", font=T.F_UI, command=self._xl_kill)
        kill_btn.pack(side="right", padx=2)
        for lbl, cmd in [("Close", self._xl_close), ("Open", self._xl_open)]:
            tk.Button(right, text=lbl, bg=T.BG_INPUT, fg=T.TEXT_PRI,
                      activebackground=T.BORDER, activeforeground=T.TEXT_PRI,
                      relief="flat", bd=0, padx=10, pady=4,
                      cursor="hand2", font=T.F_UI, command=cmd).pack(side="right", padx=2)

        # Excel status dot
        xl_f = tk.Frame(right, bg=T.BG_DARK)
        xl_f.pack(side="right", padx=(0, 8))
        tk.Label(xl_f, text="●", fg=T.TEXT_MUTED, bg=T.BG_DARK, font=("Segoe UI", 8)).pack(side="left")
        tk.Label(xl_f, text=" Excel not running", fg=T.TEXT_SEC, bg=T.BG_DARK, font=T.F_SMALL).pack(side="left")

    def _nav_to(self, name: str):
        for n, btn in self._nav_btns.items():
            if n == name:
                btn.config(bg=T.BG_CARD, fg=T.ACCENT, font=T.F_BOLD)
            else:
                btn.config(bg=T.BG_DARK, fg=T.TEXT_SEC, font=T.F_UI)
        self._bc_section.config(text=self._SECTIONS.get(name, ""))
        self._bc_title.config(text=name)
        for n, panel in self._panels.items():
            if n == name:
                panel.grid(row=0, column=0, sticky="nsew")
            else:
                panel.grid_remove()
        if name == "Support Logs" and not self._logs_files:
            self._logs_load()

    # ── Panel factory ──────────────────────────────────────────────────────────

    def _build_all_panels(self):
        for name, _ in self._NAV:
            f = tk.Frame(self._content, bg=T.BG_MAIN)
            self._panels[name] = f
        self._build_dashboard(self._panels["Dashboard"])
        self._build_deploy(self._panels["Deploy"])
        self._build_configs(self._panels["Configs"])
        self._build_logs(self._panels["Support Logs"])
        self._build_registry(self._panels["Registry"])

    # ── Layout helpers ─────────────────────────────────────────────────────────

    def _card(self, parent: tk.Widget, title: str = "") -> tk.Frame:
        f = tk.Frame(parent, bg=T.BG_CARD)
        if title:
            tk.Label(f, text=title, fg=T.TEXT_SEC, bg=T.BG_CARD,
                     font=T.F_SMALL).pack(anchor="w", padx=14, pady=(10, 2))
        return f

    # ══════════════════════════════════════════════════════════════════════════
    # DASHBOARD
    # ══════════════════════════════════════════════════════════════════════════

    def _build_dashboard(self, parent: tk.Frame):
        parent.columnconfigure((0, 1, 2), weight=1, uniform="dash_top")
        parent.rowconfigure(3, weight=1)

        # ── Row 0: top 3 cards ────────────────────────────────────────────────

        # Environment card
        env_c = self._card(parent, "Environment")
        env_c.grid(row=0, column=0, sticky="ew", padx=(16, 5), pady=(16, 5))

        badge_row = tk.Frame(env_c, bg=T.BG_CARD)
        badge_row.pack(fill="x", padx=14, pady=(4, 8))
        self._dash_env_badge = tk.Label(badge_row, text="—", fg=T.BG_DARK, bg=T.ACCENT,
                                        font=("Segoe UI", 7, "bold"), padx=8, pady=2)
        self._dash_env_badge.pack(side="right")

        _label(env_c, "CURRENT ENV", tiny=True).pack(anchor="w", padx=14, pady=(0, 3))
        self._env_dd = ttk.Combobox(env_c, textvariable=self._env_var,
                                    values=["Prod", "UAT", "Dev", "local"],
                                    state="readonly", width=22)
        self._env_dd.pack(fill="x", padx=14, pady=(0, 8))

        btn_row = tk.Frame(env_c, bg=T.BG_CARD)
        btn_row.pack(fill="x", padx=14, pady=(0, 10))
        _btn(btn_row, "⚡  Apply", lambda: self._save_env("ENV", self._env_var.get()),
             style="accent").pack(side="left")

        _label(env_c, ENV_FILE_PATH, tiny=True, bg=T.BG_CARD).pack(
            anchor="w", padx=14, pady=(0, 10))

        # Active Build card
        build_c = self._card(parent, "Active Build")
        build_c.grid(row=0, column=1, sticky="ew", padx=5, pady=(16, 5))

        _label(build_c, "VERSION (V8)", tiny=True).pack(anchor="w", padx=14, pady=(8, 3))
        ver_row = tk.Frame(build_c, bg=T.BG_CARD)
        ver_row.pack(fill="x", padx=14, pady=(0, 6))
        self._v8_var = tk.StringVar(value="Checking…")
        self._v8_dd = ttk.Combobox(ver_row, textvariable=self._v8_var, state="readonly", width=14)
        self._v8_dd.pack(side="left", fill="x", expand=True, padx=(0, 6))
        _btn(ver_row, "Save", lambda: self._save_env("VERSION", self._v8_var.get()),
             style="ghost_card").pack(side="left")

        self._counts_var = tk.StringVar(value="v8: —  ·  v6: —  ·  local: —")
        _label(build_c, bg=T.BG_CARD, textvariable=self._counts_var, muted=True).pack(
            anchor="w", padx=14, pady=(0, 6))

        deploy_row = tk.Frame(build_c, bg=T.BG_CARD)
        deploy_row.pack(fill="x", padx=14, pady=(0, 10))
        _btn(deploy_row, "Deploy →", lambda: self._nav_to("Deploy"),
             style="ghost_card").pack(side="right")

        # Quick Actions card
        qa_c = self._card(parent, "Quick Actions")
        qa_c.grid(row=0, column=2, sticky="ew", padx=(5, 16), pady=(16, 5))

        qa_grid = tk.Frame(qa_c, bg=T.BG_CARD)
        qa_grid.pack(fill="x", padx=14, pady=(4, 10))
        qa_grid.columnconfigure((0, 1), weight=1)
        qa_actions = [
            ("⊕  Support Logs", lambda: self._nav_to("Support Logs")),
            ("≡  View Configs",  lambda: self._nav_to("Configs")),
            ("⋯  Registry",      lambda: self._nav_to("Registry")),
            ("↑  Deploy Build",  lambda: self._nav_to("Deploy")),
        ]
        for i, (lbl, cmd) in enumerate(qa_actions):
            _btn(qa_grid, lbl, cmd, style="ghost_card", anchor="w").grid(
                row=i // 2, column=i % 2, sticky="ew", padx=2, pady=2)

        # ── Row 1: Health Checks ──────────────────────────────────────────────
        hc_c = self._card(parent, "Health Checks")
        hc_c.grid(row=1, column=0, columnspan=3, sticky="ew", padx=16, pady=5)

        hc_header = tk.Frame(hc_c, bg=T.BG_CARD)
        hc_header.pack(fill="x", padx=14, pady=(4, 2))
        self._hc_last_run = _label(hc_header, "Not yet run", tiny=True, bg=T.BG_CARD)
        self._hc_last_run.pack(side="left")
        _btn(hc_header, "↺  Refresh", self._hc_run,
             style="ghost_card").pack(side="right")

        self._hc_inner = tk.Frame(hc_c, bg=T.BG_CARD)
        self._hc_inner.pack(fill="x", padx=14, pady=(0, 10))
        _label(self._hc_inner, "Click Refresh to run health checks", muted=True, bg=T.BG_CARD).pack(anchor="w")

        # ── Row 2: Local Directories ──────────────────────────────────────────
        dirs_c = self._card(parent, "Local Directories")
        dirs_c.grid(row=2, column=0, columnspan=3, sticky="ew", padx=16, pady=5)

        _DIR_ROWS = [
            ("Logs",      T.ACCENT,    LOG_DIR_PATH),
            ("Loader",    T.ACCENT,    LOADER_PATH),
            ("Add-In v8", T.INFO,      ADD_IN_PATH),
            ("Add-In v6", T.INFO,      LOCAL_TEST_PATH),
            ("Test",      T.TEXT_SEC,  BASE_LOCAL_PATH),
            ("Templates", T.WARN,      TEMPLATES_PATH),
        ]
        for name, color, path in _DIR_ROWS:
            row = tk.Frame(dirs_c, bg=T.BG_CARD)
            row.pack(fill="x", padx=14, pady=1)
            tk.Frame(row, bg=color, width=3).pack(side="left", fill="y", padx=(0, 10))
            _label(row, name, bold=True, bg=T.BG_CARD, width=10, anchor="w").pack(side="left")
            _label(row, path, mono=True, muted=True, bg=T.BG_CARD,
                   anchor="w").pack(side="left", fill="x", expand=True)
            _btn(row, "Open", lambda p=path: self._explore(p),
                 style="ghost_card", pady=3).pack(side="right", padx=(0, 2))

        tk.Frame(dirs_c, bg=T.BG_CARD, height=6).pack()

    def _hc_run(self):
        for w in self._hc_inner.winfo_children():
            w.destroy()
        self._hc_last_run.config(text="Checking…")

        checks = [
            ("Loader path",    lambda: os.path.exists(LOADER_PATH)),
            ("Network share v8", lambda: os.path.isdir(NETWORK_PATH_8)),
            ("Network share v6", lambda: os.path.isdir(NETWORK_PATH_6)),
            ("Registry: NinetyOne", lambda: scan_registry_for_ninetyone()[1] != []),
            ("Excel installed", lambda: os.path.exists(
                r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")),
            ("Log directory",  lambda: os.path.isdir(LOG_DIR_PATH)),
        ]

        self._hc_inner.columnconfigure((0, 1, 2), weight=1)

        def run():
            results = []
            for label, test in checks:
                try:
                    ok = bool(test())
                except Exception:
                    ok = False
                results.append((label, ok))
            self.root.after(0, lambda: self._hc_update(results))

        threading.Thread(target=run, daemon=True).start()

    def _hc_update(self, results):
        for w in self._hc_inner.winfo_children():
            w.destroy()
        self._hc_inner.columnconfigure((0, 1, 2), weight=1)
        self._hc_last_run.config(text="Just run")
        for i, (label, ok) in enumerate(results):
            f = tk.Frame(self._hc_inner, bg=T.BG_CARD)
            f.grid(row=i // 3, column=i % 3, sticky="w", padx=(0, 20), pady=4)
            tk.Label(f, text="●", fg=T.ACCENT if ok else T.ERROR,
                     bg=T.BG_CARD, font=("Segoe UI", 8)).pack(side="left")
            _label(f, f"  {label}", bold=True, bg=T.BG_CARD).pack(side="left")

    # ══════════════════════════════════════════════════════════════════════════
    # DEPLOY
    # ══════════════════════════════════════════════════════════════════════════

    def _build_deploy(self, parent: tk.Frame):
        parent.columnconfigure(0, weight=3)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(2, weight=1)

        # ── Auth card ─────────────────────────────────────────────────────────
        auth_c = self._card(parent, "Azure DevOps Authentication")
        auth_c.grid(row=0, column=0, columnspan=2, sticky="ew", padx=16, pady=(16, 5))

        inner = tk.Frame(auth_c, bg=T.BG_CARD)
        inner.pack(fill="x", padx=14, pady=(4, 10))
        inner.columnconfigure(0, weight=1)

        _label(inner, "PERSONAL ACCESS TOKEN", tiny=True, bg=T.BG_CARD).grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 4))

        self._d_pat_var = tk.StringVar()
        ttk.Entry(inner, textvariable=self._d_pat_var, show="●", width=50).grid(
            row=1, column=0, sticky="ew", padx=(0, 8))

        self._d_load_btn = _btn(inner, "↺  Reload Builds", self._d_load)
        self._d_load_btn.grid(row=1, column=1, padx=(0, 6))
        _btn(inner, "?", self._d_open_guide).grid(row=1, column=2)

        self._d_remember = tk.BooleanVar(value=False)
        remember_row = tk.Frame(inner, bg=T.BG_CARD)
        remember_row.grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Checkbutton(remember_row, text="Remember PAT", variable=self._d_remember,
                        style="TCheckbutton").pack(side="left")
        self._d_pat_status = _label(remember_row, "", muted=True, bg=T.BG_CARD)
        self._d_pat_status.pack(side="left", padx=12)

        # ── Builds table ─────────────────────────────────────────────────────
        self._d_builds_title = tk.StringVar(value=f"Recent Builds  ·  Pipeline #{ADO_PIPELINE_ID}")
        builds_c = self._card(parent)
        builds_c.grid(row=1, column=0, rowspan=2, sticky="nsew", padx=(16, 5), pady=5)
        builds_c.rowconfigure(1, weight=1)
        builds_c.columnconfigure(0, weight=1)

        bh = tk.Frame(builds_c, bg=T.BG_CARD)
        bh.pack(fill="x", padx=14, pady=(10, 4))
        _label(bh, textvariable=self._d_builds_title, bold=True, bg=T.BG_CARD).pack(side="left")

        tree_f = tk.Frame(builds_c, bg=T.BG_CARD)
        tree_f.pack(fill="both", expand=True, padx=14, pady=(0, 4))
        tree_f.rowconfigure(0, weight=1)
        tree_f.columnconfigure(0, weight=1)

        cols = ("build", "result", "finished", "branch", "by")
        hdrs = [("Build #", 130), ("Result", 90), ("Finished (UTC)", 150), ("Branch", 120), ("Requested By", 150)]
        self._d_tree = ttk.Treeview(tree_f, columns=cols, show="headings", selectmode="browse")
        for (h, w), c in zip(hdrs, cols):
            self._d_tree.heading(c, text=h)
            self._d_tree.column(c, width=w, minwidth=70)
        d_sb = ttk.Scrollbar(tree_f, orient="vertical", command=self._d_tree.yview)
        self._d_tree.configure(yscrollcommand=d_sb.set)
        self._d_tree.grid(row=0, column=0, sticky="nsew")
        d_sb.grid(row=0, column=1, sticky="ns")
        self._d_tree.bind("<<TreeviewSelect>>", self._d_on_select)
        self._d_builds: dict = {}

        # ── Deployment log ────────────────────────────────────────────────────
        log_c = self._card(parent, "Deployment Log")
        log_c.grid(row=3, column=0, sticky="ew", padx=(16, 5), pady=(0, 16))

        self._d_log = tk.Text(log_c, bg=T.BG_DARK, fg=T.ACCENT, font=T.F_MONO,
                              height=7, relief="flat", bd=0, state="disabled",
                              insertbackground=T.ACCENT, wrap="none")
        d_log_sb = ttk.Scrollbar(log_c, orient="vertical", command=self._d_log.yview)
        self._d_log.configure(yscrollcommand=d_log_sb.set)

        log_inner = tk.Frame(log_c, bg=T.BG_DARK)
        log_inner.pack(fill="x", padx=14, pady=(4, 10))
        self._d_log.pack(side="left", fill="both", expand=True, in_=log_inner)
        d_log_sb.pack(side="right", fill="y", in_=log_inner)

        # ── Right column ──────────────────────────────────────────────────────
        right = tk.Frame(parent, bg=T.BG_MAIN)
        right.grid(row=1, column=1, rowspan=3, sticky="nsew", padx=(0, 16), pady=(5, 16))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        # Deploy Selected card
        ds_c = self._card(right, "Deploy Selected")
        ds_c.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        fields_f = tk.Frame(ds_c, bg=T.BG_CARD)
        fields_f.pack(fill="x", padx=14, pady=(4, 8))
        self._d_src_build  = tk.StringVar(value="—")
        self._d_src_branch = tk.StringVar(value="—")
        self._d_tgt_ver    = tk.StringVar(value="—")
        self._d_tgt_share  = tk.StringVar(value=NETWORK_PATH_8)
        for row_lbl, var in [("Source build", self._d_src_build),
                              ("Branch",       self._d_src_branch),
                              ("Target version", self._d_tgt_ver),
                              ("Target share",  self._d_tgt_share)]:
            row = tk.Frame(fields_f, bg=T.BG_CARD)
            row.pack(fill="x", pady=2)
            _label(row, row_lbl, muted=True, bg=T.BG_CARD, width=14, anchor="w").pack(side="left")
            _label(row, bg=T.BG_CARD, textvariable=var, bold=True).pack(side="left")

        self._d_progress = ttk.Progressbar(ds_c, mode="determinate")
        self._d_progress.pack(fill="x", padx=14, pady=(2, 0))
        self._d_progress_lbl = _label(ds_c, "", muted=True, bg=T.BG_CARD)
        self._d_progress_lbl.pack(anchor="w", padx=14, pady=(2, 4))

        ds_btns = tk.Frame(ds_c, bg=T.BG_CARD)
        ds_btns.pack(fill="x", padx=14, pady=(4, 12))
        self._d_deploy_btn = _btn(ds_btns, "⚡  Deploy", self._d_run, style="accent")
        self._d_deploy_btn.pack(side="left", padx=(0, 6))
        self._d_deploy_btn.config(state="disabled")
        _btn(ds_btns, "Rollback…", self._d_rollback).pack(side="left")

        # Existing Versions card
        ev_c = self._card(right, "Existing Versions")
        ev_c.grid(row=1, column=0, sticky="nsew")
        ev_c.rowconfigure(1, weight=1)

        ev_hdr = tk.Frame(ev_c, bg=T.BG_CARD)
        ev_hdr.pack(fill="x", padx=14, pady=(4, 4))
        self._d_ev_count = _label(ev_hdr, "", muted=True, bg=T.BG_CARD)
        self._d_ev_count.pack(side="left")
        _btn(ev_hdr, "↺", self._d_refresh_versions,
             style="ghost_card", padx=6, pady=3).pack(side="right")

        ev_scroll_f = tk.Frame(ev_c, bg=T.BG_CARD)
        ev_scroll_f.pack(fill="both", expand=True, padx=14, pady=(0, 10))
        self._d_ev_canvas = tk.Canvas(ev_scroll_f, bg=T.BG_CARD, bd=0, highlightthickness=0)
        ev_vs = ttk.Scrollbar(ev_scroll_f, orient="vertical", command=self._d_ev_canvas.yview)
        self._d_ev_canvas.configure(yscrollcommand=ev_vs.set)
        self._d_ev_canvas.pack(side="left", fill="both", expand=True)
        ev_vs.pack(side="right", fill="y")
        self._d_ev_inner = tk.Frame(self._d_ev_canvas, bg=T.BG_CARD)
        self._d_ev_win = self._d_ev_canvas.create_window((0, 0), window=self._d_ev_inner, anchor="nw")
        self._d_ev_inner.bind("<Configure>", lambda e: self._d_ev_canvas.configure(
            scrollregion=self._d_ev_canvas.bbox("all")))
        self._d_ev_canvas.bind("<Configure>", lambda e: self._d_ev_canvas.itemconfig(
            self._d_ev_win, width=e.width))

        # Load saved PAT
        self._d_load_saved_pat()

    def _d_load_saved_pat(self):
        saved = _read_pat()
        if saved:
            self._d_pat_var.set(saved)
            self._d_remember.set(True)
            self._d_pat_status.config(text="Loaded from saved file")

    def _d_log_write(self, msg: str):
        self._d_log.configure(state="normal")
        self._d_log.insert(tk.END, msg + "\n")
        self._d_log.see(tk.END)
        self._d_log.configure(state="disabled")

    def _d_open_guide(self):
        if os.path.exists(_PAT_GUIDE):
            os.startfile(_PAT_GUIDE)
        else:
            messagebox.showinfo("PAT Guide", f"Not found at:\n{_PAT_GUIDE}")

    def _d_load(self):
        pat = self._d_pat_var.get().strip()
        if not pat:
            messagebox.showwarning("PAT Required", "Enter a Personal Access Token.")
            return
        self._d_load_btn.config(state="disabled")
        self._d_tree.delete(*self._d_tree.get_children())
        self._d_log_write("Fetching builds from Azure DevOps…")
        threading.Thread(target=self._d_load_bg, args=(pat,), daemon=True).start()

    def _d_load_bg(self, pat):
        ok, result = fetch_builds(pat)
        self.root.after(0, lambda: self._d_on_loaded(ok, result, pat))

    def _d_on_loaded(self, ok, result, pat):
        self._d_load_btn.config(state="normal")
        if not ok:
            self._d_log_write(f"Error: {result}")
            messagebox.showerror("Load Failed", result)
            return
        if self._d_remember.get():
            _write_pat(pat)
            self._d_pat_status.config(text="PAT saved")
        self._d_builds = {b["id"]: b for b in result}
        for b in result:
            finished = (b.get("finishTime", "") or "")[:16].replace("T", " ")
            branch = b.get("sourceBranch", "").replace("refs/heads/", "")
            by = b.get("requestedFor", {}).get("displayName", "")
            outcome = b.get("result") or b.get("status", "")
            self._d_tree.insert("", "end", iid=str(b["id"]),
                                values=(b.get("buildNumber", b["id"]), outcome, finished, branch, by))
        self._d_log_write(f"Loaded {len(result)} build(s). Select one to deploy.")
        self._d_refresh_versions()

    def _d_on_select(self, _=None):
        sel = self._d_tree.selection()
        if not sel:
            self._d_deploy_btn.config(state="disabled")
            return
        build = self._d_builds.get(int(sel[0]), {})
        self._d_src_build.set(build.get("buildNumber", sel[0]))
        self._d_src_branch.set(build.get("sourceBranch", "").replace("refs/heads/", ""))
        self._d_tgt_ver.set("Checking…")
        self._d_deploy_btn.config(state="disabled")
        threading.Thread(target=self._d_resolve_ver, daemon=True).start()

    def _d_resolve_ver(self):
        ok, ver = get_next_version(NETWORK_PATH_8)
        def upd():
            if ok:
                self._d_tgt_ver.set(ver)
                self._d_deploy_btn.config(state="normal")
            else:
                self._d_tgt_ver.set(f"Error: {ver}")
        self.root.after(0, upd)

    def _d_run(self):
        sel = self._d_tree.selection()
        if not sel: return
        build_id = int(sel[0])
        build = self._d_builds.get(build_id, {})
        build_num = build.get("buildNumber", build_id)
        pat = self._d_pat_var.get().strip()
        ok, version = get_next_version(NETWORK_PATH_8)
        if not ok:
            messagebox.showerror("Version Error", version)
            return
        if not messagebox.askyesno("Confirm Deploy",
                                   f"Deploy build {build_num} as {version}?\n\n"
                                   f"Target:\n  {NETWORK_PATH_8}\\{version}"):
            return
        self._d_deploy_btn.config(state="disabled")
        self._d_load_btn.config(state="disabled")
        self._d_progress["value"] = 0
        self._d_log_write(f"─── Deploying {build_num} as {version} ───")
        threading.Thread(target=self._d_run_bg, args=(pat, build_id, version), daemon=True).start()

    def _d_run_bg(self, pat, build_id, version):
        self.root.after(0, lambda: self._d_log_write("Downloading artifact…"))

        def on_progress(done, total):
            pct = int(done / total * 100)
            mb_d, mb_t = done / 1048576, total / 1048576
            self.root.after(0, lambda p=pct: self._d_progress.configure(value=p))
            self.root.after(0, lambda p=pct, d=mb_d, t=mb_t:
                            self._d_progress_lbl.config(text=f"{p}%  ({d:.1f} / {t:.1f} MB)"))

        ok, result = download_artifact_zip(pat, build_id, progress_cb=on_progress)
        if not ok:
            self.root.after(0, lambda: self._d_finish(False, result))
            return
        self.root.after(0, lambda: self._d_log_write("Extracting to network share…"))
        ok, dest = deploy_zip_to_network(result, version)
        self.root.after(0, lambda: self._d_finish(ok, dest))

    def _d_finish(self, ok, result):
        self._d_deploy_btn.config(state="normal")
        self._d_load_btn.config(state="normal")
        if ok:
            self._d_progress["value"] = 100
            self._d_log_write(f"✓ Deployed to: {result}")
            messagebox.showinfo("Done", f"Deployed to:\n{result}")
            self._d_refresh_versions()
        else:
            self._d_progress["value"] = 0
            self._d_log_write(f"✗ Failed: {result}")
            messagebox.showerror("Failed", result)

    def _d_rollback(self):
        ok, versions = get_existing_versions(NETWORK_PATH_8)
        if not ok:
            messagebox.showerror("Error", versions)
            return
        if not versions:
            messagebox.showinfo("Rollback", "No existing versions found.")
            return
        dlg = tk.Toplevel(self.root)
        dlg.title("Rollback — Select Version")
        dlg.geometry("320x300")
        dlg.configure(bg=T.BG_MAIN)
        dlg.grab_set()
        _label(dlg, "Select version to copy as new release:",
               bg=T.BG_MAIN).pack(padx=14, pady=(14, 4), anchor="w")
        lb = tk.Listbox(dlg, bg=T.BG_CARD, fg=T.TEXT_PRI,
                        selectbackground=T.ACCENT, selectforeground="#000000",
                        font=T.F_MONO, relief="flat", bd=0)
        lb.pack(fill="both", expand=True, padx=14, pady=4)
        for v in versions:
            lb.insert(tk.END, v)
        if versions:
            lb.selection_set(0)
        btn_row = tk.Frame(dlg, bg=T.BG_MAIN)
        btn_row.pack(fill="x", padx=14, pady=(4, 14))

        def confirm():
            sel = lb.curselection()
            if sel:
                ver = lb.get(sel[0])
                dlg.destroy()
                self._d_do_rollback(ver)

        _btn(btn_row, "Cancel", dlg.destroy, style="ghost").pack(side="right")
        _btn(btn_row, "Rollback to Selected", confirm,
             style="accent").pack(side="right", padx=(0, 6))

    def _d_do_rollback(self, source):
        ok, nv = get_next_version(NETWORK_PATH_8)
        if not ok:
            messagebox.showerror("Error", nv)
            return
        if not messagebox.askyesno("Confirm", f"Copy {source} → {nv}?"):
            return
        self._d_log_write(f"─── Rolling back: {source} → {nv} ───")
        self._d_load_btn.config(state="disabled")
        threading.Thread(target=self._d_rollback_bg, args=(source,), daemon=True).start()

    def _d_rollback_bg(self, source):
        ok, result = rollback_to_network(source)
        def finish():
            self._d_load_btn.config(state="normal")
            if ok:
                nv, path = result
                self._d_log_write(f"✓ Rollback: {nv} at {path}")
                messagebox.showinfo("Done", f"Created {nv} from {source}.")
                self._d_refresh_versions()
            else:
                self._d_log_write(f"✗ Rollback failed: {result}")
                messagebox.showerror("Failed", result)
        self.root.after(0, finish)

    def _d_refresh_versions(self):
        for w in self._d_ev_inner.winfo_children():
            w.destroy()
        _label(self._d_ev_inner, "Loading…", muted=True, bg=T.BG_CARD).pack(anchor="w")
        threading.Thread(target=self._d_refresh_versions_bg, daemon=True).start()

    def _d_refresh_versions_bg(self):
        ok, versions = get_existing_versions(NETWORK_PATH_8)
        current = self._ver_var.get()
        def upd():
            for w in self._d_ev_inner.winfo_children():
                w.destroy()
            if not ok:
                _label(self._d_ev_inner, str(versions), fg=T.ERROR, bg=T.BG_CARD).pack(anchor="w")
                return
            self._d_ev_count.config(text=f"{len(versions)} on share")
            for v in versions:
                row = tk.Frame(self._d_ev_inner, bg=T.BG_CARD)
                row.pack(fill="x", pady=2)
                _label(row, v, mono=True, bg=T.BG_CARD).pack(side="left")
                if v == current:
                    _label(row, "Current", accent=True, bg=T.BG_CARD).pack(side="left", padx=6)
                else:
                    _btn(row, "Roll back", lambda ver=v: self._d_do_rollback(ver),
                         style="ghost_card", padx=8, pady=2).pack(side="right")
        self.root.after(0, upd)

    # ══════════════════════════════════════════════════════════════════════════
    # CONFIGS
    # ══════════════════════════════════════════════════════════════════════════

    def _build_configs(self, parent: tk.Frame):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        # Controls card
        ctrl_c = self._card(parent, "Config Variables Viewer")
        ctrl_c.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 5))

        ctrl_inner = tk.Frame(ctrl_c, bg=T.BG_CARD)
        ctrl_inner.pack(fill="x", padx=14, pady=(4, 10))

        # Source
        src_f = tk.Frame(ctrl_inner, bg=T.BG_CARD)
        src_f.pack(side="left", padx=(0, 20))
        _label(src_f, "SOURCE", tiny=True, bg=T.BG_CARD).pack(anchor="w")
        self._cfg_src_var = tk.StringVar(value="Network .Net 8")
        self._cfg_src_cb = ttk.Combobox(src_f, textvariable=self._cfg_src_var,
                                        values=["Network .Net 8", "Network .Net 6", "Local Test Path"],
                                        state="readonly", width=18)
        self._cfg_src_cb.pack(pady=(3, 0))
        self._cfg_src_cb.bind("<<ComboboxSelected>>", self._cfg_on_src)

        # Version
        ver_f = tk.Frame(ctrl_inner, bg=T.BG_CARD)
        ver_f.pack(side="left", padx=(0, 20))
        _label(ver_f, "VERSION", tiny=True, bg=T.BG_CARD).pack(anchor="w")
        self._cfg_ver_var = tk.StringVar()
        self._cfg_ver_cb = ttk.Combobox(ver_f, textvariable=self._cfg_ver_var, state="readonly", width=14)
        self._cfg_ver_cb.pack(pady=(3, 0))
        self._cfg_ver_cb.bind("<<ComboboxSelected>>", self._cfg_on_ver)

        # Environment
        env_f = tk.Frame(ctrl_inner, bg=T.BG_CARD)
        env_f.pack(side="left", padx=(0, 20))
        _label(env_f, "ENVIRONMENT", tiny=True, bg=T.BG_CARD).pack(anchor="w")
        self._cfg_env_var = tk.StringVar()
        self._cfg_env_cb = ttk.Combobox(env_f, textvariable=self._cfg_env_var, state="readonly", width=12)
        self._cfg_env_cb.pack(pady=(3, 0))
        self._cfg_env_cb.bind("<<ComboboxSelected>>", self._cfg_on_env)

        self._cfg_data: list = []

        # Table card
        tbl_c = self._card(parent)
        tbl_c.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        tbl_c.rowconfigure(1, weight=1)
        tbl_c.columnconfigure(0, weight=1)

        self._cfg_title_lbl = _label(tbl_c, "", bold=True, bg=T.BG_CARD)
        self._cfg_title_lbl.pack(anchor="w", padx=14, pady=(10, 4))

        tree_f = tk.Frame(tbl_c, bg=T.BG_CARD)
        tree_f.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        tree_f.rowconfigure(0, weight=1)
        tree_f.columnconfigure(0, weight=1)

        self._cfg_tree = ttk.Treeview(tree_f, columns=("P", "V"), show="headings")
        self._cfg_tree.heading("P", text="PARAMETER")
        self._cfg_tree.heading("V", text="VALUE")
        self._cfg_tree.column("P", width=220, minwidth=150)
        self._cfg_tree.column("V", width=500, stretch=True)
        cfg_sb = ttk.Scrollbar(tree_f, orient="vertical", command=self._cfg_tree.yview)
        self._cfg_tree.configure(yscrollcommand=cfg_sb.set)
        self._cfg_tree.grid(row=0, column=0, sticky="nsew")
        cfg_sb.grid(row=0, column=1, sticky="ns")

        self.root.after(600, self._cfg_on_src)

    def _cfg_on_src(self, _=None):
        s = self._cfg_src_var.get()
        versions = (self._versions["v8"] if "8" in s else
                    self._versions["v6"] if "6" in s else
                    self._versions["local"])
        self._cfg_ver_cb["values"] = versions
        if versions:
            self._cfg_ver_cb.set(versions[0])
            self._cfg_on_ver()
        else:
            self._cfg_ver_cb.set("N/A")
            self._cfg_tree.delete(*self._cfg_tree.get_children())

    def _cfg_on_ver(self, _=None):
        from config import NETWORK_PATH_8 as N8, NETWORK_PATH_6 as N6, CONFIG_FILE_NAME as CF
        s, v = self._cfg_src_var.get(), self._cfg_ver_var.get()
        if v == "N/A": return
        base = N8 if "8" in s else (N6 if "6" in s else LOCAL_TEST_PATH)
        self._cfg_load(os.path.join(base, v, CF))
        self._cfg_title_lbl.config(text=f"{v}  ·  {self._cfg_env_var.get()}")

    def _cfg_load(self, path):
        self._cfg_tree.delete(*self._cfg_tree.get_children())
        self._cfg_env_cb["values"] = []
        self._cfg_env_cb.set("")
        self._cfg_data = []
        if not os.path.exists(path):
            self._cfg_tree.insert("", "end", values=("File not found", path))
            return
        try:
            with open(path, encoding="utf-8") as f:
                self._cfg_data = json.load(f)
            envs = [i.get("environment") for i in self._cfg_data if isinstance(i, dict)]
            self._cfg_env_cb["values"] = envs
            cur = self._env_var.get()
            self._cfg_env_cb.set(cur if cur in envs else (envs[0] if envs else ""))
            self._cfg_on_env()
        except Exception as e:
            self._cfg_tree.insert("", "end", values=("JSON Error", str(e)))

    def _cfg_on_env(self, _=None):
        self._cfg_tree.delete(*self._cfg_tree.get_children())
        env = self._cfg_env_var.get()
        for item in self._cfg_data:
            if item.get("environment") == env:
                for k, v in item.get("parameters", {}).items():
                    self._cfg_tree.insert("", "end", values=(k, v))
                break
        self._cfg_title_lbl.config(
            text=f"{self._cfg_ver_var.get()}  ·  {env}" if self._cfg_ver_var.get() else env)

    # ══════════════════════════════════════════════════════════════════════════
    # SUPPORT LOGS
    # ══════════════════════════════════════════════════════════════════════════

    def _build_logs(self, parent: tk.Frame):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        self._logs_files: list = []
        self._logs_svc_map: dict = {}
        self._logs_entries: list = []
        self._logs_search_hits: list = []
        self._logs_search_idx  = -1
        self._logs_search_q    = ""

        # ── Toolbar card ──────────────────────────────────────────────────────
        tb_c = self._card(parent)
        tb_c.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 0))

        tb_hdr = tk.Frame(tb_c, bg=T.BG_CARD)
        tb_hdr.pack(fill="x", padx=14, pady=(10, 4))

        title_f = tk.Frame(tb_hdr, bg=T.BG_CARD)
        title_f.pack(side="left")
        _label(title_f, "FOXL Support Log Viewer", bold=True, bg=T.BG_CARD).pack(anchor="w")
        self._logs_stats = _label(title_f, "", muted=True, bg=T.BG_CARD)
        self._logs_stats.pack(anchor="w")

        right_tb = tk.Frame(tb_hdr, bg=T.BG_CARD)
        right_tb.pack(side="right")
        _btn(right_tb, "⊞  Compress & Save Logs", self._logs_export).pack(side="right", padx=(6, 0))
        _btn(right_tb, "↺  Refresh", self._logs_load, style="ghost_card").pack(side="right")

        # Filter bar
        fb = tk.Frame(tb_c, bg=T.BG_CARD)
        fb.pack(fill="x", padx=14, pady=(0, 10))

        _label(fb, "SERVICE", tiny=True, bg=T.BG_CARD).pack(side="left")
        self._logs_svc_var = tk.StringVar(value="All Services")
        self._logs_svc_dd = ttk.Combobox(fb, textvariable=self._logs_svc_var,
                                         state="readonly", width=16)
        self._logs_svc_dd.pack(side="left", padx=(4, 16))
        self._logs_svc_dd.bind("<<ComboboxSelected>>", lambda _: self._logs_reload_unified())

        _label(fb, "ORDER", tiny=True, bg=T.BG_CARD).pack(side="left")
        self._logs_order_var = tk.StringVar(value="Oldest First")
        order_dd = ttk.Combobox(fb, textvariable=self._logs_order_var,
                                values=["Oldest First", "Newest First"], state="readonly", width=12)
        order_dd.pack(side="left", padx=(4, 16))
        order_dd.bind("<<ComboboxSelected>>", lambda _: self._logs_reload_unified())

        _label(fb, "FIND", tiny=True, bg=T.BG_CARD).pack(side="left")
        self._logs_find_var = tk.StringVar()
        find_entry = ttk.Entry(fb, textvariable=self._logs_find_var, width=22)
        find_entry.pack(side="left", padx=(4, 2))
        find_entry.bind("<Return>", self._logs_search)
        _btn(fb, "▲", self._logs_find_prev, padx=6, pady=3).pack(side="left", padx=1)
        _btn(fb, "▼", self._logs_find_next, padx=6, pady=3).pack(side="left", padx=1)

        # Timeline/File View toggle
        self._logs_view = tk.StringVar(value="Timeline")
        self._logs_toggle: dict[str, tk.Button] = {}
        toggle_f = tk.Frame(fb, bg=T.BG_CARD)
        toggle_f.pack(side="right")
        for mode in ("Timeline", "File View"):
            b = tk.Button(toggle_f, text=mode, bg=T.BG_INPUT, fg=T.TEXT_SEC,
                          activebackground=T.BORDER, relief="flat", bd=0,
                          padx=10, pady=4, cursor="hand2", font=T.F_UI,
                          command=lambda m=mode: self._logs_set_view(m))
            b.pack(side="left", padx=1)
            self._logs_toggle[mode] = b

        parent.bind_all("<Control-f>", lambda e: find_entry.focus_set())

        # ── Content card ──────────────────────────────────────────────────────
        content_c = self._card(parent)
        content_c.grid(row=1, column=0, sticky="nsew", padx=16, pady=(6, 16))
        content_c.rowconfigure(0, weight=1)
        content_c.columnconfigure(0, weight=1)

        # Timeline frame
        self._logs_tl_frame = tk.Frame(content_c, bg=T.BG_CARD)
        self._logs_tl_frame.columnconfigure(0, weight=1)
        self._logs_tl_frame.rowconfigure(0, weight=1)

        self._logs_txt = tk.Text(self._logs_tl_frame, bg=T.BG_CARD, fg=T.TEXT_PRI,
                                 font=T.F_MONO, relief="flat", bd=0, state="disabled",
                                 wrap="none", selectbackground=T.BG_INPUT,
                                 insertbackground=T.TEXT_PRI)
        tl_sby = ttk.Scrollbar(self._logs_tl_frame, orient="vertical", command=self._logs_txt.yview)
        tl_sbx = ttk.Scrollbar(self._logs_tl_frame, orient="horizontal", command=self._logs_txt.xview)
        self._logs_txt.configure(yscrollcommand=tl_sby.set, xscrollcommand=tl_sbx.set)
        self._logs_txt.grid(row=0, column=0, sticky="nsew")
        tl_sby.grid(row=0, column=1, sticky="ns")
        tl_sbx.grid(row=1, column=0, sticky="ew")

        # File View frame
        self._logs_fv_frame = tk.Frame(content_c, bg=T.BG_CARD)
        self._logs_fv_frame.rowconfigure(1, weight=1)
        self._logs_fv_frame.columnconfigure(0, weight=1)

        fv_ctrl = tk.Frame(self._logs_fv_frame, bg=T.BG_CARD)
        fv_ctrl.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        _label(fv_ctrl, "File:", muted=True, bg=T.BG_CARD).pack(side="left")
        self._logs_file_var = tk.StringVar()
        self._logs_file_dd = ttk.Combobox(fv_ctrl, textvariable=self._logs_file_var,
                                          state="readonly", width=50)
        self._logs_file_dd.pack(side="left", padx=6)
        self._logs_file_dd.bind("<<ComboboxSelected>>", self._logs_load_file)

        self._logs_file_txt = tk.Text(self._logs_fv_frame, bg=T.BG_CARD, fg=T.TEXT_PRI,
                                      font=T.F_MONO, relief="flat", bd=0, state="disabled",
                                      selectbackground=T.BG_INPUT)
        fv_sb = ttk.Scrollbar(self._logs_fv_frame, orient="vertical",
                               command=self._logs_file_txt.yview)
        self._logs_file_txt.configure(yscrollcommand=fv_sb.set)
        self._logs_file_txt.grid(row=1, column=0, sticky="nsew")
        fv_sb.grid(row=1, column=1, sticky="ns")

        # Status bar
        self._logs_status = _label(content_c, "—", muted=True, bg=T.BG_CARD)
        self._logs_status.grid(row=1, column=0, sticky="w", padx=14, pady=(2, 6))

        self._logs_set_view("Timeline")

    def _logs_set_view(self, mode: str):
        self._logs_view.set(mode)
        for m, btn in self._logs_toggle.items():
            btn.config(bg=T.ACCENT if m == mode else T.BG_INPUT,
                       fg="#000000" if m == mode else T.TEXT_SEC,
                       font=T.F_BOLD if m == mode else T.F_UI)
        if mode == "Timeline":
            self._logs_fv_frame.grid_remove()
            self._logs_tl_frame.grid(row=0, column=0, sticky="nsew",
                                     padx=14, pady=(4, 0))
        else:
            self._logs_tl_frame.grid_remove()
            self._logs_fv_frame.grid(row=0, column=0, sticky="nsew",
                                     padx=14, pady=(4, 0))

    def _logs_load(self):
        self._logs_files = get_log_files(LOG_DIR_PATH)
        services = get_services(self._logs_files)
        self._logs_svc_map = {"All Services": None}
        self._logs_svc_map.update({d: r for r, d in services})
        self._logs_svc_dd["values"] = list(self._logs_svc_map.keys())
        self._logs_svc_var.set("All Services")
        self._logs_reload_file_dd()
        self._logs_reload_unified()

    def _logs_get_filtered(self):
        svc = self._logs_svc_map.get(self._logs_svc_var.get())
        if svc is None:
            return self._logs_files
        return [f for f in self._logs_files
                if extract_service(os.path.basename(f)) == svc]

    def _logs_reload_unified(self):
        self._logs_txt.configure(state="normal")
        self._logs_txt.delete("1.0", tk.END)
        self._logs_txt.insert(tk.END, "Compiling timeline…\n")
        self._logs_txt.configure(state="disabled")
        self._logs_status.config(text="Loading…")
        threading.Thread(target=self._logs_unified_bg, daemon=True).start()

    def _logs_unified_bg(self):
        files = self._logs_get_filtered()
        reverse = self._logs_order_var.get() == "Newest First"
        entries, errors = generate_unified_timeline(files, reverse=reverse)
        self._logs_entries = entries
        data = "".join(t for _, t in entries) if entries else "No logs found."
        if errors:
            data += "\n\n--- Parse errors ---\n" + "\n".join(f"  {n}: {m}" for n, m in errors)
        err_count  = sum(1 for _, t in entries if "ERROR" in t)
        warn_count = sum(1 for _, t in entries if "WARN"  in t)

        def upd():
            self._logs_txt.configure(state="normal")
            self._logs_txt.delete("1.0", tk.END)
            self._logs_txt.insert(tk.END, data)
            self._logs_apply_hl(self._logs_txt)
            self._logs_txt.configure(state="disabled")
            self._logs_stats.config(
                text=f"{len(entries)} entries · {err_count} errors · {warn_count} warnings")
            self._logs_update_status()

        self.root.after(0, upd)

    _HL_MAP = {
        "ERROR": (T.ERROR,    "hl_err"),
        "WARN":  (T.WARN,     "hl_warn"),
        "INFO":  (T.INFO,     "hl_info"),
    }

    def _logs_apply_hl(self, txt: tk.Text):
        for kw, (color, tag) in self._HL_MAP.items():
            txt.tag_config(tag, foreground=color)
            txt.tag_remove(tag, "1.0", tk.END)
            start = "1.0"
            while True:
                pos = txt.search(kw, start, stopindex=tk.END)
                if not pos: break
                end = f"{pos}+{len(kw)}c"
                txt.tag_add(tag, pos, end)
                start = end

    def _logs_reload_file_dd(self):
        files = self._logs_get_filtered()
        names = [os.path.basename(f) for f in files]
        self._logs_file_dd["values"] = names
        if names:
            self._logs_file_dd.set(names[0])
            self._logs_load_file()
        else:
            self._logs_file_dd.set("")

    def _logs_load_file(self, _=None):
        name = self._logs_file_var.get()
        path = next((f for f in self._logs_get_filtered()
                     if os.path.basename(f) == name), None)
        self._logs_file_txt.configure(state="normal")
        self._logs_file_txt.delete("1.0", tk.END)
        if path and os.path.exists(path):
            try:
                with open(path, encoding="utf-8", errors="replace") as f:
                    self._logs_file_txt.insert(tk.END, f.read())
                self._logs_apply_hl(self._logs_file_txt)
            except OSError as e:
                self._logs_file_txt.insert(tk.END, str(e))
        self._logs_file_txt.configure(state="disabled")

    def _logs_update_status(self):
        q = self._logs_search_q
        if q and self._logs_search_hits:
            self._logs_status.config(
                text=f"{self._logs_search_idx + 1} of {len(self._logs_search_hits)} matches for '{q}'")
            return
        if q:
            self._logs_status.config(text=f"No matches for '{q}'")
            return
        entries = self._logs_entries
        if entries:
            ts = [e[0] for e in entries]
            self._logs_status.config(text=f"{len(entries)} entries  ·  {min(ts)}  →  {max(ts)}")
        else:
            self._logs_status.config(text="No entries")

    def _logs_active_txt(self) -> tk.Text:
        return self._logs_txt if self._logs_view.get() == "Timeline" else self._logs_file_txt

    def _logs_search(self, _=None):
        q = self._logs_find_var.get().strip()
        self._logs_search_q = q
        txt = self._logs_active_txt()
        for tag in ("hl_search", "hl_search_cur"):
            txt.tag_remove(tag, "1.0", tk.END)
        self._logs_search_hits = []
        if not q:
            self._logs_update_status()
            return
        start = "1.0"
        while True:
            pos = txt.search(q, start, stopindex=tk.END, nocase=True)
            if not pos: break
            self._logs_search_hits.append(pos)
            start = f"{pos}+{len(q)}c"
        txt.tag_config("hl_search", background="#ffff88", foreground="black")
        for pos in self._logs_search_hits:
            txt.tag_add("hl_search", pos, f"{pos}+{len(q)}c")
        self._logs_search_idx = 0 if self._logs_search_hits else -1
        if self._logs_search_hits:
            self._logs_scroll_to_cur(txt)
        self._logs_update_status()

    def _logs_find_next(self):
        if not self._logs_search_hits: self._logs_search(); return
        self._logs_search_idx = (self._logs_search_idx + 1) % len(self._logs_search_hits)
        self._logs_scroll_to_cur(self._logs_active_txt())
        self._logs_update_status()

    def _logs_find_prev(self):
        if not self._logs_search_hits: self._logs_search(); return
        self._logs_search_idx = (self._logs_search_idx - 1) % len(self._logs_search_hits)
        self._logs_scroll_to_cur(self._logs_active_txt())
        self._logs_update_status()

    def _logs_scroll_to_cur(self, txt: tk.Text):
        if not self._logs_search_hits or self._logs_search_idx < 0: return
        pos = self._logs_search_hits[self._logs_search_idx]
        q = self._logs_search_q
        txt.tag_remove("hl_search_cur", "1.0", tk.END)
        txt.tag_config("hl_search_cur", background="#ff9900", foreground="black")
        txt.tag_add("hl_search_cur", pos, f"{pos}+{len(q)}c")
        txt.see(pos)

    def _logs_export(self):
        if not self._logs_files:
            messagebox.showwarning("No Logs", "No log files to compress.")
            return
        dest = filedialog.asksaveasfilename(
            title="Save Support Logs", defaultextension=".zip",
            initialfile="FOXL_Support_Logs.zip",
            filetypes=[("Zip files", "*.zip")])
        if dest:
            ok, msg = export_logs_to_zip(self._logs_files, dest)
            messagebox.showinfo("Done", f"Saved to:\n{dest}") if ok else messagebox.showerror("Error", msg)

    # ══════════════════════════════════════════════════════════════════════════
    # REGISTRY
    # ══════════════════════════════════════════════════════════════════════════

    def _build_registry(self, parent: tk.Frame):
        parent.columnconfigure(0, weight=2)
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(0, weight=1)

        # Left: table
        left_c = self._card(parent)
        left_c.grid(row=0, column=0, sticky="nsew", padx=(16, 5), pady=16)
        left_c.rowconfigure(1, weight=1)
        left_c.columnconfigure(0, weight=1)

        hdr = tk.Frame(left_c, bg=T.BG_CARD)
        hdr.pack(fill="x", padx=14, pady=(10, 6))
        self._reg_title = _label(hdr, "Registry — NinetyOne Keys", bold=True, bg=T.BG_CARD)
        self._reg_title.pack(side="left")
        _btn(hdr, "⋯  Open Regedit Here", self._reg_open,
             style="ghost_card").pack(side="right", padx=(6, 0))
        _btn(hdr, "↺  Re-scan", self._reg_scan,
             style="ghost_card").pack(side="right")

        tree_f = tk.Frame(left_c, bg=T.BG_CARD)
        tree_f.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        tree_f.rowconfigure(0, weight=1)
        tree_f.columnconfigure(0, weight=1)

        cols = ("key", "value", "data")
        self._reg_tree = ttk.Treeview(tree_f, columns=cols, show="headings", selectmode="browse")
        for col, hdr_txt, w in [("key", "KEY", 240), ("value", "VALUE", 120), ("data", "DATA", 200)]:
            self._reg_tree.heading(col, text=hdr_txt)
            self._reg_tree.column(col, width=w, minwidth=80)
        reg_sb = ttk.Scrollbar(tree_f, orient="vertical", command=self._reg_tree.yview)
        self._reg_tree.configure(yscrollcommand=reg_sb.set)
        self._reg_tree.grid(row=0, column=0, sticky="nsew")
        reg_sb.grid(row=0, column=1, sticky="ns")
        self._reg_tree.bind("<<TreeviewSelect>>", self._reg_on_select)

        # Right: detail
        detail_c = self._card(parent, "Detail")
        detail_c.grid(row=0, column=1, sticky="new", padx=(0, 16), pady=16)

        detail_inner = tk.Frame(detail_c, bg=T.BG_CARD)
        detail_inner.pack(fill="x", padx=14, pady=(4, 12))

        self._reg_detail: dict[str, tk.StringVar] = {}
        for field in ("Hive", "Path", "Value", "Data"):
            row = tk.Frame(detail_inner, bg=T.BG_CARD)
            row.pack(fill="x", pady=3)
            _label(row, field, muted=True, bg=T.BG_CARD, width=6, anchor="w").pack(side="left")
            var = tk.StringVar(value="")
            _label(row, bg=T.BG_CARD, textvariable=var,
                   anchor="w", wraplength=220).pack(side="left", fill="x", expand=True)
            self._reg_detail[field] = var

        btn_row = tk.Frame(detail_c, bg=T.BG_CARD)
        btn_row.pack(fill="x", padx=14, pady=(0, 12))
        _btn(btn_row, "✎  Copy Path", self._reg_copy_path,
             style="ghost_card").pack(side="left", padx=(0, 6))
        _btn(btn_row, "⋯  Open in Regedit", self._reg_open,
             style="ghost_card").pack(side="left")

    def _reg_scan(self):
        self._reg_tree.delete(*self._reg_tree.get_children())
        self._reg_title.config(text="Registry — NinetyOne Keys  Scanning…")
        threading.Thread(target=self._reg_scan_bg, daemon=True).start()

    def _reg_scan_bg(self):
        ok, result = scan_registry_for_ninetyone()
        def upd():
            if ok and result:
                self._reg_title.config(
                    text=f"Registry — NinetyOne Keys  {len(result)} found")
                for entry in result:
                    entry_str = str(entry)
                    if "=" in entry_str:
                        key_part, data = entry_str.split("=", 1)
                        parts = key_part.strip().rsplit("\\", 1)
                        key_path  = parts[0].strip() if len(parts) > 1 else key_part.strip()
                        val_name  = parts[1].strip() if len(parts) > 1 else ""
                    else:
                        key_path, val_name, data = entry_str, "", ""
                    self._reg_tree.insert("", "end", values=(key_path, val_name, data.strip()))
            elif ok:
                self._reg_title.config(text="Registry — NinetyOne Keys  0 found")
            else:
                self._reg_title.config(text=f"Registry — Error: {result}")
        self.root.after(0, upd)

    def _reg_on_select(self, _=None):
        sel = self._reg_tree.selection()
        if not sel: return
        vals = self._reg_tree.item(sel[0])["values"]
        if len(vals) >= 3:
            key, value, data = str(vals[0]), str(vals[1]), str(vals[2])
            parts = key.split("\\")
            self._reg_detail["Hive"].set(parts[0] if parts else "")
            self._reg_detail["Path"].set(key)
            self._reg_detail["Value"].set(value)
            self._reg_detail["Data"].set(data)

    def _reg_copy_path(self):
        path = self._reg_detail["Path"].get()
        if path:
            self.root.clipboard_clear()
            self.root.clipboard_append(path)

    def _reg_open(self):
        ok, msg = open_regedit_at_path()
        if not ok:
            messagebox.showerror("Error", msg)

    # ══════════════════════════════════════════════════════════════════════════
    # SHARED
    # ══════════════════════════════════════════════════════════════════════════

    def _sync_env(self):
        e, v = read_env()
        self._env_var.set(e)
        self._ver_var.set(v)
        self._dash_env_badge.config(text=f"● {e}")

    def _save_env(self, prefix: str, val: str):
        ok, msg = update_env_param(prefix, val)
        if ok:
            if prefix == "VERSION":
                self._ver_var.set(val)
            self._sync_env()
            messagebox.showinfo("Updated", f"{prefix} updated to: {val}")
        else:
            messagebox.showerror("Error", msg)

    def _start_scans(self):
        for path, key, var, dd in [
            (NETWORK_PATH_8,  "v8",    self._v8_var, self._v8_dd),
            (NETWORK_PATH_6,  "v6",    None,          None),
            (LOCAL_TEST_PATH, "local", None,          None),
        ]:
            threading.Thread(target=self._bg_scan, args=(path, key, var, dd),
                             daemon=True).start()

    def _bg_scan(self, path, key, var, dd):
        ok, result = scan_path_sync(path)
        if ok:
            self._versions[key] = result
        def upd():
            if var and dd:
                dd["values"] = self._versions[key]
                var.set(self._versions[key][0] if self._versions[key] else "No versions")
            self._update_counts()
        self.root.after(0, upd)

    def _update_counts(self):
        v8 = len(self._versions["v8"])
        v6 = len(self._versions["v6"])
        lo = len(self._versions["local"])
        self._counts_var.set(f"v8: {v8}  ·  v6: {v6}  ·  local: {lo}")

    def _explore(self, path: str):
        if not open_in_explorer(path):
            messagebox.showwarning("Not Found", path)

    def _xl_open(self):
        ok, err = launch_excel()
        if not ok: messagebox.showerror("Error", err)

    def _xl_close(self):
        ok, err = close_excel()
        if not ok: messagebox.showerror("Error", err)

    def _xl_kill(self):
        if messagebox.askyesno("Force Kill", "Kill Excel? Unsaved work will be lost."):
            ok, err = kill_excel()
            if not ok: messagebox.showerror("Error", err)

    def _open_user_view(self):
        from ui.user_window import UserWindow
        child = tk.Toplevel(self.root)
        UserWindow(child)
        child.grab_set()
