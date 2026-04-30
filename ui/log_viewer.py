import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import threading

from config import LOG_DIR_PATH
from core.log_ops import (
    get_log_files, get_services, extract_service,
    generate_unified_timeline, export_logs_to_zip,
)
import ui.theme as T

_HL = {
    "ERROR": (T.ERROR, "hl_err"),
    "WARN":  (T.WARN,  "hl_warn"),
    "INFO":  (T.INFO,  "hl_info"),
}


def _btn(parent, text, cmd, style="normal", **kw):
    _s = {
        "normal": (T.BG_INPUT, T.TEXT_PRI, T.BORDER),
        "ghost":  (T.BG_CARD,  T.TEXT_SEC, T.BG_INPUT),
    }
    bg, fg, abg = _s.get(style, _s["normal"])
    padx = kw.pop("padx", 8)
    pady = kw.pop("pady", 5)
    return tk.Button(parent, text=text, command=cmd,
                     bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
                     relief="flat", bd=0, padx=padx, pady=pady,
                     cursor="hand2", font=T.F_UI, **kw)


class LogViewer(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("FOXL Support Log Viewer")
        self.geometry("1000x680")
        T.apply(self)

        self.log_files         = get_log_files(LOG_DIR_PATH)
        self._service_filter   = None
        self._service_map: dict = {}
        self._last_entries: list = []
        self._search_hits: list  = []
        self._search_idx         = -1
        self._search_q           = ""

        self._build()
        self.load_data()

    # ── Build ──────────────────────────────────────────────────────────────────

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
        tk.Label(logo_f, text="  FOXL Support Log Viewer", fg=T.TEXT_PRI, bg=T.BG_DARK,
                 font=("Segoe UI", 10, "bold")).pack(side="left")

        right_tb = tk.Frame(topbar, bg=T.BG_DARK)
        right_tb.pack(side="right", padx=12)
        _btn(right_tb, "⊞  Compress & Save Logs for Support",
             self.export_zip).pack(side="right", padx=(6, 0))
        _btn(right_tb, "↺  Refresh", self.refresh,
             style="ghost").pack(side="right")

        # Filter bar
        fb = tk.Frame(self, bg=T.BG_CARD)
        fb.pack(fill="x", padx=16, pady=(12, 0))

        fb_inner = tk.Frame(fb, bg=T.BG_CARD)
        fb_inner.pack(fill="x", padx=14, pady=8)

        tk.Label(fb_inner, text="SERVICE", fg=T.TEXT_MUTED, bg=T.BG_CARD,
                 font=T.F_TINY).pack(side="left")
        self.service_var = tk.StringVar(value="All Services")
        self.service_dd = ttk.Combobox(fb_inner, textvariable=self.service_var,
                                       state="readonly", width=16)
        self.service_dd.pack(side="left", padx=(4, 16))
        self.service_dd.bind("<<ComboboxSelected>>", self._on_service_change)

        tk.Label(fb_inner, text="ORDER", fg=T.TEXT_MUTED, bg=T.BG_CARD,
                 font=T.F_TINY).pack(side="left")
        self.order_var = tk.StringVar(value="Oldest First")
        ttk.Combobox(fb_inner, textvariable=self.order_var,
                     values=["Oldest First", "Newest First"],
                     state="readonly", width=12).pack(side="left", padx=(4, 16))
        self.order_var.trace_add("write", lambda *_: self._reload_unified())

        tk.Label(fb_inner, text="FIND", fg=T.TEXT_MUTED, bg=T.BG_CARD,
                 font=T.F_TINY).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(fb_inner, textvariable=self.search_var, width=24)
        self.search_entry.pack(side="left", padx=(4, 2))
        self.search_entry.bind("<Return>", self._search)
        _btn(fb_inner, "▲", self._find_prev, padx=6, pady=3).pack(side="left", padx=1)
        _btn(fb_inner, "▼", self._find_next, padx=6, pady=3).pack(side="left", padx=1)
        self.bind("<Control-f>", lambda _: self.search_entry.focus_set())

        # Toggle
        self._view_mode = tk.StringVar(value="Timeline")
        self._toggle_btns: dict[str, tk.Button] = {}
        tog_f = tk.Frame(fb_inner, bg=T.BG_CARD)
        tog_f.pack(side="right")
        for mode in ("Timeline", "File View"):
            b = tk.Button(tog_f, text=mode, relief="flat", bd=0,
                          padx=10, pady=4, cursor="hand2", font=T.F_UI,
                          command=lambda m=mode: self._set_view(m))
            b.pack(side="left", padx=1)
            self._toggle_btns[mode] = b

        # Content area
        content = tk.Frame(self, bg=T.BG_CARD)
        content.pack(fill="both", expand=True, padx=16, pady=(6, 0))
        content.rowconfigure(0, weight=1)
        content.columnconfigure(0, weight=1)

        # Timeline frame
        self._tl_frame = tk.Frame(content, bg=T.BG_CARD)
        self._tl_frame.rowconfigure(0, weight=1)
        self._tl_frame.columnconfigure(0, weight=1)

        self.txt_unified = tk.Text(self._tl_frame, bg=T.BG_CARD, fg=T.TEXT_PRI,
                                   font=T.F_MONO, relief="flat", bd=0,
                                   state="disabled", wrap="none",
                                   selectbackground=T.BG_INPUT)
        tl_sby = ttk.Scrollbar(self._tl_frame, orient="vertical",  command=self.txt_unified.yview)
        tl_sbx = ttk.Scrollbar(self._tl_frame, orient="horizontal", command=self.txt_unified.xview)
        self.txt_unified.configure(yscrollcommand=tl_sby.set, xscrollcommand=tl_sbx.set)
        self.txt_unified.grid(row=0, column=0, sticky="nsew")
        tl_sby.grid(row=0, column=1, sticky="ns")
        tl_sbx.grid(row=1, column=0, sticky="ew")

        # File View frame
        self._fv_frame = tk.Frame(content, bg=T.BG_CARD)
        self._fv_frame.rowconfigure(1, weight=1)
        self._fv_frame.columnconfigure(0, weight=1)

        fv_ctrl = tk.Frame(self._fv_frame, bg=T.BG_CARD)
        fv_ctrl.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        tk.Label(fv_ctrl, text="File:", fg=T.TEXT_SEC, bg=T.BG_CARD,
                 font=T.F_UI).pack(side="left", padx=(0, 6))
        self.file_var = tk.StringVar()
        self.file_dd  = ttk.Combobox(fv_ctrl, textvariable=self.file_var,
                                     state="readonly", width=52)
        self.file_dd.pack(side="left")
        self.file_dd.bind("<<ComboboxSelected>>", self.on_file_select)

        self.txt_file = tk.Text(self._fv_frame, bg=T.BG_CARD, fg=T.TEXT_PRI,
                                font=T.F_MONO, relief="flat", bd=0,
                                state="disabled", selectbackground=T.BG_INPUT)
        fv_sb = ttk.Scrollbar(self._fv_frame, orient="vertical", command=self.txt_file.yview)
        self.txt_file.configure(yscrollcommand=fv_sb.set)
        self.txt_file.grid(row=1, column=0, sticky="nsew")
        fv_sb.grid(row=1, column=1, sticky="ns")

        # Status bar
        self.status_var = tk.StringVar(value="Loading…")
        status_bar = tk.Frame(self, bg=T.BG_DARK, height=28)
        status_bar.pack(fill="x", padx=16, pady=(0, 16))
        status_bar.pack_propagate(False)
        tk.Label(status_bar, textvariable=self.status_var,
                 fg=T.TEXT_SEC, bg=T.BG_DARK, font=T.F_SMALL,
                 anchor="w").pack(side="left", padx=12, pady=6)

        # Toolbar (copy/prettify) in status bar
        for lbl, cmd in [("Copy Selected", self.copy_selection),
                         ("Copy All",       self.copy_all),
                         ("Prettify JSON",  self.prettify_json)]:
            _btn(status_bar, lbl, cmd, style="ghost", pady=3).pack(side="right", padx=4)

        self._set_view("Timeline")

    # ── View toggle ────────────────────────────────────────────────────────────

    def _set_view(self, mode: str):
        self._view_mode.set(mode)
        for m, btn in self._toggle_btns.items():
            btn.config(bg=T.ACCENT if m == mode else T.BG_INPUT,
                       fg="#000000" if m == mode else T.TEXT_SEC,
                       font=T.F_BOLD if m == mode else T.F_UI)
        content = self._tl_frame.master
        if mode == "Timeline":
            self._fv_frame.grid_remove()
            self._tl_frame.grid(row=0, column=0, sticky="nsew", padx=14, pady=(4, 0))
        else:
            self._tl_frame.grid_remove()
            self._fv_frame.grid(row=0, column=0, sticky="nsew", padx=14, pady=(4, 0))

    # ── Data loading ───────────────────────────────────────────────────────────

    def load_data(self):
        services = get_services(self.log_files)
        self._service_map = {"All Services": None}
        self._service_map.update({d: r for r, d in services})
        self.service_dd["values"] = list(self._service_map.keys())
        self.service_var.set("All Services")
        self._service_filter = None
        self._reload_file_dropdown()
        self._reload_unified()

    def _get_filtered(self):
        if self._service_filter is None:
            return self.log_files
        return [f for f in self.log_files
                if extract_service(os.path.basename(f)) == self._service_filter]

    def _reload_unified(self):
        self.txt_unified.configure(state="normal")
        self.txt_unified.delete("1.0", tk.END)
        self.txt_unified.insert(tk.END, "Compiling unified timeline…\n")
        self.txt_unified.configure(state="disabled")
        self.status_var.set("Loading…")
        threading.Thread(target=self._load_unified_bg, daemon=True).start()

    def _load_unified_bg(self):
        files   = self._get_filtered()
        reverse = self.order_var.get() == "Newest First"
        entries, errors = generate_unified_timeline(files, reverse=reverse)
        self._last_entries = entries
        data = "".join(t for _, t in entries) if entries else "No logs found."
        if errors:
            data += "\n\n--- Parse errors ---\n" + "\n".join(f"  {n}: {m}" for n, m in errors)

        def upd():
            self.txt_unified.configure(state="normal")
            self.txt_unified.delete("1.0", tk.END)
            self.txt_unified.insert(tk.END, data)
            self._apply_hl(self.txt_unified)
            self.txt_unified.configure(state="disabled")
            self._update_status()

        self.after(0, upd)

    def _reload_file_dropdown(self):
        files = self._get_filtered()
        names = [os.path.basename(f) for f in files]
        self.file_dd["values"] = names
        if names:
            self.file_dd.set(names[0])
            self.on_file_select()
        else:
            self.file_dd.set("")
            self.txt_file.configure(state="normal")
            self.txt_file.delete("1.0", tk.END)
            self.txt_file.configure(state="disabled")

    def refresh(self):
        self.log_files = get_log_files(LOG_DIR_PATH)
        self._clear_search()
        self.load_data()

    def on_file_select(self, _=None):
        name   = self.file_var.get()
        files  = self._get_filtered()
        target = next((f for f in files if os.path.basename(f) == name), None)
        self.txt_file.configure(state="normal")
        self.txt_file.delete("1.0", tk.END)
        if target and os.path.exists(target):
            try:
                with open(target, encoding="utf-8", errors="replace") as f:
                    self.txt_file.insert(tk.END, f.read())
                self._apply_hl(self.txt_file)
                self._update_status()
            except OSError as e:
                self.txt_file.insert(tk.END, str(e))
        self.txt_file.configure(state="disabled")

    # ── Highlighting ───────────────────────────────────────────────────────────

    def _apply_hl(self, txt: tk.Text):
        for kw, (color, tag) in _HL.items():
            txt.tag_config(tag, foreground=color)
            txt.tag_remove(tag, "1.0", tk.END)
            start = "1.0"
            while True:
                pos = txt.search(kw, start, stopindex=tk.END)
                if not pos: break
                end = f"{pos}+{len(kw)}c"
                txt.tag_add(tag, pos, end)
                start = end

    # ── Search ─────────────────────────────────────────────────────────────────

    def _active_txt(self) -> tk.Text:
        return self.txt_unified if self._view_mode.get() == "Timeline" else self.txt_file

    def _search(self, _=None):
        self._clear_search_tags()
        q = self.search_var.get().strip()
        self._search_q = q
        self._search_hits = []
        if not q:
            self._update_status()
            return
        txt   = self._active_txt()
        start = "1.0"
        while True:
            pos = txt.search(q, start, stopindex=tk.END, nocase=True)
            if not pos: break
            self._search_hits.append(pos)
            start = f"{pos}+{len(q)}c"
        txt.tag_config("hl_search", background="#ffff88", foreground="black")
        for pos in self._search_hits:
            txt.tag_add("hl_search", pos, f"{pos}+{len(q)}c")
        self._search_idx = 0 if self._search_hits else -1
        if self._search_hits:
            self._scroll_to_cur(txt)
        self._update_status()

    def _find_next(self):
        if not self._search_hits: self._search(); return
        self._search_idx = (self._search_idx + 1) % len(self._search_hits)
        self._scroll_to_cur(self._active_txt())
        self._update_status()

    def _find_prev(self):
        if not self._search_hits: self._search(); return
        self._search_idx = (self._search_idx - 1) % len(self._search_hits)
        self._scroll_to_cur(self._active_txt())
        self._update_status()

    def _scroll_to_cur(self, txt: tk.Text):
        if not self._search_hits or self._search_idx < 0: return
        pos = self._search_hits[self._search_idx]
        q   = self._search_q
        txt.tag_remove("hl_search_cur", "1.0", tk.END)
        txt.tag_config("hl_search_cur", background="#ff9900", foreground="black")
        txt.tag_add("hl_search_cur", pos, f"{pos}+{len(q)}c")
        txt.see(pos)

    def _clear_search_tags(self):
        for txt in (self.txt_unified, self.txt_file):
            for tag in ("hl_search", "hl_search_cur"):
                txt.tag_remove(tag, "1.0", tk.END)

    def _clear_search(self):
        self._clear_search_tags()
        self._search_hits = []
        self._search_idx  = -1
        self._search_q    = ""

    # ── Status bar ─────────────────────────────────────────────────────────────

    def _update_status(self):
        q = self._search_q
        if q and self._search_hits:
            self.status_var.set(f"{self._search_idx+1} of {len(self._search_hits)} matches for '{q}'")
            return
        if q:
            self.status_var.set(f"No matches for '{q}'")
            return
        entries = self._last_entries
        if entries:
            ts = [e[0] for e in entries]
            self.status_var.set(f"{len(entries)} entries  ·  {min(ts)}  →  {max(ts)}")
        else:
            self.status_var.set("No entries")

    # ── Filter handlers ────────────────────────────────────────────────────────

    def _on_service_change(self, _=None):
        self._service_filter = self._service_map.get(self.service_var.get())
        self._clear_search()
        self._reload_unified()
        self._reload_file_dropdown()

    # ── Toolbar actions ────────────────────────────────────────────────────────

    def copy_selection(self):
        txt = self._active_txt()
        try:
            self.clipboard_clear()
            self.clipboard_append(txt.get(tk.SEL_FIRST, tk.SEL_LAST))
        except tk.TclError:
            messagebox.showwarning("Warning", "No text selected.")

    def copy_all(self):
        txt = self._active_txt()
        self.clipboard_clear()
        self.clipboard_append(txt.get("1.0", tk.END))

    def prettify_json(self):
        txt = self._active_txt()
        try:
            s, e   = txt.index(tk.SEL_FIRST), txt.index(tk.SEL_LAST)
            parsed = json.loads(txt.get(s, e))
            txt.configure(state="normal")
            txt.delete(s, e)
            txt.insert(s, json.dumps(parsed, indent=4))
            txt.configure(state="disabled")
        except tk.TclError:
            messagebox.showwarning("Warning", "Highlight the JSON text first.")
        except json.JSONDecodeError:
            messagebox.showerror("Error", "Selected text is not valid JSON.")

    def export_zip(self):
        if not self.log_files:
            messagebox.showwarning("Warning", "No log files to compress.")
            return
        dest = filedialog.asksaveasfilename(
            title="Save Support Logs", defaultextension=".zip",
            initialfile="FOXL_Support_Logs.zip",
            filetypes=[("Zip files", "*.zip")])
        if dest:
            ok, msg = export_logs_to_zip(self.log_files, dest)
            messagebox.showinfo("Done", f"Saved to:\n{dest}") if ok else messagebox.showerror("Error", msg)
