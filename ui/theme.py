import tkinter as tk
from tkinter import ttk

# ── Palette ───────────────────────────────────────────────────────────────────
BG_DARK    = "#1a1d22"
BG_MAIN    = "#1e2127"
BG_CARD    = "#252930"
BG_INPUT   = "#2d3139"
ACCENT     = "#00c896"
ACCENT_DIM = "#009e78"
TEXT_PRI   = "#e8e8e8"
TEXT_SEC   = "#8b8fa8"
TEXT_MUTED = "#4e5364"
BORDER     = "#31363f"
ERROR      = "#e05252"
WARN       = "#c8960c"
INFO       = "#5b9bd5"

# ── Fonts ─────────────────────────────────────────────────────────────────────
F_UI     = ("Segoe UI", 9)
F_BOLD   = ("Segoe UI", 9, "bold")
F_SMALL  = ("Segoe UI", 8)
F_TINY   = ("Segoe UI", 7)
F_LARGE  = ("Segoe UI", 11, "bold")
F_MONO   = ("Consolas", 9)


def apply(root: tk.Misc) -> None:
    root.configure(bg=BG_MAIN)
    s = ttk.Style(root)
    s.theme_use("clam")

    s.configure(".", background=BG_MAIN, foreground=TEXT_PRI, font=F_UI,
                borderwidth=0, relief="flat")

    # Frames
    s.configure("TFrame",    background=BG_MAIN)
    s.configure("Card.TFrame", background=BG_CARD)

    # Labels
    s.configure("TLabel",       background=BG_MAIN, foreground=TEXT_PRI)
    s.configure("Muted.TLabel", background=BG_MAIN, foreground=TEXT_SEC, font=F_SMALL)

    # Treeview
    s.configure("Treeview",
                background=BG_CARD, foreground=TEXT_PRI,
                fieldbackground=BG_CARD, borderwidth=0,
                rowheight=28, font=F_UI)
    s.configure("Treeview.Heading",
                background=BG_DARK, foreground=TEXT_MUTED,
                font=F_TINY, relief="flat", borderwidth=0)
    s.map("Treeview",
          background=[("selected", "#2d5a4a")],
          foreground=[("selected", ACCENT)])
    s.map("Treeview.Heading", background=[("active", BG_DARK)])

    # Entry
    s.configure("TEntry",
                fieldbackground=BG_INPUT, foreground=TEXT_PRI,
                insertcolor=TEXT_PRI, bordercolor=BORDER,
                lightcolor=BG_INPUT, darkcolor=BG_INPUT,
                relief="flat", padding=(8, 6))
    s.map("TEntry",
          fieldbackground=[("disabled", BG_DARK)],
          foreground=[("disabled", TEXT_SEC)])

    # Combobox
    s.configure("TCombobox",
                fieldbackground=BG_INPUT, background=BG_INPUT,
                foreground=TEXT_PRI, arrowcolor=TEXT_SEC,
                bordercolor=BG_INPUT, lightcolor=BG_INPUT, darkcolor=BG_INPUT,
                relief="flat", padding=(8, 6))
    s.map("TCombobox",
          fieldbackground=[("readonly", BG_INPUT), ("disabled", BG_DARK)],
          foreground=[("disabled", TEXT_SEC)],
          arrowcolor=[("disabled", TEXT_MUTED)])

    # Checkbutton
    s.configure("TCheckbutton", background=BG_CARD, foreground=TEXT_PRI)
    s.map("TCheckbutton",
          background=[("active", BG_CARD)],
          foreground=[("active", TEXT_PRI)])

    # Progressbar
    s.configure("TProgressbar",
                background=ACCENT, troughcolor=BG_INPUT,
                borderwidth=0, thickness=4)

    # Scrollbar
    s.configure("TScrollbar",
                background=BG_INPUT, troughcolor=BG_CARD,
                arrowcolor=TEXT_MUTED, borderwidth=0, relief="flat")
    s.map("TScrollbar", background=[("active", BORDER)])

    # Separator
    s.configure("TSeparator", background=BORDER)
