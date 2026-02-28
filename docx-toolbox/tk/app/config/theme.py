"""ttk.Style 主题配置 — 浅色主题"""

from tkinter import ttk

# 色板
PRIMARY = "#3B82F6"
BG = "#F8FAFC"
CARD = "#FFFFFF"
TEXT = "#111827"
BORDER = "#D1D5DB"
SUCCESS = "#16A34A"
ERROR = "#DC2626"
MUTED = "#6B7280"


def apply_theme(style: ttk.Style) -> None:
    """应用统一浅色主题"""
    style.theme_use("clam")

    style.configure(".", background=BG, foreground=TEXT, bordercolor=BORDER,
                     focuscolor=PRIMARY, font=("system", 12))

    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD, relief="solid", borderwidth=1)

    style.configure("TLabel", background=BG, foreground=TEXT, font=("system", 12))
    style.configure("Title.TLabel", font=("system", 14, "bold"))
    style.configure("Muted.TLabel", foreground=MUTED, font=("system", 11))

    style.configure("TButton", font=("system", 12), padding=(12, 6))
    style.configure("Primary.TButton", background=PRIMARY, foreground="white",
                     font=("system", 12, "bold"))
    style.map("Primary.TButton",
              background=[("active", "#2563EB"), ("disabled", BORDER)])

    style.configure("Danger.TButton", background=ERROR, foreground="white")
    style.map("Danger.TButton",
              background=[("active", "#B91C1C"), ("disabled", BORDER)])

    style.configure("TEntry", fieldbackground=CARD, bordercolor=BORDER, padding=4)
    style.configure("TCheckbutton", background=BG, font=("system", 12))
    style.configure("TSpinbox", fieldbackground=CARD, padding=4)
    style.configure("TCombobox", fieldbackground=CARD, padding=4)

    style.configure("TNotebook", background=BG, tabmargins=(4, 4, 4, 0))
    style.configure("TNotebook.Tab", background=CARD, padding=(14, 6),
                     font=("system", 12))
    style.map("TNotebook.Tab",
              background=[("selected", PRIMARY)],
              foreground=[("selected", "white")])

    style.configure("Horizontal.TProgressbar", troughcolor=BORDER,
                     background=PRIMARY, thickness=8)

    style.configure("Treeview", background=CARD, fieldbackground=CARD,
                     foreground=TEXT, rowheight=28, font=("system", 11))
    style.configure("Treeview.Heading", font=("system", 11, "bold"),
                     background=BG, foreground=TEXT)
