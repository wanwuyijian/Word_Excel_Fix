"""界面主题：简洁专业蓝白风。"""

from __future__ import annotations

from tkinter import font as tkfont
from tkinter import ttk

# ---- 配色 ----
BG = "#F5F7FA"          # 窗口底色
CARD = "#FFFFFF"        # 卡片
BORDER = "#DCE3EC"      # 分隔线
TEXT = "#1F2937"        # 主文字
TEXT_MUTED = "#6B7280"  # 次要文字
PRIMARY = "#2563EB"     # 主蓝
PRIMARY_DARK = "#1D4ED8"
PRIMARY_LIGHT = "#DBE7FE"
SUCCESS = "#10B981"
DANGER = "#DC2626"
WARN = "#D97706"
ROW_ALT = "#F9FBFD"
WARN_BG = "#FFF7E6"
WARN_BORDER = "#F5C97B"
DROP_BG = "#F2F7FF"
DROP_BORDER = "#C3D7F7"


def pick_font(candidates, fallback="Microsoft YaHei UI") -> str:
    """挑选系统里存在的字体。"""
    try:
        available = set(tkfont.families())
    except Exception:  # pragma: no cover
        return fallback
    for name in candidates:
        if name in available:
            return name
    return fallback


def ui_font(size: int = 10, weight: str = "normal") -> tuple:
    family = pick_font(["Microsoft YaHei UI", "微软雅黑", "Microsoft YaHei", "Segoe UI"])
    return (family, size, weight)


def mono_font(size: int = 9) -> tuple:
    return (pick_font(["Consolas", "Cascadia Mono", "Courier New"]), size, "normal")


def apply_theme(root) -> ttk.Style:
    """把 ttk 默认主题改造成蓝白风。"""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:  # pragma: no cover
        pass

    base = ui_font(10)
    root.option_add("*Font", base)

    style.configure(".", font=base, background=BG, foreground=TEXT)
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD)
    style.configure("TLabel", background=BG, foreground=TEXT)
    style.configure("Card.TLabel", background=CARD, foreground=TEXT)
    style.configure("Muted.TLabel", background=BG, foreground=TEXT_MUTED,
                    font=ui_font(9))
    style.configure("CardMuted.TLabel", background=CARD, foreground=TEXT_MUTED,
                    font=ui_font(9))
    style.configure("H1.TLabel", background=BG, foreground=TEXT, font=ui_font(15, "bold"))
    style.configure("H2.TLabel", background=CARD, foreground=TEXT, font=ui_font(11, "bold"))
    style.configure("Primary.TButton", font=ui_font(11, "bold"),
                    background=PRIMARY, foreground="#FFFFFF",
                    borderwidth=0, focusthickness=0, padding=(18, 9))
    style.map("Primary.TButton",
              background=[("active", PRIMARY_DARK), ("disabled", "#93B4F0")],
              foreground=[("disabled", "#F1F5F9")])
    style.configure("TButton", padding=(12, 6), background="#FFFFFF",
                    foreground=TEXT, bordercolor=BORDER, relief="flat")
    style.map("TButton", background=[("active", PRIMARY_LIGHT), ("disabled", "#EEF1F5")])
    style.configure("Link.TButton", padding=(6, 2), background=BG,
                    foreground=PRIMARY, borderwidth=0, relief="flat")
    style.map("Link.TButton", background=[("active", BG)], foreground=[("active", PRIMARY_DARK)])

    style.configure("TEntry", fieldbackground="#FFFFFF", bordercolor=BORDER,
                    lightcolor=BORDER, darkcolor=BORDER, padding=4)
    style.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF",
                    bordercolor=BORDER, arrowcolor=PRIMARY, padding=4)
    style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")])
    style.configure("TCheckbutton", background=CARD, foreground=TEXT)
    style.map("TCheckbutton", background=[("active", CARD)])
    style.configure("TRadiobutton", background=CARD, foreground=TEXT)
    style.configure("TPanedwindow", background=BG)
    style.configure("TSeparator", background=BORDER)

    style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF",
                    foreground=TEXT, rowheight=26, borderwidth=0)
    style.configure("Treeview.Heading", font=ui_font(9, "bold"),
                    background="#F1F5F9", foreground=TEXT_MUTED, relief="flat",
                    padding=(6, 6))
    style.map("Treeview.Heading", background=[("active", "#E6EDF7")])
    style.map("Treeview", background=[("selected", PRIMARY_LIGHT)],
              foreground=[("selected", TEXT)])

    style.configure("Blue.Horizontal.TProgressbar", troughcolor="#E8EDF5",
                    background=PRIMARY, bordercolor="#E8EDF5",
                    lightcolor=PRIMARY, darkcolor=PRIMARY, thickness=10)

    style.configure("Card.TLabelframe", background=CARD, bordercolor=BORDER)
    style.configure("Card.TLabelframe.Label", background=CARD, foreground=TEXT,
                    font=ui_font(10, "bold"))

    # ---- v1.1 新增：分组卡片、徽标、警示条 ----
    style.configure("Group.TFrame", background="#F8FAFD")
    style.configure("Group.TLabel", background="#F8FAFD", foreground=TEXT)
    style.configure("GroupMuted.TLabel", background="#F8FAFD", foreground=TEXT_MUTED,
                    font=ui_font(9))
    style.configure("GroupTitle.TLabel", background="#F8FAFD", foreground=TEXT,
                    font=ui_font(9, "bold"))
    style.configure("Group.TCheckbutton", background="#F8FAFD", foreground=TEXT)
    style.map("Group.TCheckbutton", background=[("active", "#F8FAFD")])
    style.configure("Card.TCheckbutton", background=CARD, foreground=TEXT)
    style.map("Card.TCheckbutton", background=[("active", CARD)])

    style.configure("Warning.TFrame", background=WARN_BG)
    style.configure("Warning.TLabel", background=WARN_BG, foreground="#92400E",
                    font=ui_font(9))
    style.configure("WarningTitle.TLabel", background=WARN_BG, foreground="#92400E",
                    font=ui_font(9, "bold"))

    style.configure("Badge.TLabel", background=PRIMARY_LIGHT, foreground=PRIMARY_DARK,
                    font=ui_font(8, "bold"), padding=(6, 2))
    style.configure("BadgeGray.TLabel", background="#EEF2F7", foreground=TEXT_MUTED,
                    font=ui_font(8, "bold"), padding=(6, 2))

    style.configure("Ghost.TButton", padding=(10, 6), background="#FFFFFF",
                    foreground=TEXT, bordercolor=BORDER, relief="flat")
    style.map("Ghost.TButton",
              background=[("active", "#F1F5F9"), ("disabled", "#F7F8FA")],
              foreground=[("disabled", "#A8B0BD")])

    style.configure("Soft.Horizontal.TProgressbar", troughcolor="#E8EDF5",
                    background=PRIMARY, bordercolor="#E8EDF5",
                    lightcolor=PRIMARY, darkcolor=PRIMARY, thickness=8)
    return style
