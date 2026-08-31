"""
Palet warna & helper styling bersama, dipakai semua tool supaya tampilan
konsisten (flat/modern) di seluruh aplikasi. Tidak ada logika bisnis di
sini - murni tampilan.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk

FONT_FAMILY = "Segoe UI"

BG = "#f1f5f9"            # latar area konten (di luar card putih)
CARD_BG = "#ffffff"       # latar card/frame tiap tool
SIDEBAR_BG = "#ffffff"
BORDER = "#e2e8f0"
BORDER_STRONG = "#cbd5e1"

TEXT = "#0f172a"
TEXT_MUTED = "#64748b"
TEXT_FAINT = "#94a3b8"

ACCENT = "#2563eb"
ACCENT_HOVER = "#1d4ed8"
ACCENT_SOFT = "#eff6ff"
ACCENT_SOFT_BORDER = "#bfdbfe"

DANGER = "#dc2626"


def resource_path(*parts):
    """Path ke file resource, benar baik dijalankan dari source maupun exe PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, *parts)


def apply_app_style(root):
    """Setup ttk.Style global - dipanggil sekali di app.py. 'clam' dipakai
    karena (beda dari tema default Windows) benar-benar menghormati semua
    warna kustom yang di-set lewat style.configure/map, jadi tampilannya
    tidak lagi mengikuti chrome Windows lawas."""
    style = ttk.Style(root)
    style.theme_use("clam")

    root.option_add("*Font", (FONT_FAMILY, 10))

    # Default TFrame/TLabel = putih (CARD_BG): tiap tool membangun 1 frame
    # akar bergaya "Card.TFrame" (putih + border) di atas kanvas abu-abu
    # (BG) punya app.py; semua frame/label BERSARANG di dalamnya otomatis
    # ikut putih tanpa perlu di-style manual satu-satu.
    style.configure(".", background=CARD_BG, foreground=TEXT, font=(FONT_FAMILY, 10))
    style.configure("TFrame", background=CARD_BG)
    style.configure("TLabel", background=CARD_BG, foreground=TEXT)
    style.configure("TSeparator", background=BORDER)

    style.configure(
        "Card.TFrame",
        background=CARD_BG,
        relief="solid",
        borderwidth=1,
        bordercolor=BORDER,
    )
    style.configure("Card.TLabel", background=CARD_BG, foreground=TEXT)

    style.configure(
        "TEntry",
        fieldbackground="#ffffff",
        foreground=TEXT,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        insertcolor=TEXT,
        padding=6,
    )
    style.map(
        "TEntry",
        bordercolor=[("focus", ACCENT)],
        lightcolor=[("focus", ACCENT)],
        darkcolor=[("focus", ACCENT)],
    )

    # Tombol default (flat, aksen biru) - dipakai semua tombol "Proses"/aksi utama.
    style.configure(
        "TButton",
        background=ACCENT,
        foreground="#ffffff",
        borderwidth=0,
        focusthickness=0,
        padding=(14, 8),
        font=(FONT_FAMILY, 10, "bold"),
    )
    style.map(
        "TButton",
        background=[("disabled", "#c7d2fe"), ("pressed", ACCENT_HOVER), ("active", ACCENT_HOVER)],
        foreground=[("disabled", "#eef2ff")],
    )

    # Tombol sekunder (mis. "Pilih...", "Hapus Terpilih") - abu netral, tidak
    # bersaing visual dengan tombol aksi utama.
    style.configure(
        "Secondary.TButton",
        background="#e2e8f0",
        foreground=TEXT,
        borderwidth=0,
        focusthickness=0,
        padding=(12, 6),
        font=(FONT_FAMILY, 9),
    )
    style.map(
        "Secondary.TButton",
        background=[("pressed", BORDER_STRONG), ("active", "#cbd5e1")],
    )

    style.configure(
        "TScrollbar",
        background="#e2e8f0",
        troughcolor=BG,
        bordercolor=BG,
        arrowcolor=TEXT_MUTED,
        gripcount=0,
    )
    style.map("TScrollbar", background=[("active", BORDER_STRONG)])

    # Sidebar
    style.configure("Sidebar.TFrame", background=SIDEBAR_BG)
    style.configure("SidebarTitle.TLabel", background=SIDEBAR_BG, foreground=TEXT, font=(FONT_FAMILY, 15, "bold"))
    style.configure("SidebarSubtitle.TLabel", background=SIDEBAR_BG, foreground=TEXT_MUTED, font=(FONT_FAMILY, 8))
    style.configure(
        "SidebarGroup.TLabel",
        background=SIDEBAR_BG,
        foreground=TEXT_FAINT,
        font=(FONT_FAMILY, 8, "bold"),
    )

    style.configure(
        "Sidebar.TButton",
        background=SIDEBAR_BG,
        foreground=TEXT,
        borderwidth=0,
        focusthickness=0,
        anchor="w",
        padding=(14, 9),
        font=(FONT_FAMILY, 10),
    )
    style.map(
        "Sidebar.TButton",
        background=[("active", "#f1f5f9")],
        foreground=[("active", TEXT)],
    )

    style.configure(
        "SidebarActive.TButton",
        background=ACCENT_SOFT,
        foreground=ACCENT,
        borderwidth=0,
        focusthickness=0,
        anchor="w",
        padding=(14, 9),
        font=(FONT_FAMILY, 10, "bold"),
    )
    style.map(
        "SidebarActive.TButton",
        background=[("active", ACCENT_SOFT)],
        foreground=[("active", ACCENT)],
    )

    return style


def style_text_widget(widget, focus_border=True):
    """Terapkan tampilan flat modern ke widget tk.Text/tk.Listbox mentah
    (bukan ttk - tk.Text/Listbox tidak ikut ttk.Style, jadi di-set manual)."""
    widget.configure(
        relief="flat",
        borderwidth=0,
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT if focus_border else BORDER,
        background="#ffffff",
        foreground=TEXT,
        insertbackground=TEXT,
        font=(FONT_FAMILY, 10),
        padx=8,
        pady=6,
    )


def style_listbox(widget):
    """Sama seperti style_text_widget, tapi untuk tk.Listbox - opsi yang
    valid beda dari tk.Text (tidak ada padx/pady/insertbackground)."""
    widget.configure(
        relief="flat",
        borderwidth=0,
        highlightthickness=1,
        highlightbackground=BORDER,
        highlightcolor=ACCENT,
        background="#ffffff",
        foreground=TEXT,
        font=(FONT_FAMILY, 10),
        selectbackground=ACCENT_SOFT,
        selectforeground=ACCENT,
        activestyle="none",
    )


def card_frame(parent, **kwargs):
    """ttk.Frame dengan style card (latar putih) + padding default 24."""
    kwargs.setdefault("style", "Card.TFrame")
    kwargs.setdefault("padding", 24)
    return ttk.Frame(parent, **kwargs)
