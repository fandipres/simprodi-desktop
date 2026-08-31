"""
Tool UI: Cek Berita Acara Perwalian.
Placeholder - fitur validasi isi PDF Berita Acara (tanggal, tanda tangan,
daftar mahasiswa, dll) belum diimplementasikan.
"""

from tkinter import ttk

from tools import theme

NAME = "Cek Berita Acara Perwalian"
LABEL = "Cek Berita Acara"
ICON = "\U0001F4CB"  # 📋
DESCRIPTION = "Cek otomatis apakah PDF Berita Acara Perwalian sudah diisi lengkap - tanggal, tanda tangan, daftar mahasiswa, dan lainnya."


def build_frame(parent):
    frame = ttk.Frame(parent, padding=24, style="Card.TFrame")

    ttk.Label(frame, text=NAME, font=(theme.FONT_FAMILY, 15, "bold")).grid(
        row=0, column=0, sticky="w"
    )

    ttk.Label(frame, text=DESCRIPTION, foreground=theme.TEXT_MUTED).grid(
        row=1, column=0, sticky="w", pady=(4, 28)
    )

    badge = ttk.Frame(frame, style="Card.TFrame")
    badge.grid(row=2, column=0, sticky="w")
    ttk.Label(
        badge,
        text="  🚧  Coming Soon  ",
        font=(theme.FONT_FAMILY, 11, "bold"),
        foreground=theme.TEXT_MUTED,
        background="#f1f5f9",
        padding=(10, 8),
    ).pack()

    return frame
