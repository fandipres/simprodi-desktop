"""
SIMPRODI Desktop (alat bantu administrasi program studi) - jendela utama.

Menu di sisi kiri dikelompokkan per kategori. Untuk menambah tool baru:
1. Buat file baru di folder tools/, mis. tools/tool_baru.py
2. Di file itu, sediakan:
     NAME = "Nama Tool"       (judul di halaman konten)
     LABEL = "Label Sidebar"  (teks tombol di sidebar, di bawah kategorinya)
     ICON = "\U0001f4c1"      (opsional, 1 emoji - default "•" kalau tidak diisi)
     def build_frame(parent) -> tk.Widget: ...
3. Import modulnya dan tambahkan ke grup yang sesuai di TOOL_GROUPS di
   bawah (atau buat kategori baru).
"""

import tkinter as tk
from tkinter import ttk

from tools import (
    bandingkan_dokumen,
    cek_berita_acara,
    eskalasi_sp,
    folder_dosen_penasihat_akademik,
    folder_mata_kuliah,
    pdf_peserta,
    theme,
)

TOOL_GROUPS = [
    ("Dosen Wali", [folder_dosen_penasihat_akademik, cek_berita_acara, eskalasi_sp]),
    ("Mata Kuliah", [folder_mata_kuliah]),
    ("Sertifikasi", [pdf_peserta]),
    ("Dokumen", [bandingkan_dokumen]),
]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SIMPRODI Desktop - Alat Bantu Administrasi Program Studi")
        self.geometry("1180x740")
        self.minsize(1000, 620)
        self.configure(background=theme.BG)
        self._set_icon()

        self._build_layout()
        self._show_tool(TOOL_GROUPS[0][1][0])

    def _set_icon(self):
        try:
            self.iconbitmap(theme.resource_path("assets", "icon.ico"))
        except tk.TclError:
            pass  # ikon opsional - jangan sampai app gagal jalan gara-gara ini

    def _build_layout(self):
        theme.apply_app_style(self)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        sidebar = ttk.Frame(container, width=250, style="Sidebar.TFrame")
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        header = ttk.Frame(sidebar, style="Sidebar.TFrame", padding=(20, 22, 20, 16))
        header.pack(fill="x")
        ttk.Label(
            header, text="SIMPRODI Desktop", style="SidebarTitle.TLabel", wraplength=210, justify="left"
        ).pack(anchor="w")
        ttk.Label(
            header,
            text="Alat Bantu Administrasi Program Studi",
            style="SidebarSubtitle.TLabel",
            wraplength=210,
            justify="left",
        ).pack(anchor="w", pady=(3, 0))

        ttk.Separator(sidebar).pack(fill="x", padx=20, pady=(2, 10))

        nav = ttk.Frame(sidebar, style="Sidebar.TFrame", padding=(12, 0))
        nav.pack(fill="both", expand=True)

        self._tool_buttons = {}

        for i, (group_name, modules) in enumerate(TOOL_GROUPS):
            ttk.Label(
                nav, text=group_name.upper(), style="SidebarGroup.TLabel"
            ).pack(anchor="w", padx=8, pady=(14 if i else 4, 4))

            for module in modules:
                icon = getattr(module, "ICON", "•")
                btn = ttk.Button(
                    nav,
                    text=f"{icon}  {module.LABEL}",
                    style="Sidebar.TButton",
                    command=lambda m=module: self._show_tool(m),
                    cursor="hand2",
                )
                btn.pack(fill="x", pady=2)
                self._tool_buttons[module] = btn

        content_outer = tk.Frame(container, background=theme.BG)
        content_outer.pack(side="left", fill="both", expand=True)

        self.content = tk.Frame(content_outer, background=theme.BG)
        self.content.pack(fill="both", expand=True, padx=28, pady=24)

    def _show_tool(self, module):
        for m, btn in self._tool_buttons.items():
            btn.configure(style="SidebarActive.TButton" if m is module else "Sidebar.TButton")

        for child in self.content.winfo_children():
            child.destroy()
        frame = module.build_frame(self.content)
        frame.pack(fill="both", expand=True)


if __name__ == "__main__":
    App().mainloop()
