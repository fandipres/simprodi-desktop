"""
Komponen UI kecil yang dipakai bersama oleh semua tool.
"""

import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from tools import theme


def app_root():
    """Folder tempat "SIMPRODI Desktop.exe" berada di disk (atau folder proyek
    kalau dijalankan dari source) - BEDA dari sys._MEIPASS yang dipakai
    theme.resource_path untuk aset yang dibundel (folder sementara PyInstaller,
    bukan lokasi exe sesungguhnya), supaya folder data/ default konsisten ada
    di sebelah exe/proyeknya, bukan di folder temp yang hilang tiap sesi."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def data_dir(*parts):
    """Path ke subfolder data/ default (dibuat kalau belum ada), dipakai
    sebagai folder awal dialog pilih/simpan file tiap tool."""
    path = os.path.join(app_root(), "data", *parts)
    os.makedirs(path, exist_ok=True)
    return path


def build_folder_creator_frame(
    parent, name, description, default_prefix, names_label, create_fn,
    default_output_dir=None,
):
    """
    Frame tool generik: isi daftar nama (satu nama per baris) di kotak
    teks, tentukan prefix & folder output, lalu buat satu folder per baris
    nama lewat create_fn(names, output_dir, prefix).

    Dipakai oleh tool-tool "Buat Folder ..." yang polanya identik, hanya
    beda nama, prefix default, label daftar nama, dan fungsi core-nya.
    """
    frame = ttk.Frame(parent, padding=24, style="Card.TFrame")

    ttk.Label(frame, text=name, font=(theme.FONT_FAMILY, 15, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w"
    )

    ttk.Label(frame, text=description, foreground=theme.TEXT_MUTED).grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(4, 20)
    )

    prefix_var = tk.StringVar(value=default_prefix)
    output_var = tk.StringVar(value=default_output_dir or "")

    # Prefix biasanya pendek, jadi field-nya sengaja tidak dilebarkan
    # mengikuti kolom (sticky="w" saja, bukan "we").
    ttk.Label(frame, text="Prefix Nama Folder").grid(row=2, column=0, sticky="w", pady=6)
    ttk.Entry(frame, textvariable=prefix_var, width=28).grid(
        row=2, column=1, sticky="w", padx=10, pady=6
    )

    def pick_output():
        path = filedialog.askdirectory(
            title="Pilih/tentukan folder output",
            initialdir=output_var.get() or default_output_dir or None,
        )
        if path:
            output_var.set(path)

    ttk.Label(frame, text="Folder Output").grid(row=3, column=0, sticky="w", pady=6)
    ttk.Entry(frame, textvariable=output_var, width=52).grid(
        row=3, column=1, sticky="we", padx=10, pady=6
    )
    ttk.Button(frame, text="Pilih...", style="Secondary.TButton", command=pick_output).grid(
        row=3, column=2, pady=6
    )

    ttk.Label(frame, text=names_label).grid(row=4, column=0, sticky="nw", pady=(14, 6))
    names_text = tk.Text(frame, height=8, width=56, wrap="word")
    theme.style_text_widget(names_text)
    names_text.grid(row=4, column=1, columnspan=2, sticky="nsew", padx=10, pady=(14, 6))
    frame.rowconfigure(4, weight=1)

    frame.columnconfigure(1, weight=1)

    process_btn = ttk.Button(frame, text="Proses")
    process_btn.grid(row=5, column=0, columnspan=3, pady=(18, 8), sticky="w")

    result_row = ttk.Frame(frame)
    result_row.grid(row=6, column=0, columnspan=3, sticky="w")
    open_folder_btn = ttk.Button(result_row, text="Buka Folder Hasil", style="Secondary.TButton")

    ttk.Label(frame, text="Log", foreground=theme.TEXT_MUTED, font=(theme.FONT_FAMILY, 9, "bold")).grid(
        row=7, column=0, sticky="w", pady=(14, 4)
    )
    log_text = tk.Text(frame, height=8, width=90, state="disabled", wrap="word")
    theme.style_text_widget(log_text, focus_border=False)
    log_text.grid(row=8, column=0, columnspan=3, sticky="nsew")
    frame.rowconfigure(8, weight=1)

    def log(msg):
        log_text.configure(state="normal")
        log_text.insert("end", msg + "\n")
        log_text.see("end")
        log_text.configure(state="disabled")

    result_queue = queue.Queue()

    def run_worker(names, output_dir, prefix):
        try:
            created, existing = create_fn(names, output_dir, prefix)
            result_queue.put(("ok", created, existing, output_dir))
        except Exception as e:
            result_queue.put(("error", str(e)))

    def poll_queue():
        try:
            item = result_queue.get_nowait()
        except queue.Empty:
            frame.after(100, poll_queue)
            return

        process_btn.configure(state="normal")
        if item[0] == "ok":
            _, created, existing, output_dir = item
            log(f"Berhasil. Folder dibuat: {len(created)}, sudah ada sebelumnya: {len(existing)}")
            for folder_name in created:
                log(f"  + {folder_name}")
            for folder_name in existing:
                log(f"  = {folder_name} (sudah ada, dilewati)")
            log(f"Lokasi: {output_dir}")
            open_folder_btn.grid(row=0, column=0)
        else:
            log(f"Gagal: {item[1]}")
            messagebox.showerror("Gagal memproses", item[1])

    def start_process():
        output_dir = output_var.get().strip()
        prefix = prefix_var.get()
        names = [line.strip() for line in names_text.get("1.0", "end").splitlines() if line.strip()]

        if not names:
            messagebox.showwarning("Belum lengkap", "Isi daftar nama terlebih dahulu (satu nama per baris).")
            return
        if not output_dir:
            messagebox.showwarning("Belum lengkap", "Tentukan folder output terlebih dahulu.")
            return

        open_folder_btn.grid_remove()
        process_btn.configure(state="disabled")
        log("Memproses...")
        threading.Thread(
            target=run_worker, args=(names, output_dir, prefix), daemon=True
        ).start()
        frame.after(100, poll_queue)

    process_btn.configure(command=start_process)
    open_folder_btn.configure(command=lambda: os.startfile(output_var.get()))

    return frame
