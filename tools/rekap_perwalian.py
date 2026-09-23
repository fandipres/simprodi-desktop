"""
Tool UI: Rekap Perwalian.
Wrapper Tkinter di sekitar fungsi load_perwalian_roster()/
generate_from_folders()/export_excel() dari core/rekap_perwalian.py.

Mulai dari roster perwalian (semua anak wali tiap dosen), lalu nilainya
dicari di seluruh folder data nilai yang dipilih (boleh lebih dari 1
folder/semester sekaligus) kalau ada.
"""

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import rekap_perwalian as core
from tools import theme
from tools.common import data_dir

NAME = "Rekap Perwalian"
LABEL = "Rekap Perwalian"
ICON = "\U0001F465"  # 👥
DESCRIPTION = (
    "Bikin rekap nilai per dosen wali mulai dari roster perwalian - nilai "
    "tiap anak wali dicari di seluruh folder data nilai yang dipilih."
)


def _build_folder_list_section(frame, row, label_text, initialdir, height=3):
    """Baris daftar folder (bisa lebih dari 1) - Listbox + Tambah/Hapus/
    Bersihkan, mirip _build_file_list_section di tools/eskalasi_sp.py tapi
    untuk folder (askdirectory dipanggil berulang, sekali per folder)."""
    ttk.Label(frame, text=label_text).grid(row=row, column=0, sticky="nw", pady=6)

    list_frame = ttk.Frame(frame)
    list_frame.grid(row=row, column=1, columnspan=2, sticky="nsew", padx=10, pady=6)
    list_frame.columnconfigure(0, weight=1)
    list_frame.rowconfigure(0, weight=1)
    frame.rowconfigure(row, weight=1)

    listbox = tk.Listbox(list_frame, height=height, selectmode="extended")
    theme.style_listbox(listbox)
    listbox.grid(row=0, column=0, sticky="nsew")
    scroll = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
    scroll.grid(row=0, column=1, sticky="ns")
    listbox.configure(yscrollcommand=scroll.set)

    folder_paths = []

    def refresh():
        listbox.delete(0, "end")
        for path in folder_paths:
            listbox.insert("end", os.path.basename(path.rstrip("/\\")) or path)

    def add_folder():
        path = filedialog.askdirectory(title=f"Pilih {label_text}", initialdir=initialdir)
        if path and path not in folder_paths:
            folder_paths.append(path)
            refresh()

    def remove_selected():
        for idx in reversed(listbox.curselection()):
            del folder_paths[idx]
        refresh()

    def clear_all():
        folder_paths.clear()
        refresh()

    buttons_row = ttk.Frame(frame)
    buttons_row.grid(row=row + 1, column=1, columnspan=2, sticky="w", padx=10, pady=(0, 6))
    ttk.Button(buttons_row, text="Tambah Folder...", style="Secondary.TButton", command=add_folder).pack(side="left")
    ttk.Button(buttons_row, text="Hapus Terpilih", style="Secondary.TButton", command=remove_selected).pack(
        side="left", padx=(8, 0)
    )
    ttk.Button(buttons_row, text="Bersihkan Semua", style="Secondary.TButton", command=clear_all).pack(
        side="left", padx=(8, 0)
    )

    return folder_paths


def build_frame(parent):
    frame = ttk.Frame(parent, padding=24, style="Card.TFrame")

    ttk.Label(frame, text=NAME, font=(theme.FONT_FAMILY, 15, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w"
    )
    ttk.Label(frame, text=DESCRIPTION, foreground=theme.TEXT_MUTED).grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(4, 20)
    )

    rekap_dir = data_dir("rekap-nilai")
    perwalian_dir = data_dir("rekap-nilai", "perwalian")

    perwalian_var = tk.StringVar()

    def pick_perwalian():
        path = filedialog.askdirectory(title="Pilih folder data perwalian (PDF)", initialdir=perwalian_dir)
        if path:
            perwalian_var.set(path)

    ttk.Label(frame, text="Folder Data Perwalian").grid(row=2, column=0, sticky="w", pady=6)
    ttk.Entry(frame, textvariable=perwalian_var, width=52, state="readonly").grid(
        row=2, column=1, sticky="we", padx=10, pady=6
    )
    ttk.Button(frame, text="Pilih Folder...", style="Secondary.TButton", command=pick_perwalian).grid(
        row=2, column=2, pady=6
    )

    nilai_folders = _build_folder_list_section(
        frame, 3, "Folder-Folder Data Nilai", initialdir=rekap_dir
    )

    output_var = tk.StringVar()
    output_is_auto = {"value": True}

    def pick_output():
        initial = output_var.get()
        initialdir = os.path.dirname(initial) if initial else rekap_dir
        initialfile = os.path.basename(initial) if initial else "Rekap Perwalian.xlsx"
        path = filedialog.asksaveasfilename(
            title="Simpan hasil sebagai",
            defaultextension=".xlsx",
            filetypes=[("File Excel", "*.xlsx")],
            initialdir=initialdir,
            initialfile=initialfile,
        )
        if path:
            output_var.set(path)
            output_is_auto["value"] = False

    ttk.Label(frame, text="Simpan Hasil Sebagai").grid(row=5, column=0, sticky="w", pady=6)
    output_entry = ttk.Entry(frame, textvariable=output_var, width=52)
    output_entry.grid(row=5, column=1, sticky="we", padx=10, pady=6)
    output_entry.bind("<Key>", lambda e: output_is_auto.__setitem__("value", False))
    ttk.Button(frame, text="Pilih...", style="Secondary.TButton", command=pick_output).grid(
        row=5, column=2, pady=6
    )

    frame.columnconfigure(1, weight=1)

    process_btn = ttk.Button(frame, text="Proses")
    process_btn.grid(row=6, column=0, columnspan=3, pady=(18, 8), sticky="w")

    result_row = ttk.Frame(frame)
    result_row.grid(row=7, column=0, columnspan=3, sticky="w")
    open_file_btn = ttk.Button(result_row, text="Buka File Hasil", style="Secondary.TButton")

    ttk.Label(frame, text="Log", foreground=theme.TEXT_MUTED, font=(theme.FONT_FAMILY, 9, "bold")).grid(
        row=8, column=0, sticky="w", pady=(14, 4)
    )
    log_text = tk.Text(frame, height=10, width=90, state="disabled", wrap="word")
    theme.style_text_widget(log_text, focus_border=False)
    log_text.grid(row=9, column=0, columnspan=3, sticky="nsew")
    frame.rowconfigure(9, weight=1)

    def log(msg):
        log_text.configure(state="normal")
        log_text.insert("end", msg + "\n")
        log_text.see("end")
        log_text.configure(state="disabled")

    result_queue = queue.Queue()

    def run_worker(perwalian_path, folders, output_path):
        try:
            roster, roster_skipped = core.load_perwalian_roster(perwalian_path)
            courses, skipped = core.generate_from_folders(folders)
            core.export_excel(roster, courses, output_path)
            result_queue.put(("ok", roster, roster_skipped, courses, skipped, output_path))
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
            _, roster, roster_skipped, courses, skipped, output_path = item
            log(f"Data perwalian terbaca: {len(roster)} mahasiswa")
            for name, msg in roster_skipped:
                log(f"  Dilewati (perwalian) - {name}: {msg}")
            log(f"Mata kuliah tergabung ({len(courses)}): {', '.join(c['course_name'] for c in courses)}")
            for name, msg in skipped:
                log(f"  Dilewati (nilai) - {name}: {msg}")
            for c in courses:
                if c["parse_error"]:
                    log(f"  Tidak dianalisis - {c['course_name']}: {c['parse_error']}")
            log(f"Tersimpan di: {output_path}")
            open_file_btn.grid(row=0, column=0)
        else:
            log(f"Gagal: {item[1]}")
            messagebox.showerror("Gagal memproses", item[1])

    def start_process():
        perwalian_path = perwalian_var.get().strip()
        folders = list(nilai_folders)
        output_path = output_var.get().strip()

        if not perwalian_path:
            messagebox.showwarning("Belum lengkap", "Pilih folder data perwalian terlebih dahulu.")
            return
        if not folders:
            messagebox.showwarning("Belum lengkap", "Tambahkan minimal 1 folder data nilai terlebih dahulu.")
            return
        if not output_path:
            messagebox.showwarning("Belum lengkap", "Tentukan lokasi file hasil terlebih dahulu.")
            return

        open_file_btn.grid_remove()
        process_btn.configure(state="disabled")
        log("Memproses...")
        threading.Thread(
            target=run_worker, args=(perwalian_path, folders, output_path), daemon=True
        ).start()
        frame.after(100, poll_queue)

    process_btn.configure(command=start_process)
    open_file_btn.configure(command=lambda: os.startfile(output_var.get()))

    return frame
