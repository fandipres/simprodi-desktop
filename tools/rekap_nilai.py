"""
Tool UI: Rekap Nilai.
Wrapper Tkinter di sekitar fungsi generate()/export_excel() dari
core/rekap_nilai.py.
"""

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import rekap_nilai as core
from tools import theme
from tools.common import data_dir

NAME = "Rekap Nilai"
LABEL = "Rekap Nilai"
ICON = "\U0001F4CA"  # 📊
DESCRIPTION = (
    "Gabungkan file nilai per mata kuliah jadi 1 file, plus sheet Ringkasan "
    "berisi rekap kelulusan dan daftar mahasiswa tidak lulus per dosen wali."
)


def build_frame(parent):
    frame = ttk.Frame(parent, padding=24, style="Card.TFrame")

    ttk.Label(frame, text=NAME, font=(theme.FONT_FAMILY, 15, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w"
    )
    ttk.Label(frame, text=DESCRIPTION, foreground=theme.TEXT_MUTED).grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(4, 20)
    )

    rekap_dir = data_dir("rekap-nilai")

    folder_var = tk.StringVar()

    def pick_folder():
        path = filedialog.askdirectory(
            title="Pilih folder kumpulan file nilai", initialdir=rekap_dir
        )
        if path:
            folder_var.set(path)
            if output_is_auto["value"]:
                output_var.set(os.path.join(rekap_dir, f"Rekap Nilai {os.path.basename(path)}.xlsx"))

    ttk.Label(frame, text="Folder Kumpulan Nilai").grid(row=2, column=0, sticky="w", pady=6)
    ttk.Entry(frame, textvariable=folder_var, width=52, state="readonly").grid(
        row=2, column=1, sticky="we", padx=10, pady=6
    )
    ttk.Button(frame, text="Pilih Folder...", style="Secondary.TButton", command=pick_folder).grid(
        row=2, column=2, pady=6
    )

    perwalian_dir = data_dir("rekap-nilai", "perwalian")
    perwalian_var = tk.StringVar()

    def pick_perwalian():
        path = filedialog.askdirectory(
            title="Pilih folder data perwalian (PDF)", initialdir=perwalian_dir
        )
        if path:
            perwalian_var.set(path)

    ttk.Label(frame, text="Folder Data Perwalian (Opsional)").grid(row=3, column=0, sticky="w", pady=6)
    ttk.Entry(frame, textvariable=perwalian_var, width=52, state="readonly").grid(
        row=3, column=1, sticky="we", padx=10, pady=6
    )
    ttk.Button(frame, text="Pilih Folder...", style="Secondary.TButton", command=pick_perwalian).grid(
        row=3, column=2, pady=6
    )

    output_var = tk.StringVar()
    output_is_auto = {"value": True}

    def pick_output():
        initial = output_var.get()
        initialdir = os.path.dirname(initial) if initial else rekap_dir
        initialfile = os.path.basename(initial) if initial else "Rekap Nilai.xlsx"
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

    ttk.Label(frame, text="Simpan Hasil Sebagai").grid(row=4, column=0, sticky="w", pady=6)
    output_entry = ttk.Entry(frame, textvariable=output_var, width=52)
    output_entry.grid(row=4, column=1, sticky="we", padx=10, pady=6)
    output_entry.bind("<Key>", lambda e: output_is_auto.__setitem__("value", False))
    ttk.Button(frame, text="Pilih...", style="Secondary.TButton", command=pick_output).grid(
        row=4, column=2, pady=6
    )

    frame.columnconfigure(1, weight=1)

    process_btn = ttk.Button(frame, text="Proses")
    process_btn.grid(row=5, column=0, columnspan=3, pady=(18, 8), sticky="w")

    result_row = ttk.Frame(frame)
    result_row.grid(row=6, column=0, columnspan=3, sticky="w")
    open_file_btn = ttk.Button(result_row, text="Buka File Hasil", style="Secondary.TButton")

    ttk.Label(frame, text="Log", foreground=theme.TEXT_MUTED, font=(theme.FONT_FAMILY, 9, "bold")).grid(
        row=7, column=0, sticky="w", pady=(14, 4)
    )
    log_text = tk.Text(frame, height=10, width=90, state="disabled", wrap="word")
    theme.style_text_widget(log_text, focus_border=False)
    log_text.grid(row=8, column=0, columnspan=3, sticky="nsew")
    frame.rowconfigure(8, weight=1)

    def log(msg):
        log_text.configure(state="normal")
        log_text.insert("end", msg + "\n")
        log_text.see("end")
        log_text.configure(state="disabled")

    result_queue = queue.Queue()

    def run_worker(folder_path, perwalian_path, output_path):
        try:
            courses, skipped = core.generate(folder_path)
            perwalian, perwalian_skipped = None, []
            if perwalian_path:
                perwalian, perwalian_skipped = core.load_perwalian(perwalian_path)
            core.export_excel(courses, output_path, perwalian)
            result_queue.put(("ok", courses, skipped, perwalian, perwalian_skipped, output_path))
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
            _, courses, skipped, perwalian, perwalian_skipped, output_path = item
            log(f"Mata kuliah tergabung ({len(courses)}): {', '.join(c['course_name'] for c in courses)}")
            for name, msg in skipped:
                log(f"  Dilewati (gagal dibuka) - {name}: {msg}")
            for c in courses:
                if c["parse_error"]:
                    log(f"  Tidak dianalisis di Ringkasan - {c['course_name']}: {c['parse_error']}")
            if perwalian is not None:
                log(f"Data perwalian terbaca: {len(perwalian)} mahasiswa")
                for name, msg in perwalian_skipped:
                    log(f"  Dilewati (perwalian) - {name}: {msg}")
            log(f"Tersimpan di: {output_path}")
            open_file_btn.grid(row=0, column=0)
        else:
            log(f"Gagal: {item[1]}")
            messagebox.showerror("Gagal memproses", item[1])

    def start_process():
        folder_path = folder_var.get().strip()
        perwalian_path = perwalian_var.get().strip()
        output_path = output_var.get().strip()

        if not folder_path:
            messagebox.showwarning("Belum lengkap", "Pilih folder kumpulan file nilai terlebih dahulu.")
            return
        if not output_path:
            messagebox.showwarning("Belum lengkap", "Tentukan lokasi file hasil terlebih dahulu.")
            return

        open_file_btn.grid_remove()
        process_btn.configure(state="disabled")
        log("Memproses...")
        threading.Thread(
            target=run_worker, args=(folder_path, perwalian_path, output_path), daemon=True
        ).start()
        frame.after(100, poll_queue)

    process_btn.configure(command=start_process)
    open_file_btn.configure(command=lambda: os.startfile(output_var.get()))

    return frame
