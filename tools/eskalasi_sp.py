"""
Tool UI: Proses SP dan DO.
Wrapper Tkinter di sekitar fungsi find_escalations()/generate() dari
eskalasi_sp.py & generate_sp_do.py. Satu tool untuk semua kebutuhan
terkait SP/DO - generate semester terbaru, cek eskalasi dari riwayat,
atau keduanya sekaligus.
"""

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import eskalasi_sp as sp_core
from core import generate_sp_do as gen_core
from tools import theme
from tools.common import data_dir

NAME = "Proses SP dan DO"
LABEL = "Proses SP dan DO"
ICON = "⚠️"  # ⚠️
DESCRIPTION = (
    "Bikin daftar calon SP Tahap Awal, SP Tahap Akhir, dan DO untuk "
    "semester ini secara otomatis dari data mahasiswa - sekaligus bisa "
    "cek siapa saja yang levelnya naik dibanding semester sebelumnya."
)


def _build_file_list_section(frame, row, label_text, filetypes, height=3, initialdir=None):
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

    file_paths = []

    def refresh():
        listbox.delete(0, "end")
        for path in file_paths:
            listbox.insert("end", os.path.basename(path))

    def add_files():
        paths = filedialog.askopenfilenames(
            title=f"Pilih {label_text}", filetypes=filetypes, initialdir=initialdir
        )
        for path in paths:
            if path not in file_paths:
                file_paths.append(path)
        refresh()

    def remove_selected():
        for idx in reversed(listbox.curselection()):
            del file_paths[idx]
        refresh()

    def clear_all():
        file_paths.clear()
        refresh()

    buttons_row = ttk.Frame(frame)
    buttons_row.grid(row=row + 1, column=1, columnspan=2, sticky="w", padx=10, pady=(0, 6))
    ttk.Button(buttons_row, text="Tambah File...", style="Secondary.TButton", command=add_files).pack(side="left")
    ttk.Button(buttons_row, text="Hapus Terpilih", style="Secondary.TButton", command=remove_selected).pack(
        side="left", padx=(8, 0)
    )
    ttk.Button(buttons_row, text="Bersihkan Semua", style="Secondary.TButton", command=clear_all).pack(
        side="left", padx=(8, 0)
    )

    return file_paths


def _build_folder_section(frame, row, label_text, on_pick, initialdir=None):
    """Baris folder-picker: Entry (path, hanya tampilan) + tombol Pilih Folder."""
    var = tk.StringVar()

    def pick():
        path = filedialog.askdirectory(title=f"Pilih {label_text}", initialdir=initialdir)
        if path:
            var.set(path)
            on_pick(path)

    ttk.Label(frame, text=label_text).grid(row=row, column=0, sticky="w", pady=6)
    ttk.Entry(frame, textvariable=var, width=52, state="readonly").grid(
        row=row, column=1, sticky="we", padx=10, pady=6
    )
    ttk.Button(frame, text="Pilih Folder...", style="Secondary.TButton", command=pick).grid(
        row=row, column=2, pady=6
    )
    return var


def build_frame(parent):
    frame = ttk.Frame(parent, padding=24, style="Card.TFrame")

    ttk.Label(frame, text=NAME, font=(theme.FONT_FAMILY, 15, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w"
    )

    ttk.Label(frame, text=DESCRIPTION, foreground=theme.TEXT_MUTED).grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(4, 20)
    )

    filetypes = [("File Excel", "*.xls *.xlsx"), ("Semua file", "*.*")]
    sp_do_dir = data_dir("surat-peringatan")

    aktif_paths = _build_file_list_section(
        frame, 2, "File Mahasiswa Aktif (Opsional, semester ini)", filetypes, initialdir=sp_do_dir
    )
    nonaktif_now_paths = _build_file_list_section(
        frame, 4, "File Mahasiswa Non-Aktif (Opsional, semester ini)", filetypes, initialdir=sp_do_dir
    )

    nonaktif_history_files = []
    lulus_files = []
    recap_history_files = []

    def scan_and_cache(folder_path, target_list, label_text):
        target_list.clear()
        matched, skipped = gen_core.scan_folder(folder_path)
        target_list.extend(matched)
        log(f"Folder {label_text}: {len(matched)} file terbaca ({os.path.basename(folder_path)}).")
        if skipped:
            log(f"  Dilewati (nama tidak cocok pola semester): {', '.join(skipped)}")

    _build_folder_section(
        frame, 6, "Folder Riwayat Non-Aktif (Opsional)",
        lambda p: scan_and_cache(p, nonaktif_history_files, "Riwayat Non-Aktif"),
        initialdir=data_dir("surat-peringatan", "nonaktif"),
    )
    _build_folder_section(
        frame, 7, "Folder Data Lulus (Opsional)",
        lambda p: scan_and_cache(p, lulus_files, "Lulus"),
        initialdir=data_dir("surat-peringatan", "lulus"),
    )

    _build_folder_section(
        frame, 8, "Folder Riwayat SP/DO (Opsional)",
        lambda p: scan_and_cache(p, recap_history_files, "Riwayat SP/DO"),
        initialdir=data_dir("surat-peringatan", "sp-dan-do"),
    )

    output_var = tk.StringVar()

    def _guess_output_filename():
        """Saran nama file mengikuti semester TERBARU dari data mentah yang sudah dipilih."""
        parsed = [
            gen_core.parse_term(p)
            for p in (*aktif_paths, *nonaktif_now_paths)
            if gen_core.parse_term(p) is not None
        ]
        if not parsed:
            return "Data SP dan DO.xlsx"
        year_start, term = max(parsed, key=lambda t: (t[0], 0 if t[1] == "Ganjil" else 1))
        return f"Data SP dan DO {year_start}-{year_start + 1} {term}.xlsx"

    def pick_output():
        initial = output_var.get()
        initialdir = os.path.dirname(initial) if initial else sp_do_dir
        initialfile = os.path.basename(initial) if initial else _guess_output_filename()
        path = filedialog.asksaveasfilename(
            title="Simpan hasil sebagai",
            defaultextension=".xlsx",
            filetypes=[("File Excel", "*.xlsx")],
            initialdir=initialdir,
            initialfile=initialfile,
        )
        if path:
            output_var.set(path)

    ttk.Label(frame, text="Simpan Hasil Sebagai").grid(row=9, column=0, sticky="w", pady=6)
    ttk.Entry(frame, textvariable=output_var, width=52).grid(
        row=9, column=1, sticky="we", padx=10, pady=6
    )
    ttk.Button(frame, text="Pilih...", style="Secondary.TButton", command=pick_output).grid(
        row=9, column=2, pady=6
    )

    frame.columnconfigure(1, weight=1)

    process_btn = ttk.Button(frame, text="Proses")
    process_btn.grid(row=10, column=0, columnspan=3, pady=(18, 8), sticky="w")

    result_row = ttk.Frame(frame)
    result_row.grid(row=11, column=0, columnspan=3, sticky="w")
    open_file_btn = ttk.Button(result_row, text="Buka File Hasil", style="Secondary.TButton")

    ttk.Label(frame, text="Log", foreground=theme.TEXT_MUTED, font=(theme.FONT_FAMILY, 9, "bold")).grid(
        row=12, column=0, sticky="w", pady=(14, 4)
    )
    log_text = tk.Text(frame, height=7, width=90, state="disabled", wrap="word")
    theme.style_text_widget(log_text, focus_border=False)
    log_text.grid(row=13, column=0, columnspan=3, sticky="nsew")
    frame.rowconfigure(13, weight=1)

    def log(msg):
        log_text.configure(state="normal")
        log_text.insert("end", msg + "\n")
        log_text.see("end")
        log_text.configure(state="disabled")

    result_queue = queue.Queue()

    def run_worker(aktif, nonaktif, lulus, recap, output_path):
        try:
            if aktif or nonaktif:
                result = gen_core.generate(aktif, nonaktif, lulus, recap)
                gen_core.export_excel(result, output_path)
                result_queue.put(("ok_generate", result, output_path))
            else:
                rows, semester_labels = sp_core.find_escalations(recap)
                sp_core.export_excel(rows, semester_labels, output_path)
                result_queue.put(("ok_escalation", rows, semester_labels, output_path))
        except Exception as e:
            result_queue.put(("error", str(e)))

    def poll_queue():
        try:
            item = result_queue.get_nowait()
        except queue.Empty:
            frame.after(100, poll_queue)
            return

        process_btn.configure(state="normal")
        if item[0] == "ok_generate":
            _, result, output_path = item
            log(f"Semester diproses ({len(result['semester_diproses'])}): {', '.join(result['semester_diproses'])}")
            log(f"Hasil untuk semester: {result['label']}")
            log(f"  SP Tahap Awal            : {len(result['sp_tahap_awal'])} mahasiswa "
                f"({len(result['sp_tahap_awal_eskalasi'])} eskalasi)")
            log(f"  SP Tahap Akhir           : {len(result['sp_tahap_akhir'])} mahasiswa "
                f"({len(result['sp_tahap_akhir_eskalasi'])} eskalasi)")
            log(f"  DO                       : {len(result['do'])} mahasiswa "
                f"({len(result['do_eskalasi'])} eskalasi)")
            log(f"  Rekap Non-Aktif          : {len(result['nonaktif_berkepanjangan'])} mahasiswa")
            log(f"Tersimpan di: {output_path}")
            open_file_btn.grid(row=0, column=0)
        elif item[0] == "ok_escalation":
            _, rows, semester_labels, output_path = item
            log(f"Semester dibaca ({len(semester_labels)}): {', '.join(semester_labels)}")
            log(f"Berhasil. Ditemukan {len(rows)} mahasiswa dengan eskalasi SP di semester terbaru.")
            log(f"Tersimpan di: {output_path}")
            open_file_btn.grid(row=0, column=0)
        else:
            log(f"Gagal: {item[1]}")
            messagebox.showerror("Gagal memproses", item[1])

    def start_process():
        output_path = output_var.get().strip()
        aktif = list(aktif_paths)
        nonaktif = list(nonaktif_now_paths) + list(nonaktif_history_files)
        lulus = list(lulus_files)
        recap_combined = list(recap_history_files)

        if not aktif and not nonaktif and not recap_combined:
            messagebox.showwarning(
                "Belum lengkap",
                "Isi data mahasiswa aktif/non-aktif semester ini, dan/atau pilih "
                "Folder Riwayat SP/DO, terlebih dahulu.",
            )
            return
        if not aktif and not nonaktif and len(recap_combined) < 2:
            messagebox.showwarning(
                "Belum lengkap",
                "Untuk cek eskalasi tanpa data mentah, pilih Folder Riwayat SP/DO "
                "yang berisi minimal 2 file rekap.",
            )
            return
        if not output_path:
            messagebox.showwarning("Belum lengkap", "Tentukan lokasi file hasil terlebih dahulu.")
            return

        open_file_btn.grid_remove()
        process_btn.configure(state="disabled")
        log("Memproses...")
        threading.Thread(
            target=run_worker, args=(aktif, nonaktif, lulus, recap_combined, output_path), daemon=True
        ).start()
        frame.after(100, poll_queue)

    process_btn.configure(command=start_process)
    open_file_btn.configure(command=lambda: os.startfile(output_var.get()))

    return frame
