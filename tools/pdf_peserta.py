"""
Tool UI: Update PDF Daftar Peserta Sertifikasi dari file Excel.
Wrapper Tkinter di sekitar fungsi build_pdf() dari update_pdf_dari_excel.py.
"""

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import update_pdf_dari_excel as pdf_core
from tools import theme
from tools.common import data_dir

NAME = "Daftar Peserta Sertifikasi Internal"
LABEL = "Pengajuan Peserta"
ICON = "\U0001F393"  # 🎓
DESCRIPTION = "Isi otomatis tabel NIM/Nama/Email di form PDF Daftar Peserta Sertifikasi dari data Excel - tidak perlu ketik manual satu-satu."


def build_frame(parent):
    frame = ttk.Frame(parent, padding=24, style="Card.TFrame")

    ttk.Label(frame, text=NAME, font=(theme.FONT_FAMILY, 15, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w"
    )

    ttk.Label(frame, text=DESCRIPTION, foreground=theme.TEXT_MUTED).grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(4, 20)
    )

    pengajuan_dir = data_dir("pengajuan-peserta")

    excel_var = tk.StringVar()
    pdf_var = tk.StringVar()
    output_var = tk.StringVar()
    output_is_auto = {"value": True}  # False setelah user pilih/ubah sendiri

    def other_folder(preferred_var, fallback_var):
        """Folder file yang sudah dipilih, buat jadi lokasi awal dialog berikutnya -
        default ke data/pengajuan-peserta/ kalau belum ada yang dipilih sama sekali."""
        for var in (preferred_var, fallback_var):
            if var.get():
                return os.path.dirname(var.get())
        return pengajuan_dir

    def suggest_output():
        """
        Path hasil default: nama file ikut PDF template, tapi folder ikut file
        Excel (folder kelas) - karena template kadang berupa file umum yang
        ditaruh di luar folder kelas, jadi tidak relevan sebagai lokasi simpan.
        """
        if not output_is_auto["value"] or not pdf_var.get():
            return
        folder = other_folder(excel_var, pdf_var)
        if folder is None:
            folder = os.path.dirname(pdf_var.get())
        base = os.path.splitext(os.path.basename(pdf_var.get()))[0]
        output_var.set(os.path.join(folder, f"{base} (Updated).pdf"))

    def pick_excel():
        path = filedialog.askopenfilename(
            title="Pilih file Excel",
            filetypes=[("File Excel", "*.xls *.xlsx"), ("Semua file", "*.*")],
            initialdir=other_folder(excel_var, pdf_var),
        )
        if path:
            excel_var.set(path)
            suggest_output()

    def pick_pdf():
        path = filedialog.askopenfilename(
            title="Pilih file PDF template",
            filetypes=[("File PDF", "*.pdf"), ("Semua file", "*.*")],
            initialdir=other_folder(pdf_var, excel_var),
        )
        if path:
            pdf_var.set(path)
            suggest_output()

    def pick_output():
        initial = output_var.get()
        initialdir = os.path.dirname(initial) if initial else other_folder(excel_var, pdf_var)
        initialfile = os.path.basename(initial) if initial else "Hasil.pdf"
        path = filedialog.asksaveasfilename(
            title="Simpan hasil sebagai",
            defaultextension=".pdf",
            filetypes=[("File PDF", "*.pdf")],
            initialdir=initialdir,
            initialfile=initialfile,
        )
        if path:
            output_var.set(path)
            output_is_auto["value"] = False

    def make_row(row, label, var, command):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=6)
        entry = ttk.Entry(frame, textvariable=var, width=52)
        entry.grid(row=row, column=1, sticky="we", padx=10, pady=6)
        ttk.Button(frame, text="Pilih...", style="Secondary.TButton", command=command).grid(
            row=row, column=2, pady=6
        )
        return entry

    make_row(2, "File Excel", excel_var, pick_excel)
    make_row(3, "PDF Template", pdf_var, pick_pdf)
    output_entry = make_row(4, "Simpan Sebagai", output_var, pick_output)
    # Kalau nama/lokasi hasil diketik manual, berhenti menimpanya otomatis.
    output_entry.bind("<Key>", lambda e: output_is_auto.__setitem__("value", False))

    frame.columnconfigure(1, weight=1)

    process_btn = ttk.Button(frame, text="Proses")
    process_btn.grid(row=5, column=0, columnspan=3, pady=(18, 8), sticky="w")

    result_row = ttk.Frame(frame)
    result_row.grid(row=6, column=0, columnspan=3, sticky="w")
    open_file_btn = ttk.Button(result_row, text="Buka File Hasil", style="Secondary.TButton")
    open_folder_btn = ttk.Button(result_row, text="Buka Folder Hasil", style="Secondary.TButton")

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

    def run_worker(excel_path, pdf_path, output_path):
        try:
            n_participants, n_pages = pdf_core.build_pdf(excel_path, pdf_path, output_path)
            result_queue.put(("ok", n_participants, n_pages, output_path))
        except SystemExit as e:
            result_queue.put(("error", str(e.code if e.code else e)))
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
            _, n_participants, n_pages, output_path = item
            log(f"Berhasil. Jumlah peserta: {n_participants}, jumlah halaman: {n_pages}")
            log(f"Tersimpan di: {output_path}")
            open_file_btn.grid(row=0, column=0, padx=(0, 8))
            open_folder_btn.grid(row=0, column=1)
        else:
            log(f"Gagal: {item[1]}")
            messagebox.showerror("Gagal memproses", item[1])

    def start_process():
        excel_path = excel_var.get().strip()
        pdf_path = pdf_var.get().strip()
        output_path = output_var.get().strip()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showwarning("Belum lengkap", "Pilih file Excel yang valid terlebih dahulu.")
            return
        if not pdf_path or not os.path.isfile(pdf_path):
            messagebox.showwarning("Belum lengkap", "Pilih file PDF template yang valid terlebih dahulu.")
            return
        if not output_path:
            messagebox.showwarning("Belum lengkap", "Tentukan lokasi file hasil terlebih dahulu.")
            return
        if os.path.abspath(output_path) == os.path.abspath(pdf_path):
            messagebox.showwarning(
                "Nama file sama",
                "File hasil tidak boleh sama dengan file PDF template (agar template asli tidak tertimpa).",
            )
            return

        open_file_btn.grid_remove()
        open_folder_btn.grid_remove()
        process_btn.configure(state="disabled")
        log("Memproses...")
        threading.Thread(
            target=run_worker, args=(excel_path, pdf_path, output_path), daemon=True
        ).start()
        frame.after(100, poll_queue)

    process_btn.configure(command=start_process)
    open_file_btn.configure(command=lambda: os.startfile(output_var.get()))
    open_folder_btn.configure(
        command=lambda: os.startfile(os.path.dirname(os.path.abspath(output_var.get())))
    )

    return frame
