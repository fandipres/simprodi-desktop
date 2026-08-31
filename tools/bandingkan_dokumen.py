"""
Tool UI: Bandingkan Dokumen (Word/PDF).
Wrapper Tkinter di sekitar fungsi compare_documents() dari
bandingkan_dokumen.py.
"""

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core import bandingkan_dokumen as compare_core
from tools import theme

NAME = "Bandingkan Dokumen (Word/PDF)"
LABEL = "Bandingkan Dokumen"
ICON = "\U0001F50D"  # 🔍
DESCRIPTION = (
    "Bandingkan 2 versi dokumen (Word atau PDF, boleh beda format) dan "
    "langsung lihat bagian mana saja yang berubah - tidak perlu baca ulang "
    "dari awal."
)


def build_frame(parent):
    frame = ttk.Frame(parent, padding=24, style="Card.TFrame")

    ttk.Label(frame, text=NAME, font=(theme.FONT_FAMILY, 15, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w"
    )

    ttk.Label(frame, text=DESCRIPTION, foreground=theme.TEXT_MUTED).grid(
        row=1, column=0, columnspan=3, sticky="w", pady=(4, 20)
    )

    doc_a_var = tk.StringVar()
    doc_b_var = tk.StringVar()
    output_var = tk.StringVar()
    output_is_auto = {"value": True}  # False setelah user pilih/ubah sendiri

    filetypes = [("Dokumen Word/PDF", "*.docx *.pdf"), ("Semua file", "*.*")]

    def other_folder(preferred_var, fallback_var):
        for var in (preferred_var, fallback_var):
            if var.get():
                return os.path.dirname(var.get())
        return None

    def suggest_output():
        if not output_is_auto["value"] or not doc_b_var.get():
            return
        folder = os.path.dirname(doc_b_var.get())
        base = os.path.splitext(os.path.basename(doc_b_var.get()))[0]
        output_var.set(os.path.join(folder, f"Perbandingan - {base}.html"))

    def pick_doc_a():
        path = filedialog.askopenfilename(
            title="Pilih dokumen versi lama",
            filetypes=filetypes,
            initialdir=other_folder(doc_a_var, doc_b_var),
        )
        if path:
            doc_a_var.set(path)
            suggest_output()

    def pick_doc_b():
        path = filedialog.askopenfilename(
            title="Pilih dokumen versi baru",
            filetypes=filetypes,
            initialdir=other_folder(doc_b_var, doc_a_var),
        )
        if path:
            doc_b_var.set(path)
            suggest_output()

    def pick_output():
        initial = output_var.get()
        initialdir = os.path.dirname(initial) if initial else other_folder(doc_b_var, doc_a_var)
        initialfile = os.path.basename(initial) if initial else "Hasil.html"
        path = filedialog.asksaveasfilename(
            title="Simpan hasil sebagai",
            defaultextension=".html",
            filetypes=[("File HTML", "*.html")],
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

    make_row(2, "Dokumen Lama", doc_a_var, pick_doc_a)
    make_row(3, "Dokumen Baru", doc_b_var, pick_doc_b)
    output_entry = make_row(4, "Simpan Hasil Sebagai", output_var, pick_output)
    output_entry.bind("<Key>", lambda e: output_is_auto.__setitem__("value", False))

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

    def run_worker(doc_a, doc_b, output_path):
        try:
            counts, out_path = compare_core.compare_documents(doc_a, doc_b, output_path)
            result_queue.put(("ok", counts, out_path))
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
            _, counts, out_path = item
            log(
                f"Berhasil. +{counts['added']} ditambahkan, "
                f"-{counts['removed']} dihapus, ~{counts['changed']} diubah."
            )
            log(f"Tersimpan di: {out_path}")
            open_file_btn.grid(row=0, column=0)
        else:
            log(f"Gagal: {item[1]}")
            messagebox.showerror("Gagal memproses", item[1])

    def start_process():
        doc_a = doc_a_var.get().strip()
        doc_b = doc_b_var.get().strip()
        output_path = output_var.get().strip()

        if not doc_a or not os.path.isfile(doc_a):
            messagebox.showwarning("Belum lengkap", "Pilih dokumen versi lama yang valid terlebih dahulu.")
            return
        if not doc_b or not os.path.isfile(doc_b):
            messagebox.showwarning("Belum lengkap", "Pilih dokumen versi baru yang valid terlebih dahulu.")
            return
        if not output_path:
            messagebox.showwarning("Belum lengkap", "Tentukan lokasi file hasil terlebih dahulu.")
            return

        open_file_btn.grid_remove()
        process_btn.configure(state="disabled")
        log("Memproses...")
        threading.Thread(
            target=run_worker, args=(doc_a, doc_b, output_path), daemon=True
        ).start()
        frame.after(100, poll_queue)

    process_btn.configure(command=start_process)
    open_file_btn.configure(command=lambda: os.startfile(output_var.get()))

    return frame
