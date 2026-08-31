"""
Tools: Update tabel Peserta Sertifikasi di PDF (form FM-SPK-01-01) agar sesuai
dengan data NIM/Nama/Email pada file Excel (.xls/.xlsx).

Cara pakai (dari folder ini):
    python update_pdf_dari_excel.py

Default: otomatis mencari 1 file .xls/.xlsx dan 1 file .pdf di folder ini,
lalu menyimpan hasilnya sebagai "<nama_pdf> (Updated).pdf".

Bisa juga eksplisit:
    python update_pdf_dari_excel.py --excel "AISD-A Pagi.xls" --pdf "template.pdf" --output "hasil.pdf"

Cara kerja:
- Header/kop form (Tahun Ajaran, Prodi, Kelas, Nama Dosen, dst) di halaman 1
  PDF dianggap sudah benar dan TIDAK diubah.
- Baris tabel peserta (No/NIM/NAMA/EMAIL) di semua halaman dibersihkan lalu
  diisi ulang dari Excel, dengan jumlah halaman menyesuaikan jumlah peserta
  (setiap halaman tambahan adalah salinan persis halaman 1, sehingga header,
  garis tabel, dan logo tetap identik dengan form asli).
- Field "Hal. : X dari Y" ikut diperbarui sesuai jumlah halaman hasil.
"""

import argparse
import glob
import math
import os
import sys

import fitz  # PyMuPDF
import pandas as pd

CALIBRI_REGULAR = r"C:\Windows\Fonts\calibri.ttf"
CALIBRI_BOLD = r"C:\Windows\Fonts\calibrib.ttf"
FONT_REGULAR = "PesertaCalibri"
FONT_BOLD = "PesertaCalibriBold"
DATA_FONT_SIZE = 6.5

_calibri_measure = None


def text_width(text, fontsize):
    # Lazy-loaded (bukan di level modul) supaya font Calibri yang hilang/
    # dipindah di komputer lain tidak membuat SELURUH aplikasi crash saat
    # start (app.py mengimpor semua tool di awal) - error-nya baru muncul
    # kalau tool ini yang benar-benar dipakai.
    global _calibri_measure
    if _calibri_measure is None:
        _calibri_measure = fitz.Font(fontfile=CALIBRI_REGULAR)
    return _calibri_measure.text_length(text, fontsize=fontsize)


def find_single(patterns, kind):
    matches = []
    for p in patterns:
        matches.extend(glob.glob(p))
    matches = sorted(set(matches))
    if len(matches) == 0:
        sys.exit(f"Tidak ada file {kind} ditemukan di folder ini.")
    if len(matches) > 1:
        sys.exit(
            f"Ditemukan lebih dari satu file {kind}: {matches}\n"
            f"Gunakan argumen eksplisit untuk memilih salah satu."
        )
    return matches[0]


def load_participants(excel_path):
    xls = pd.ExcelFile(excel_path)
    raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)

    header_row = None
    for i in range(min(15, len(raw))):
        row_values = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if "nim" in row_values and any("nama" in v for v in row_values) and any(
            "email" in v for v in row_values
        ):
            header_row = i
            break
    if header_row is None:
        sys.exit("Tidak menemukan baris header (NIM/Nama/Email) di file Excel.")

    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    col_nim = next(c for c in df.columns if c.strip().lower() == "nim")
    col_nama = next(c for c in df.columns if "nama" in c.strip().lower())
    col_email = next(c for c in df.columns if "email" in c.strip().lower())

    df = df.dropna(subset=[col_nim])

    participants = []
    for _, row in df.iterrows():
        nim_raw = row[col_nim]
        try:
            nim = str(int(nim_raw))
        except (ValueError, TypeError):
            nim = str(nim_raw).strip()
        nama = str(row[col_nama]).strip()
        email = "" if pd.isna(row[col_email]) else str(row[col_email]).strip()
        participants.append((nim, nama, email))

    if not participants:
        sys.exit("Tidak ada baris data peserta yang terbaca dari Excel.")
    return participants


def get_table_grid(page):
    draws = page.get_drawings()

    horiz = [
        d["rect"]
        for d in draws
        if d["type"] == "f"
        and (d["rect"].y1 - d["rect"].y0) < 2
        and (d["rect"].x1 - d["rect"].x0) > 400
        and 285 < d["rect"].y0 < 730
    ]
    vert = [
        d["rect"]
        for d in draws
        if d["type"] == "f"
        and (d["rect"].x1 - d["rect"].x0) < 2
        and (d["rect"].y1 - d["rect"].y0) > 300
        and d["rect"].y0 > 240
    ]

    if len(horiz) < 3 or len(vert) < 3:
        sys.exit("Tidak bisa mendeteksi grid tabel peserta pada PDF template.")

    # Batas atas data: tepat di bawah baris label kolom (No/NIM/NAMA/*EMAIL).
    header_hits = page.search_for("*EMAIL") or page.search_for("EMAIL")
    if not header_hits:
        sys.exit("Tidak menemukan label kolom *EMAIL pada tabel peserta.")
    header_bottom = max(h.y1 for h in header_hits) - 3

    # Batas bawah data: ujung bawah garis vertikal tabel (bukan bingkai luar
    # form yang sedikit lebih rendah).
    table_bottom = min(v.y1 for v in vert) + 1

    row_lines = sorted(
        set(round(r.y0, 2) for r in horiz if header_bottom <= r.y0 <= table_bottom)
    )
    col_lines = sorted(set(round(r.x0, 2) for r in vert))

    if len(col_lines) < 5:
        sys.exit(f"Jumlah garis kolom tidak sesuai dugaan (dapat {len(col_lines)}, butuh 5).")

    # 5 boundary x: border-kiri, No|NIM, NIM|NAMA, NAMA|EMAIL, border-kanan
    col_lines = col_lines[:5]

    return row_lines, col_lines


def find_hal_rect(page):
    hits = page.search_for("Hal.")
    if not hits:
        return None
    label = hits[0]
    # Ambil area di kanan label ": X dari Y" sampai margin kanan halaman
    return fitz.Rect(label.x1, label.y0, label.x1 + 90, label.y1)


def ensure_fonts(page):
    page.insert_font(fontname=FONT_REGULAR, fontfile=CALIBRI_REGULAR)
    page.insert_font(fontname=FONT_BOLD, fontfile=CALIBRI_BOLD)


def draw_text(page, point, text, fontname, fontsize, color=(0, 0, 0)):
    fontfile = CALIBRI_BOLD if fontname == FONT_BOLD else CALIBRI_REGULAR
    page.insert_text(point, text, fontname=fontname, fontfile=fontfile, fontsize=fontsize, color=color)


def fit_text_size(text, max_width, start_size=DATA_FONT_SIZE, min_size=4.5):
    size = start_size
    while size > min_size:
        if text_width(text, size) <= max_width:
            return size
        size -= 0.25
    return min_size


def build_pdf(excel_path, pdf_path, output_path):
    participants = load_participants(excel_path)

    doc = fitz.open(pdf_path)
    template_page = doc[0]
    row_lines, col_lines = get_table_grid(template_page)
    rows_per_page = len(row_lines) - 1

    border_left, div_no_nim, div_nim_nama, div_nama_email, border_right = col_lines

    n_pages = max(1, math.ceil(len(participants) / rows_per_page))

    # Buang semua halaman selain halaman pertama (yang jadi master template),
    # lalu gandakan halaman pertama sebanyak yang dibutuhkan.
    while doc.page_count > 1:
        doc.delete_page(1)
    if n_pages > 1:
        # insert_pdf melakukan deep-copy halaman (xref independen), berbeda
        # dari copy_page yang hanya menautkan ke content stream yang sama.
        # Sumbernya harus dokumen terpisah, jadi buka ulang file template.
        src = fitz.open(pdf_path)
        for _ in range(n_pages - 1):
            doc.insert_pdf(src, from_page=0, to_page=0, start_at=doc.page_count)
        src.close()

    for page_idx in range(n_pages):
        page = doc[page_idx]
        ensure_fonts(page)

        # Perbaiki "Hal. : X dari Y"
        hal_rect = find_hal_rect(page)
        if hal_rect is not None:
            page.add_redact_annot(hal_rect, fill=(1, 1, 1))
            page.apply_redactions()
            draw_text(
                page,
                (hal_rect.x0 + 1, hal_rect.y1 - 1.6),
                f": {page_idx + 1} dari {n_pages}",
                FONT_REGULAR,
                8.04,
            )

        start = page_idx * rows_per_page
        page_participants = participants[start : start + rows_per_page]

        # Bersihkan seluruh sel tabel di halaman ini dulu (No/NIM/NAMA/EMAIL)
        for r in range(rows_per_page):
            y_top, y_bot = row_lines[r], row_lines[r + 1]
            cells = [
                fitz.Rect(border_left + 1, y_top + 0.9, div_no_nim - 1.0, y_bot - 0.9),
                fitz.Rect(div_no_nim + 1.0, y_top + 0.9, div_nim_nama - 1.0, y_bot - 0.9),
                fitz.Rect(div_nim_nama + 1.0, y_top + 0.9, div_nama_email - 1.0, y_bot - 0.9),
                fitz.Rect(div_nama_email + 1.0, y_top + 0.9, border_right - 1.0, y_bot - 0.9),
            ]
            for cell in cells:
                page.add_redact_annot(cell, fill=(1, 1, 1))
        page.apply_redactions()

        # Isi ulang data
        for r in range(rows_per_page):
            y_top, y_bot = row_lines[r], row_lines[r + 1]
            baseline = y_bot - 1.25

            if r < len(page_participants):
                nim, nama, email = page_participants[r]
                no_text = str(start + r + 1)

                no_width = text_width(no_text, DATA_FONT_SIZE)
                no_x = div_no_nim - 1.0 - no_width
                draw_text(page, (no_x, baseline), no_text, FONT_REGULAR, DATA_FONT_SIZE)

                nim_size = fit_text_size(nim, div_nim_nama - div_no_nim - 3.0)
                draw_text(page, (div_no_nim + 2.0, baseline), nim, FONT_REGULAR, nim_size)

                nama_size = fit_text_size(nama, div_nama_email - div_nim_nama - 2.0)
                draw_text(page, (div_nim_nama + 0.8, baseline), nama, FONT_REGULAR, nama_size)

                if email:
                    email_size = fit_text_size(email, border_right - div_nama_email - 2.0)
                    draw_text(page, (div_nama_email + 1.5, baseline), email, FONT_REGULAR, email_size)

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    return len(participants), n_pages


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--excel", help="Path file Excel (.xls/.xlsx) sumber data NIM/Nama/Email")
    parser.add_argument("--pdf", help="Path file PDF template (form FM-SPK-01-01)")
    parser.add_argument("--output", help="Path file PDF hasil")
    args = parser.parse_args()

    excel_path = args.excel or find_single(["*.xls", "*.xlsx"], "Excel")
    pdf_path = args.pdf or find_single(["*.pdf"], "PDF")

    if args.output:
        output_path = args.output
    else:
        base, ext = os.path.splitext(pdf_path)
        output_path = f"{base} (Updated){ext}"

    n_participants, n_pages = build_pdf(excel_path, pdf_path, output_path)
    print(f"Excel : {excel_path}")
    print(f"PDF   : {pdf_path}")
    print(f"Hasil : {output_path}")
    print(f"Jumlah peserta : {n_participants}")
    print(f"Jumlah halaman : {n_pages}")


if __name__ == "__main__":
    main()
