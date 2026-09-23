"""
Tools: Gabungkan file nilai per mata kuliah (format "Evaluasi Ketercapaian
MK"/OBE hasil export sistem akademik) jadi 1 file Excel - satu sheet per mata
kuliah (disalin apa adanya, termasuk formatnya), diurutkan abjad berdasarkan
nama mata kuliah, ditambah 1 sheet "Ringkasan" di depan berisi rekap
kelulusan per mata kuliah dan daftar mahasiswa yang tidak lulus.

Tiap mata kuliah boleh beda jumlah Sub-CPMK/kolom penilaian (makanya lebar
tabelnya beda-beda per file) - kolom yang dibutuhkan (NIM, Nama Mahasiswa,
Status Kelulusan MK) dicari lewat teks header-nya, bukan posisi kolom tetap,
supaya tetap kebaca walau jumlah kolom di antaranya berubah. Kalau suatu file
formatnya tidak dikenali sama sekali (bukan format ini), sheet-nya TETAP
disalin ke file gabungan, cuma tidak ikut dihitung di Ringkasan.

Opsional: folder berisi PDF "Daftar Hadir Mahasiswa Pembimbingan Akademik"
(form FM-FKT-02-01) bisa diikutkan lewat load_perwalian()/--perwalian, supaya
tabel mahasiswa tidak lulus ditandai dosen wali SAAT INI (bukan dosen wali
waktu nilai itu diambil - keduanya bisa beda kalau ada pergantian dosen wali
antar semester) dan dikelompokkan per dosen wali.

Cara pakai (dari folder ini):
    python rekap_nilai.py --folder "path/ke/folder/nilai" --output "hasil.xlsx" \\
        [--perwalian "path/ke/folder/perwalian"]
"""
import argparse
import copy
import os
import re

import fitz  # PyMuPDF
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

LULUS = "L"


def scan_folder(folder_path):
    """Daftar file Excel (.xls/.xlsx) di folder_path, diurutkan nama file -
    dipakai cuma sebagai urutan awal sebelum di-parse (urutan sheet final
    tetap ikut nama mata kuliah hasil parsing, lihat generate())."""
    matched = []
    with os.scandir(folder_path) as entries:
        for entry in sorted(entries, key=lambda e: e.name):
            if not entry.is_file() or entry.name.startswith("~$"):
                continue
            if entry.name.lower().endswith((".xls", ".xlsx")):
                matched.append(entry.path)
    return matched


def scan_pdf_folder(folder_path):
    """Daftar file PDF di folder_path, diurutkan nama file."""
    matched = []
    with os.scandir(folder_path) as entries:
        for entry in sorted(entries, key=lambda e: e.name):
            if entry.is_file() and not entry.name.startswith("~$") and entry.name.lower().endswith(".pdf"):
                matched.append(entry.path)
    return matched


_DOSEN_WALI_RE = re.compile(r"Dosen Penasihat Akademik\s*\n\s*:\s*([^\n]+)")
_KELAS_RE = re.compile(r"Kelas\s*\n\s*:\s*([^\n]+)")
_SEMESTER_RE = re.compile(r"Semester\s*\n\s*:\s*([^\n]+)")
_PERWALIAN_ROW_RE = re.compile(r"(\d+)\.\s*\n\s*(\d{6,10})\s*\n\s*([^\n]+)")


def _parse_perwalian_pdf(path):
    """
    Baca 1 PDF "Daftar Hadir Mahasiswa Pembimbingan Akademik" (form
    FM-FKT-02-01) - kembalikan dict {dosen, kelas, semester, students}
    (students = {nim: nama_mahasiswa}). Raise ValueError kalau markah yang
    dibutuhkan (baris "Dosen Penasihat Akademik : ..." atau baris NIM
    mahasiswa) tidak ditemukan - dianggap bukan format ini.
    """
    doc = fitz.open(path)
    text = "\n".join(page.get_text() for page in doc)
    doc.close()

    dosen_matches = _DOSEN_WALI_RE.findall(text)
    if not dosen_matches:
        raise ValueError('Baris "Dosen Penasihat Akademik : ..." tidak ditemukan.')
    dosen = dosen_matches[0].strip()

    rows = _PERWALIAN_ROW_RE.findall(text)
    if not rows:
        raise ValueError("Tidak ada baris NIM mahasiswa yang terbaca.")

    kelas_matches = _KELAS_RE.findall(text)
    sem_matches = _SEMESTER_RE.findall(text)

    students = {nim.strip(): nama.strip() for _, nim, nama in rows}
    return {
        "dosen": dosen,
        "kelas": kelas_matches[0].strip() if kelas_matches else None,
        "semester": sem_matches[0].strip() if sem_matches else None,
        "students": students,
    }


def load_perwalian(folder_path):
    """
    Baca semua PDF di folder_path. Kembalikan (mapping, skipped): mapping =
    {nim: nama_dosen_wali_saat_ini}, skipped = list (nama_file, pesan_error)
    untuk PDF yang formatnya tidak dikenali (dilewati, tidak menggagalkan
    yang lain).
    """
    mapping, skipped = {}, []
    for path in scan_pdf_folder(folder_path):
        try:
            info = _parse_perwalian_pdf(path)
        except Exception as e:
            skipped.append((os.path.basename(path), str(e)))
            continue
        for nim in info["students"]:
            mapping[nim] = info["dosen"]
    return mapping, skipped


def load_perwalian_roster(folder_path):
    """
    Baca semua PDF di folder_path. Kembalikan (roster, skipped): roster =
    {nim: {"nama", "dosen", "kelas", "semester"}} (satu entri per mahasiswa
    - dipakai fitur "Rekap Perwalian" yang mulai dari roster, bukan dari
    file nilai), skipped = list (nama_file, pesan_error) untuk PDF yang
    formatnya tidak dikenali.
    """
    roster, skipped = {}, []
    for path in scan_pdf_folder(folder_path):
        try:
            info = _parse_perwalian_pdf(path)
        except Exception as e:
            skipped.append((os.path.basename(path), str(e)))
            continue
        for nim, nama in info["students"].items():
            roster[nim] = {
                "nama": nama,
                "dosen": info["dosen"],
                "kelas": info["kelas"],
                "semester": info["semester"],
            }
    return roster, skipped


def _label_value(grid, nrows, ncols, label):
    """Cari sel yang isinya persis label (case-insensitive), kembalikan nilai
    sel pertama yang terisi di baris yang sama, sebelah kanannya."""
    label_lower = label.strip().lower()
    for r in range(nrows):
        for c in range(ncols):
            v = grid[r][c]
            if isinstance(v, str) and v.strip().lower() == label_lower:
                for cc in range(c + 1, ncols):
                    if grid[r][cc] not in (None, ""):
                        return grid[r][cc]
                return None
    return None


def _find_cell(grid, predicate, row_range, col_range):
    for r in row_range:
        for c in col_range:
            if predicate(grid[r][c]):
                return r, c
    return None


def _looks_like_nim(v):
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str) and v.strip().isdigit() and len(v.strip()) >= 5:
        return True
    return False


def _parse_students(grid, nrows, ncols):
    """Cari tabel 'NIM/Nama Mahasiswa/.../Status Kelulusan MK' lewat teks
    header (bukan posisi tetap) dan baca baris datanya. Raise ValueError
    dengan pesan jelas kalau markah yang dibutuhkan tidak ketemu."""
    nim_pos = _find_cell(
        grid, lambda v: isinstance(v, str) and v.strip().upper() == "NIM",
        range(nrows), range(ncols),
    )
    if nim_pos is None:
        raise ValueError('Kolom "NIM" tidak ditemukan - format tabel mahasiswa tidak dikenali.')
    header_row, nim_col = nim_pos
    nama_col = nim_col + 1
    if nama_col >= ncols:
        raise ValueError('Kolom di sebelah kanan "NIM" (harusnya "Nama Mahasiswa") tidak ada.')

    marker_rows = range(header_row, min(header_row + 4, nrows))
    status_pos = _find_cell(
        grid,
        lambda v: isinstance(v, str) and "kelulusan" in v.lower() and " mk" in v.lower(),
        marker_rows, range(ncols),
    )
    status_col = status_pos[1] if status_pos else ncols - 1

    letter_pos = _find_cell(
        grid,
        lambda v: isinstance(v, str) and "huruf" in v.lower() and "mk" in v.lower(),
        marker_rows, range(ncols),
    )
    letter_col = letter_pos[1] if letter_pos else None

    score_pos = _find_cell(
        grid,
        lambda v: isinstance(v, str) and "mentah akhir mk" in v.lower(),
        marker_rows, range(ncols),
    )
    score_col = score_pos[1] if score_pos else None

    data_start = None
    for r in range(header_row + 1, nrows):
        if _looks_like_nim(grid[r][nim_col]):
            data_start = r
            break
    if data_start is None:
        raise ValueError('Baris data mahasiswa (di bawah header "NIM") tidak ditemukan.')

    students = []
    r = data_start
    while r < nrows and _looks_like_nim(grid[r][nim_col]):
        nim_raw = grid[r][nim_col]
        nim = str(int(nim_raw)) if isinstance(nim_raw, float) else str(nim_raw).strip()
        nama_raw = grid[r][nama_col]
        status_raw = grid[r][status_col] if status_col < ncols else None
        students.append({
            "nim": nim,
            "nama": str(nama_raw).strip() if nama_raw not in (None, "") else "",
            "nilai_akhir": grid[r][score_col] if score_col is not None and score_col < ncols else None,
            "nilai_huruf": grid[r][letter_col] if letter_col is not None and letter_col < ncols else None,
            "status": str(status_raw).strip().upper() if status_raw not in (None, "") else "",
        })
        r += 1
    if not students:
        raise ValueError("Tidak ada baris data mahasiswa yang terbaca di bawah header.")
    return students


def parse_course_file(path):
    """
    Baca 1 file nilai mata kuliah. Selalu berhasil selama filenya adalah
    Excel yang valid - worksheet sumbernya (buat disalin apa adanya) dan
    metadata (nama MK, kelas, dosen) tetap dikembalikan walau tabel
    mahasiswanya gagal dikenali; dalam kasus itu "students" kosong dan
    "parse_error" berisi alasannya (sheet TETAP disalin, cuma tidak ikut
    dihitung di Ringkasan).
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = [[cell.value for cell in row] for row in ws.iter_rows()]
    nrows = len(rows)
    ncols = max((len(r) for r in rows), default=0)
    grid = [r + [None] * (ncols - len(r)) for r in rows]

    course_name = _label_value(grid, nrows, ncols, "Nama MK")
    if not course_name:
        course_name = os.path.splitext(os.path.basename(path))[0]
    kelas = _label_value(grid, nrows, ncols, "Kelas")
    dosen = _label_value(grid, nrows, ncols, "Dosen")

    result = {
        "source_path": path,
        "worksheet": ws,
        "course_name": str(course_name).strip(),
        "kelas": str(kelas).strip() if kelas else None,
        "dosen": str(dosen).strip() if dosen else None,
        "students": [],
        "parse_error": None,
    }
    try:
        result["students"] = _parse_students(grid, nrows, ncols)
    except ValueError as e:
        result["parse_error"] = str(e)
    return result


def generate_from_folders(folder_paths):
    """
    Sama seperti generate(), tapi menggabungkan file dari BEBERAPA folder
    sekaligus (dipakai fitur Rekap Perwalian, buat mencari nilai mahasiswa
    lintas beberapa folder/semester data nilai dalam satu proses).
    """
    paths = []
    for folder in folder_paths:
        paths.extend(scan_folder(folder))
    if not paths:
        raise ValueError("Tidak ada file Excel (.xls/.xlsx) yang ditemukan di folder-folder ini.")

    courses, skipped = [], []
    for path in paths:
        try:
            courses.append(parse_course_file(path))
        except Exception as e:
            skipped.append((os.path.basename(path), str(e)))

    if not courses:
        raise ValueError(
            "Tidak ada file yang berhasil dibuka:\n"
            + "\n".join(f"- {name}: {msg}" for name, msg in skipped)
        )

    courses.sort(key=lambda c: c["course_name"].lower())
    return courses, skipped


def generate(folder_path):
    """
    Parse semua file di folder_path. Kembalikan (courses, skipped):
    courses = list dict hasil parse_course_file(), terurut abjad berdasarkan
    nama mata kuliah; skipped = list (nama_file, pesan_error) untuk file yang
    GAGAL DIBUKA sama sekali (bukan .xlsx/.xls valid - beda dari
    parse_error, yang masih menghasilkan course dengan sheet tetap disalin).
    """
    return generate_from_folders([folder_path])


def _failing_students(courses):
    """{nim: {"nama":..., "gagal": [nama_mk, ...]}} lintas semua mata kuliah
    yang berhasil di-parse, lalu diringkas jadi list baris terurut (jumlah MK
    gagal terbanyak dulu, lalu NIM)."""
    by_nim = {}
    for course in courses:
        if course["parse_error"]:
            continue
        for s in course["students"]:
            if not s["nim"] or not s["status"] or s["status"] == LULUS:
                continue
            entry = by_nim.setdefault(s["nim"], {"nama": "", "gagal": []})
            if s["nama"]:
                entry["nama"] = s["nama"]
            entry["gagal"].append(course["course_name"])

    rows = [
        {
            "nim": nim,
            "nama": info["nama"],
            "jumlah_gagal": len(info["gagal"]),
            "daftar_gagal": ", ".join(info["gagal"]),
        }
        for nim, info in by_nim.items()
    ]
    rows.sort(key=lambda r: (-r["jumlah_gagal"], r["nim"]))
    return rows


def _grades_by_nim(courses):
    """{nim: {"lulus": [nama_mk,...], "gagal": [nama_mk,...]}} dari seluruh
    mata kuliah yang berhasil di-parse - dipakai buat mencari nilai
    mahasiswa dari roster perwalian (lihat generate_perwalian_rekap)."""
    by_nim = {}
    for course in courses:
        if course["parse_error"]:
            continue
        for s in course["students"]:
            if not s["nim"] or not s["status"]:
                continue
            entry = by_nim.setdefault(s["nim"], {"lulus": [], "gagal": []})
            if s["status"] == LULUS:
                entry["lulus"].append(course["course_name"])
            else:
                entry["gagal"].append(course["course_name"])
    return by_nim


_TANPA_DATA_NILAI = "(Tidak ada data nilai ditemukan di folder yang dipilih)"


def generate_perwalian_rekap(roster, courses):
    """
    Mulai dari roster perwalian (bukan dari file nilai) - buat 1 baris per
    mahasiswa di roster, cari nilainya di seluruh 'courses' (hasil
    generate_from_folders(), boleh gabungan banyak folder/semester). Kembalikan
    list baris terurut: per dosen wali (abjad), lalu di dalam tiap dosen wali
    - yang gagal >=1 MK dulu (paling banyak gagal duluan), baru yang lulus
    semua, baru yang datanya belum ketemu sama sekali.
    """
    grades = _grades_by_nim(courses)

    rows = []
    for nim, info in roster.items():
        g = grades.get(nim, {"lulus": [], "gagal": []})
        jumlah_ditemukan = len(g["lulus"]) + len(g["gagal"])
        jumlah_gagal = len(g["gagal"])
        if jumlah_ditemukan == 0:
            daftar_gagal = _TANPA_DATA_NILAI
        elif jumlah_gagal == 0:
            daftar_gagal = "-"
        else:
            daftar_gagal = ", ".join(g["gagal"])
        rows.append({
            "dosen": info["dosen"] or "",
            "nim": nim,
            "nama": info["nama"],
            "jumlah_ditemukan": jumlah_ditemukan,
            "jumlah_gagal": jumlah_gagal,
            "daftar_gagal": daftar_gagal,
        })

    def tier(row):
        if row["jumlah_gagal"] > 0:
            return 0
        if row["jumlah_ditemukan"] > 0:
            return 1
        return 2

    rows.sort(key=lambda r: (r["dosen"], tier(r), -r["jumlah_gagal"], r["nim"]))
    return rows


def _copy_sheet(src_ws, wb, title):
    """Salin 1 worksheet APA ADANYA (nilai, style per sel, merge cell, lebar
    kolom/tinggi baris) ke workbook tujuan - dipakai supaya file nilai asli
    per mata kuliah tetap terlihat sama persis sebagai salah satu sheet di
    file gabungan."""
    ws = wb.create_sheet(title)
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy.copy(cell.font)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.border = copy.copy(cell.border)
                new_cell.alignment = copy.copy(cell.alignment)
                new_cell.number_format = cell.number_format
    for merged_range in src_ws.merged_cells.ranges:
        ws.merge_cells(str(merged_range))
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            ws.column_dimensions[col_letter].width = dim.width
    for row_idx, dim in src_ws.row_dimensions.items():
        if dim.height:
            ws.row_dimensions[row_idx].height = dim.height
    return ws


def _course_sheet_name(course, used):
    """
    Nama sheet dari mata kuliah + kelas (kalau ada) - PENTING dipakai kelas
    (bukan nama mata kuliah saja) karena satu mata kuliah bisa punya
    beberapa kelas paralel (mis. IF-A Pagi, IF-A Sore, IF-B Pagi, dst)
    waktu beberapa folder data nilai digabung sekaligus - tanpa kelas,
    sheet-nya cuma kebedakan lewat akhiran " (2)", " (3)" yang tidak
    informatif.

    Batas nama sheet Excel cuma 31 karakter - kalau nama mata kuliah +
    kelas kepanjangan, bagian KELAS diprioritaskan tetap utuh (nama mata
    kuliahnya yang dipotong duluan), supaya tetap bisa dibedakan tanpa
    perlu buka isi sheet-nya dulu.
    """
    name = "".join(ch for ch in course["course_name"] if ch not in r":\/?*[]").strip() or "Mata Kuliah"
    suffix = f" - {course['kelas']}" if course["kelas"] else ""
    suffix = "".join(ch for ch in suffix if ch not in r":\/?*[]")[:31]

    def build(n=None):
        extra = f" ({n})" if n else ""
        max_name_len = 31 - len(suffix) - len(extra)
        if max_name_len < 1:
            return (suffix + extra)[:31] if suffix else (name[:31 - len(extra)] + extra)
        return name[:max_name_len] + suffix + extra

    candidate = build()
    n = 2
    while candidate.lower() in used:
        candidate = build(n)
        n += 1
    used.add(candidate.lower())
    return candidate


def export_excel(courses, output_path, perwalian=None):
    wb = Workbook()
    wb.remove(wb.active)
    ws_ringkasan = wb.create_sheet("Ringkasan", 0)

    used_names = set()
    for course in courses:
        title = _course_sheet_name(course, used_names)
        _copy_sheet(course["worksheet"], wb, title)

    _write_ringkasan(ws_ringkasan, courses, perwalian)
    wb.save(output_path)


def export_perwalian_excel(roster, courses, output_path):
    """Versi export_excel() yang mulai dari roster perwalian - lihat
    generate_perwalian_rekap()."""
    wb = Workbook()
    wb.remove(wb.active)
    ws_ringkasan = wb.create_sheet("Ringkasan", 0)

    used_names = set()
    for course in courses:
        title = _course_sheet_name(course, used_names)
        _copy_sheet(course["worksheet"], wb, title)

    rows = generate_perwalian_rekap(roster, courses)
    _write_perwalian_ringkasan(ws_ringkasan, rows)
    wb.save(output_path)


_HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
_FAIL_FILL = PatternFill("solid", fgColor="FCE4E4")
_NODATA_FILL = PatternFill("solid", fgColor="F2F2F2")
_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center")


_LEFT = Alignment(horizontal="left")


def _write_table(ws, start_row, headers, rows, col_widths, highlight_rows=None, left_align_cols=None, row_fills=None):
    """row_fills: {row_index: PatternFill} - fill per baris data (0-based,
    relatif ke rows), lebih fleksibel dari highlight_rows (yang selalu
    _FAIL_FILL) kalau butuh lebih dari 1 warna sekaligus."""
    left_align_cols = left_align_cols or set()
    row_fills = dict(row_fills or {})
    for i in (highlight_rows or set()):
        row_fills.setdefault(i, _FAIL_FILL)
    for c, header in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c, value=header)
        cell.font = Font(bold=True)
        cell.fill = _HEADER_FILL
        cell.border = _BORDER
        cell.alignment = _CENTER
    for i, row in enumerate(rows):
        r = start_row + 1 + i
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = _BORDER
            cell.alignment = _LEFT if c in left_align_cols else _CENTER
            if i in row_fills:
                cell.fill = row_fills[i]
    for c, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = width
    return start_row + 1 + len(rows)


_TANPA_PERWALIAN = "(Tidak ada di data perwalian)"


def _write_ringkasan(ws, courses, perwalian=None):
    ws.cell(row=1, column=1, value="Ringkasan Nilai").font = Font(bold=True, size=14)

    kelas_set = sorted({c["kelas"] for c in courses if c["kelas"]})
    if kelas_set:
        ws.cell(row=2, column=1, value=f"Kelas: {', '.join(kelas_set)}")
    ws.cell(row=3, column=1, value=f"Jumlah Mata Kuliah: {len(courses)}")

    row = 5
    ws.cell(row=row, column=1, value="Rekap per Mata Kuliah").font = Font(bold=True, size=12)
    row += 1

    course_rows = []
    for course in courses:
        if course["parse_error"]:
            course_rows.append((
                course["course_name"], course["kelas"] or "", course["dosen"] or "",
                "", "", "", f"Tidak dianalisis - {course['parse_error']}",
            ))
            continue
        total = len(course["students"])
        lulus = sum(1 for s in course["students"] if s["status"] == LULUS)
        dinilai = sum(1 for s in course["students"] if s["status"])
        tidak_lulus = dinilai - lulus
        persen = f"{lulus / dinilai * 100:.0f}%" if dinilai else ""
        course_rows.append((
            course["course_name"], course["kelas"] or "", course["dosen"] or "",
            total, lulus, tidak_lulus, persen,
        ))

    # Kolom "Kelas" penting kalau mata kuliah yang sama punya beberapa
    # kelas paralel (lihat catatan di _course_sheet_name) - tanpa ini, baris
    # "Mata Kuliah" yang sama berulang tanpa cara membedakan kelasnya.
    row = _write_table(
        ws, row,
        ["Mata Kuliah", "Kelas", "Dosen", "Jumlah Mahasiswa", "Lulus", "Tidak Lulus", "% Lulus"],
        course_rows,
        [30, 16, 30, 16, 10, 12, 24],
        left_align_cols={1, 2, 3},
    )

    row += 2
    ws.cell(row=row, column=1, value="Mahasiswa Tidak Lulus (minimal 1 mata kuliah)").font = Font(bold=True, size=12)
    row += 1

    failing = _failing_students(courses)
    if failing:
        # Kolom "Catatan" sengaja dikosongkan - dipakai dosen wali menandai
        # tindak lanjut per mahasiswa (mis. "sudah dihubungi", "remedial"),
        # bukan diisi otomatis oleh aplikasi.
        if perwalian is not None:
            for f in failing:
                f["dosen_wali"] = perwalian.get(f["nim"]) or _TANPA_PERWALIAN
            # Dikelompokkan per dosen wali (abjad, yang tidak ketemu datanya
            # ditaruh paling akhir), baru di dalam tiap dosen diurutkan jumlah
            # MK gagal terbanyak dulu.
            failing.sort(key=lambda f: (
                f["dosen_wali"] == _TANPA_PERWALIAN, f["dosen_wali"],
                -f["jumlah_gagal"], f["nim"],
            ))
            headers = ["NIM", "Nama", "Dosen Wali (Saat Ini)", "Jumlah MK Gagal", "Mata Kuliah yang Gagal", "Catatan"]
            fail_rows = [
                (f["nim"], f["nama"], f["dosen_wali"], f["jumlah_gagal"], f["daftar_gagal"], "")
                for f in failing
            ]
            widths = [16, 30, 30, 16, 60, 30]
            left_cols = {2, 3, 5}
        else:
            headers = ["NIM", "Nama", "Jumlah MK Gagal", "Mata Kuliah yang Gagal", "Catatan"]
            fail_rows = [(f["nim"], f["nama"], f["jumlah_gagal"], f["daftar_gagal"], "") for f in failing]
            widths = [16, 30, 16, 60, 30]
            left_cols = {2, 4}
        row = _write_table(
            ws, row, headers, fail_rows, widths,
            highlight_rows=set(range(len(fail_rows))),
            left_align_cols=left_cols,
        )
    else:
        ws.cell(row=row, column=1, value="Tidak ada mahasiswa yang tidak lulus di mata kuliah manapun.")
        row += 1

    row += 1
    note_text = (
        "Catatan: data ini belum memperhitungkan nilai remedial, dan belum "
        "mengecualikan mahasiswa yang sudah keluar/nonaktif di semester "
        "berikutnya - cek ulang manual untuk kasus-kasus tersebut sebelum "
        "dipakai sebagai dasar keputusan."
    )
    if perwalian is not None:
        note_text += (
            f' Kolom "Dosen Wali (Saat Ini)" diambil dari file data perwalian '
            f'yang dipilih - status "{_TANPA_PERWALIAN}" berarti NIM tersebut '
            "tidak ditemukan di file perwalian manapun yang dipilih (bisa "
            "karena filenya tidak disertakan, atau mahasiswanya sudah tidak "
            "aktif)."
        )
    note = ws.cell(
        row=row, column=1,
        value=(
            note_text
        ),
    )
    note.font = Font(italic=True, color="808080")


def _write_perwalian_ringkasan(ws, rows):
    ws.cell(row=1, column=1, value="Ringkasan Perwalian").font = Font(bold=True, size=14)

    dosen_set = sorted({r["dosen"] for r in rows if r["dosen"]})
    ws.cell(row=2, column=1, value=f"Jumlah Dosen Wali: {len(dosen_set)}")
    ws.cell(row=3, column=1, value=f"Jumlah Mahasiswa (roster perwalian): {len(rows)}")

    row = 5
    # Kolom "Catatan" sengaja dikosongkan - dipakai dosen wali menandai
    # tindak lanjut per mahasiswa (mis. "sudah dihubungi", "remedial"),
    # bukan diisi otomatis oleh aplikasi.
    table_rows = [
        (r["dosen"], r["nim"], r["nama"], r["jumlah_ditemukan"], r["jumlah_gagal"], r["daftar_gagal"], "")
        for r in rows
    ]
    row_fills = {}
    for i, r in enumerate(rows):
        if r["jumlah_gagal"] > 0:
            row_fills[i] = _FAIL_FILL
        elif r["jumlah_ditemukan"] == 0:
            row_fills[i] = _NODATA_FILL

    row = _write_table(
        ws, row,
        ["Dosen Wali", "NIM", "Nama", "Jumlah MK Ditemukan", "Jumlah MK Gagal", "Daftar MK Gagal", "Catatan"],
        table_rows,
        [30, 16, 30, 18, 16, 60, 30],
        left_align_cols={1, 3, 6},
        row_fills=row_fills,
    )

    row += 1
    note = ws.cell(
        row=row, column=1,
        value=(
            'Baris merah = tidak lulus minimal 1 mata kuliah. Baris abu-abu = '
            '"Jumlah MK Ditemukan" = 0, artinya TIDAK ADA data nilai mahasiswa '
            "itu di folder data nilai yang dipilih (bukan berarti lulus semua) "
            "- kemungkinan datanya ada di folder lain yang belum diikutkan, "
            "atau mahasiswanya belum/tidak mengambil mata kuliah apa pun di "
            "semester yang dicek. Data ini juga belum memperhitungkan nilai "
            "remedial dan status mahasiswa yang sudah nonaktif - cek ulang "
            "manual sebelum dipakai sebagai dasar keputusan."
        ),
    )
    note.font = Font(italic=True, color="808080")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--folder", required=True, help="Folder berisi file Excel nilai per mata kuliah")
    parser.add_argument("--output", required=True, help="Path file Excel hasil")
    parser.add_argument("--perwalian", help="Folder berisi PDF data perwalian (opsional)")
    args = parser.parse_args()

    courses, skipped = generate(args.folder)

    perwalian, perwalian_skipped = None, []
    if args.perwalian:
        perwalian, perwalian_skipped = load_perwalian(args.perwalian)

    export_excel(courses, args.output, perwalian)

    print(f"Mata kuliah tergabung ({len(courses)}): {', '.join(c['course_name'] for c in courses)}")
    for name, msg in skipped:
        print(f"Dilewati - {name}: {msg}")
    for c in courses:
        if c["parse_error"]:
            print(f"Tidak dianalisis - {c['course_name']}: {c['parse_error']}")
    if args.perwalian:
        print(f"Data perwalian terbaca: {len(perwalian)} mahasiswa")
        for name, msg in perwalian_skipped:
            print(f"Dilewati (perwalian) - {name}: {msg}")
    print(f"Hasil: {args.output}")


if __name__ == "__main__":
    main()
