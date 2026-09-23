"""
Tools: Rekap nilai per dosen wali, mulai dari roster perwalian (bukan dari
file nilai). Untuk SEMUA mahasiswa di roster perwalian, cari nilainya di
seluruh file nilai per mata kuliah (format "Evaluasi Ketercapaian MK"/OBE
hasil export sistem akademik) yang diberikan - boleh gabungan beberapa
folder/semester sekaligus.

Hasilnya 1 file Excel: 1 sheet "Ringkasan" (mahasiswa dikelompokkan per
dosen wali, yang gagal paling banyak ditaruh duluan), ditambah 1 sheet
gabungan per NAMA mata kuliah (semua kelas paralel yang mengajarkannya
digabung jadi 1 sheet, dengan kolom Kelas & Dosen Pengajar supaya konteks
tiap baris tetap jelas).

Tiap mata kuliah boleh beda jumlah Sub-CPMK/kolom penilaian (makanya lebar
tabel sumbernya beda-beda per file) - kolom yang dibutuhkan (NIM, Nama
Mahasiswa, Status Kelulusan MK) dicari lewat teks header-nya, bukan posisi
kolom tetap, supaya tetap kebaca walau jumlah kolom di antaranya berubah.
Kalau suatu file nilai formatnya tidak dikenali sama sekali, mahasiswa di
mata kuliah itu dianggap belum ada datanya (bukan dianggap lulus) - lihat
catatan di sheet Ringkasan. Kalau file PDF perwalian formatnya tidak
dikenali, dilewati dengan pesan jelas (tidak menggagalkan yang lain).

Opsional: --per-dosen <folder> sekaligus memecah hasilnya jadi 1 file
Excel terpisah per dosen wali di folder itu (cuma berisi anak wali dosen
itu sendiri) - siap dikirim langsung ke masing-masing dosen wali.

Cara pakai (dari folder ini):
    python rekap_nilai.py --perwalian "path/ke/folder/perwalian" \\
        --folder "path/ke/folder/nilai1" "path/ke/folder/nilai2" \\
        --output "hasil.xlsx" [--per-dosen "path/ke/folder/output"]
"""
import argparse
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
    tetap ikut nama mata kuliah hasil parsing)."""
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


def load_perwalian_roster(folder_path):
    """
    Baca semua PDF di folder_path. Kembalikan (roster, skipped): roster =
    {nim: {"nama", "dosen", "kelas", "semester"}} (satu entri per
    mahasiswa), skipped = list (nama_file, pesan_error) untuk PDF yang
    formatnya tidak dikenali (dilewati, tidak menggagalkan yang lain).
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
    Excel yang valid - metadata (nama MK, kelas, dosen) tetap dikembalikan
    walau tabel mahasiswanya gagal dikenali; dalam kasus itu "students"
    kosong dan "parse_error" berisi alasannya.
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
    Parse semua file Excel di seluruh folder_paths sekaligus (boleh lebih
    dari 1 folder/semester). Kembalikan (courses, skipped): courses = list
    dict hasil parse_course_file(), terurut abjad berdasarkan nama mata
    kuliah; skipped = list (nama_file, pesan_error) untuk file yang GAGAL
    DIBUKA sama sekali (bukan .xlsx/.xls valid - beda dari parse_error,
    yang masih menghasilkan course, cuma "students"-nya kosong).
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
    generate_from_folders(), boleh gabungan banyak folder/semester).
    Kembalikan list baris terurut: per dosen wali (abjad), lalu per
    kelompok kelas+semester (satu dosen bisa jadi wali >1 kelompok
    sekaligus - mis. kelas peminatan di semester lain - jadi tiap
    kelompoknya tetap dipisah, bukan digabung rata ke 1 dosen; semester
    WAJIB ikut jadi pembeda, bukan cuma kelas, karena field "Kelas" di
    PDF-nya sendiri kadang sama persis walau semesternya beda, mis. kelas
    "IF/B" dipakai baik di semester 5 (AISD-B Sore) maupun semester 7
    (IF-B Sore) - dua kelompok mahasiswa yang sama sekali beda orangnya),
    lalu di dalam tiap kelompok - yang gagal >=1 MK dulu (paling banyak
    gagal duluan), baru yang lulus semua, baru yang datanya belum ketemu
    sama sekali.
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
            "kelas": info["kelas"] or "",
            "semester": info["semester"] or "",
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

    def sem_key(r):
        # Tuple (0, angka) vs (1, teks) - supaya tidak pernah membandingkan
        # int dengan str langsung kalau semester-nya kosong/bukan angka
        # (TypeError di Python 3).
        try:
            return (0, int(r["semester"]))
        except (TypeError, ValueError):
            return (1, r["semester"] or "")

    rows.sort(key=lambda r: (r["dosen"], r["kelas"], sem_key(r), tier(r), -r["jumlah_gagal"], r["nim"]))
    return rows


def export_excel(roster, courses, output_path):
    wb = Workbook()
    wb.remove(wb.active)
    ws_ringkasan = wb.create_sheet("Ringkasan", 0)

    # 1 sheet per NAMA mata kuliah (gabungan semua kelas paralel yang
    # mengajarkannya) - dosen wali cuma perlu tahu status kelulusan tiap
    # anak walinya, bukan rincian Sub-CPMK mentah; kolom Kelas & Dosen
    # Pengajar menjaga konteks asal tiap baris tetap kelihatan.
    by_course_name = {}
    for course in courses:
        by_course_name.setdefault(course["course_name"], []).append(course)

    used_names = set()
    for course_name in sorted(by_course_name, key=str.lower):
        _write_course_summary_sheet(wb, course_name, by_course_name[course_name], used_names)

    rows = generate_perwalian_rekap(roster, courses)
    _write_ringkasan(ws_ringkasan, rows)
    wb.save(output_path)


def _write_course_summary_sheet(wb, course_name, course_list, used_names):
    """1 sheet gabungan untuk 1 nama mata kuliah, isinya mahasiswa dari
    SEMUA kelas paralel yang mengajarkan mata kuliah itu."""
    title_base = "".join(ch for ch in course_name if ch not in r":\/?*[]").strip() or "Mata Kuliah"
    base = title_base[:31]
    title, n = base, 2
    while title.lower() in used_names:
        suffix = f" ({n})"
        title = base[: 31 - len(suffix)] + suffix
        n += 1
    used_names.add(title.lower())

    ws = wb.create_sheet(title)

    rows = []
    for course in course_list:
        if course["parse_error"]:
            rows.append((None, None, course["kelas"] or "", course["dosen"] or "",
                         None, f"Tidak dianalisis - {course['parse_error']}"))
            continue
        for s in course["students"]:
            rows.append((
                s["nim"], s["nama"], course["kelas"] or "", course["dosen"] or "",
                s["nilai_huruf"] if s["nilai_huruf"] not in (None, "") else "",
                s["status"] or "",
            ))
    rows.sort(key=lambda r: ((r[2] or ""), (r[0] or "")))

    _write_table(
        ws, 1,
        ["NIM", "Nama", "Kelas", "Dosen Pengajar", "Nilai Huruf MK", "Status Kelulusan MK"],
        rows,
        [16, 30, 16, 30, 16, 24],
        left_align_cols={1, 2, 3},
    )


_HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
_FAIL_FILL = PatternFill("solid", fgColor="FCE4E4")
_NODATA_FILL = PatternFill("solid", fgColor="F2F2F2")
_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center")
_LEFT = Alignment(horizontal="left")


def _write_table(ws, start_row, headers, rows, col_widths, left_align_cols=None, row_fills=None):
    """row_fills: {row_index: PatternFill} - fill per baris data (0-based,
    relatif ke rows)."""
    left_align_cols = left_align_cols or set()
    row_fills = row_fills or {}
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


def _write_ringkasan(ws, rows, title="Ringkasan Perwalian", show_dosen_col=True):
    """show_dosen_col=False dipakai buat file per-dosen-wali terpisah (lihat
    export_per_advisor_excel) - kolom "Dosen Wali" jadi mubazir kalau
    seisi file memang cuma anak wali 1 dosen itu saja."""
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=14)

    if show_dosen_col:
        dosen_set = sorted({r["dosen"] for r in rows if r["dosen"]})
        kelompok_set = sorted({(r["dosen"], r["kelas"], r["semester"]) for r in rows if r["dosen"]})
        ws.cell(row=2, column=1, value=f"Jumlah Dosen Wali (unik): {len(dosen_set)}")
        ws.cell(row=3, column=1, value=f"Jumlah Kelompok Perwalian (dosen+kelas+semester): {len(kelompok_set)}")
        ws.cell(row=4, column=1, value=f"Jumlah Mahasiswa (roster perwalian): {len(rows)}")
        row = 6
        # Kolom "Kelas" & "Semester" penting: satu dosen bisa jadi wali
        # >1 kelompok sekaligus (mis. kelas peminatan di semester lain) -
        # dan field "Kelas" di PDF-nya sendiri kadang SAMA PERSIS walau
        # semesternya beda (mis. "IF/B" dipakai baik di semester 5 maupun
        # 7, dua kelompok mahasiswa yang sama sekali beda) - tanpa kedua
        # kolom ini, baris "Dosen Wali" yang sama berulang tanpa cara
        # membedakan kelompoknya masing-masing.
        headers = ["Dosen Wali", "Kelas", "Semester", "NIM", "Nama", "Jumlah MK Ditemukan", "Jumlah MK Gagal", "Daftar MK Gagal", "Catatan"]
        table_rows = [
            (r["dosen"], r["kelas"], r["semester"], r["nim"], r["nama"], r["jumlah_ditemukan"], r["jumlah_gagal"], r["daftar_gagal"], "")
            for r in rows
        ]
        widths = [30, 16, 12, 16, 30, 18, 16, 60, 30]
        left_cols = {1, 2, 5, 8}
    else:
        ws.cell(row=2, column=1, value=f"Jumlah Mahasiswa: {len(rows)}")
        row = 4
        headers = ["NIM", "Nama", "Jumlah MK Ditemukan", "Jumlah MK Gagal", "Daftar MK Gagal", "Catatan"]
        table_rows = [
            (r["nim"], r["nama"], r["jumlah_ditemukan"], r["jumlah_gagal"], r["daftar_gagal"], "")
            for r in rows
        ]
        widths = [16, 30, 18, 16, 60, 30]
        left_cols = {1, 4}

    # Kolom "Catatan" sengaja dikosongkan - dipakai dosen wali menandai
    # tindak lanjut per mahasiswa (mis. "sudah dihubungi", "remedial"),
    # bukan diisi otomatis oleh aplikasi.
    row_fills = {}
    for i, r in enumerate(rows):
        if r["jumlah_gagal"] > 0:
            row_fills[i] = _FAIL_FILL
        elif r["jumlah_ditemukan"] == 0:
            row_fills[i] = _NODATA_FILL

    row = _write_table(ws, row, headers, table_rows, widths, left_align_cols=left_cols, row_fills=row_fills)

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


def _safe_filename(name):
    name = "".join(ch for ch in (name or "").strip() if ch not in '\\/:*?"<>|').strip()
    return name or "Tanpa Dosen Wali"


def export_per_advisor_excel(roster, courses, output_dir):
    """
    Tulis 1 file Excel TERPISAH per (dosen wali, kelas, semester) ke
    output_dir (dibuat kalau belum ada) - satu dosen bisa jadi wali lebih
    dari 1 kelompok sekaligus (mis. kelas peminatan di semester lain),
    jadi tiap kelompoknya dipecah jadi file sendiri-sendiri (BUKAN
    digabung rata ke 1 dosen). Semester WAJIB ikut jadi pembeda nama file,
    bukan cuma kelas, karena field "Kelas" di PDF-nya sendiri kadang sama
    persis walau semesternya beda (lihat catatan di generate_perwalian_rekap).
    Kembalikan list path file yang berhasil dibuat, terurut nama dosen,
    lalu kelas, lalu semester (abjad/numerik).
    """
    rows = generate_perwalian_rekap(roster, courses)
    by_group = {}
    for r in rows:
        key = (r["dosen"] or "(Tanpa Dosen Wali)", r["kelas"] or "(Tanpa Kelas)", r["semester"] or "")
        by_group.setdefault(key, []).append(r)

    def group_sort_key(k):
        dosen, kelas, semester = k
        try:
            sem_num = (0, int(semester))
        except (TypeError, ValueError):
            sem_num = (1, semester)
        return (dosen.lower(), kelas.lower(), sem_num)

    os.makedirs(output_dir, exist_ok=True)
    written = []
    used_names = set()
    for dosen, kelas, semester in sorted(by_group, key=group_sort_key):
        wb = Workbook()
        ws = wb.active
        ws.title = "Ringkasan"
        label = f"{dosen} - {kelas} (Semester {semester})" if semester else f"{dosen} - {kelas}"
        _write_ringkasan(ws, by_group[(dosen, kelas, semester)], title=f"Rekap Nilai - {label}", show_dosen_col=False)

        filename_base = _safe_filename(label)
        filename = filename_base
        n = 2
        while filename.lower() in used_names:
            filename = f"{filename_base} ({n})"
            n += 1
        used_names.add(filename.lower())

        path = os.path.join(output_dir, f"Rekap Nilai - {filename}.xlsx")
        wb.save(path)
        written.append(path)
    return written


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--perwalian", required=True, help="Folder berisi PDF data perwalian")
    parser.add_argument("--folder", nargs="+", required=True, help="Satu atau lebih folder berisi file Excel nilai per mata kuliah")
    parser.add_argument("--output", required=True, help="Path file Excel hasil")
    parser.add_argument("--per-dosen", help="Folder tujuan - kalau diisi, pecah juga hasilnya jadi 1 file per dosen wali")
    args = parser.parse_args()

    roster, roster_skipped = load_perwalian_roster(args.perwalian)
    courses, skipped = generate_from_folders(args.folder)
    export_excel(roster, courses, args.output)

    print(f"Data perwalian terbaca: {len(roster)} mahasiswa")
    for name, msg in roster_skipped:
        print(f"Dilewati (perwalian) - {name}: {msg}")
    print(f"Mata kuliah tergabung ({len(courses)}): {', '.join(c['course_name'] for c in courses)}")
    for name, msg in skipped:
        print(f"Dilewati (nilai) - {name}: {msg}")
    for c in courses:
        if c["parse_error"]:
            print(f"Tidak dianalisis - {c['course_name']}: {c['parse_error']}")
    print(f"Hasil: {args.output}")

    if args.per_dosen:
        written = export_per_advisor_excel(roster, courses, args.per_dosen)
        print(f"Dipecah jadi {len(written)} file per dosen wali di: {args.per_dosen}")


if __name__ == "__main__":
    main()
