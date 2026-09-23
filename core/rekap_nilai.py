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

Cara pakai (dari folder ini):
    python rekap_nilai.py --folder "path/ke/folder/nilai" --output "hasil.xlsx"
"""
import argparse
import copy
import os

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


def generate(folder_path):
    """
    Parse semua file di folder_path. Kembalikan (courses, skipped):
    courses = list dict hasil parse_course_file(), terurut abjad berdasarkan
    nama mata kuliah; skipped = list (nama_file, pesan_error) untuk file yang
    GAGAL DIBUKA sama sekali (bukan .xlsx/.xls valid - beda dari
    parse_error, yang masih menghasilkan course dengan sheet tetap disalin).
    """
    paths = scan_folder(folder_path)
    if not paths:
        raise ValueError("Tidak ada file Excel (.xls/.xlsx) yang ditemukan di folder ini.")

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


def _safe_sheet_name(name, used):
    name = "".join(ch for ch in name if ch not in r":\/?*[]").strip() or "Mata Kuliah"
    base = name[:31]
    candidate = base
    n = 2
    while candidate.lower() in used:
        suffix = f" ({n})"
        candidate = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate


def export_excel(courses, output_path):
    wb = Workbook()
    wb.remove(wb.active)
    ws_ringkasan = wb.create_sheet("Ringkasan", 0)

    used_names = set()
    for course in courses:
        title = _safe_sheet_name(course["course_name"], used_names)
        _copy_sheet(course["worksheet"], wb, title)

    _write_ringkasan(ws_ringkasan, courses)
    wb.save(output_path)


_HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")
_FAIL_FILL = PatternFill("solid", fgColor="FCE4E4")
_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center")


def _write_table(ws, start_row, headers, rows, col_widths, highlight_rows=None):
    highlight_rows = highlight_rows or set()
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
            cell.alignment = _CENTER
            if i in highlight_rows:
                cell.fill = _FAIL_FILL
    for c, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = width
    return start_row + 1 + len(rows)


def _write_ringkasan(ws, courses):
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
                course["course_name"], course["dosen"] or "",
                "", "", "", f"Tidak dianalisis - {course['parse_error']}",
            ))
            continue
        total = len(course["students"])
        lulus = sum(1 for s in course["students"] if s["status"] == LULUS)
        dinilai = sum(1 for s in course["students"] if s["status"])
        tidak_lulus = dinilai - lulus
        persen = f"{lulus / dinilai * 100:.0f}%" if dinilai else ""
        course_rows.append((
            course["course_name"], course["dosen"] or "",
            total, lulus, tidak_lulus, persen,
        ))

    row = _write_table(
        ws, row,
        ["Mata Kuliah", "Dosen", "Jumlah Mahasiswa", "Lulus", "Tidak Lulus", "% Lulus"],
        course_rows,
        [30, 30, 16, 10, 12, 24],
    )

    row += 2
    ws.cell(row=row, column=1, value="Mahasiswa Tidak Lulus (minimal 1 mata kuliah)").font = Font(bold=True, size=12)
    row += 1

    failing = _failing_students(courses)
    if failing:
        # Kolom "Catatan" sengaja dikosongkan - dipakai dosen wali menandai
        # tindak lanjut per mahasiswa (mis. "sudah dihubungi", "remedial"),
        # bukan diisi otomatis oleh aplikasi.
        fail_rows = [(f["nim"], f["nama"], f["jumlah_gagal"], f["daftar_gagal"], "") for f in failing]
        row = _write_table(
            ws, row,
            ["NIM", "Nama", "Jumlah MK Gagal", "Mata Kuliah yang Gagal", "Catatan"],
            fail_rows,
            [16, 30, 16, 60, 30],
            highlight_rows=set(range(len(fail_rows))),
        )
    else:
        ws.cell(row=row, column=1, value="Tidak ada mahasiswa yang tidak lulus di mata kuliah manapun.")
        row += 1

    row += 1
    note = ws.cell(
        row=row, column=1,
        value=(
            "Catatan: data ini belum memperhitungkan nilai remedial, dan belum "
            "mengecualikan mahasiswa yang sudah keluar/nonaktif di semester "
            "berikutnya - cek ulang manual untuk kasus-kasus tersebut sebelum "
            "dipakai sebagai dasar keputusan."
        ),
    )
    note.font = Font(italic=True, color="808080")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--folder", required=True, help="Folder berisi file Excel nilai per mata kuliah")
    parser.add_argument("--output", required=True, help="Path file Excel hasil")
    args = parser.parse_args()

    courses, skipped = generate(args.folder)
    export_excel(courses, args.output)

    print(f"Mata kuliah tergabung ({len(courses)}): {', '.join(c['course_name'] for c in courses)}")
    for name, msg in skipped:
        print(f"Dilewati - {name}: {msg}")
    for c in courses:
        if c["parse_error"]:
            print(f"Tidak dianalisis - {c['course_name']}: {c['parse_error']}")
    print(f"Hasil: {args.output}")


if __name__ == "__main__":
    main()
