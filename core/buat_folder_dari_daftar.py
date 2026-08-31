"""
Tools: Buat satu folder per baris nama dari daftar nama (satu nama per
baris, baris kosong diabaikan). Nama folder = "<prefix><nama baris>".

Dipakai untuk kebutuhan yang bentuknya "buat banyak folder dari daftar
nama", mis. folder BA Perwalian per dosen penasihat akademik, atau folder
per mata kuliah - hanya beda prefix & lokasi output.

Cara pakai (dari folder ini):
    python buat_folder_dari_daftar.py --file "daftar.txt" --output "folder_tujuan" --prefix "Awalan-"
"""

import argparse
import os


def parse_names(text):
    """Pecah teks multi-baris jadi daftar nama; baris kosong diabaikan."""
    return [line.strip() for line in text.splitlines() if line.strip()]


def create_folders(names, output_dir, prefix):
    """Buat satu folder per nama di output_dir, nama folder = f'{prefix}{nama}'."""
    if not names:
        raise ValueError("Belum ada nama yang diisi.")

    os.makedirs(output_dir, exist_ok=True)

    created, existing = [], []
    for name in names:
        folder_name = f"{prefix}{name}"
        path = os.path.join(output_dir, folder_name)
        if os.path.exists(path):
            existing.append(folder_name)
        else:
            os.makedirs(path)
            created.append(folder_name)

    return created, existing


def main():
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument("--file", required=True, help="Path file teks daftar nama (satu nama per baris)")
    parser.add_argument("--output", required=True, help="Folder induk tempat folder per-nama dibuat")
    parser.add_argument("--prefix", default="", help="Teks yang ditempel di depan tiap nama untuk jadi nama folder")
    args = parser.parse_args()

    with open(args.file, encoding="utf-8-sig") as f:
        names = parse_names(f.read())

    created, existing = create_folders(names, args.output, args.prefix)
    print(f"Folder dibuat : {len(created)}")
    print(f"Sudah ada     : {len(existing)}")
    print(f"Lokasi        : {args.output}")


if __name__ == "__main__":
    main()
