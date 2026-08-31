"""
Tool UI: Buat Folder Dosen Penasihat Akademik.
Wrapper generik (build_folder_creator_frame) di sekitar
core.buat_folder_dari_daftar.create_folders().
"""

from core import buat_folder_dari_daftar as folder_core
from tools.common import build_folder_creator_frame, data_dir

NAME = "Buat Folder Dosen Penasihat Akademik"
LABEL = "Buat Folder"
ICON = "\U0001F4C1"  # 📁
DESCRIPTION = (
    "Bikin folder BA Perwalian untuk banyak dosen sekaligus, tinggal tempel "
    "daftar namanya - tidak perlu buat satu-satu."
)
DEFAULT_PREFIX = "BA Perwalian Ganjil 2627-"
NAMES_LABEL = "Daftar Nama Dosen"


def build_frame(parent):
    return build_folder_creator_frame(
        parent,
        NAME,
        DESCRIPTION,
        DEFAULT_PREFIX,
        NAMES_LABEL,
        folder_core.create_folders,
        default_output_dir=data_dir("dosen-wali"),
    )
