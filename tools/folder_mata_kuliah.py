"""
Tool UI: Buat Folder Mata Kuliah.
Wrapper generik (build_folder_creator_frame) di sekitar
core.buat_folder_dari_daftar.create_folders().
"""

from core import buat_folder_dari_daftar as folder_core
from tools.common import build_folder_creator_frame, data_dir

NAME = "Buat Folder Mata Kuliah"
LABEL = "Buat Folder"
ICON = "\U0001F4DA"  # 📚
DESCRIPTION = (
    "Bikin folder per mata kuliah (atau rumpun mata kuliah/RMK) sekaligus, "
    "tinggal tempel daftar namanya - tidak perlu buat satu-satu."
)
DEFAULT_PREFIX = "[IF-Ganjil-2627]-"
NAMES_LABEL = "Daftar Mata Kuliah"


def build_frame(parent):
    return build_folder_creator_frame(
        parent,
        NAME,
        DESCRIPTION,
        DEFAULT_PREFIX,
        NAMES_LABEL,
        folder_core.create_folders,
        default_output_dir=data_dir("mata-kuliah"),
    )
