"""One-off helper: extract UI sections from Pub-Xel.py into pubxel_core/ui/."""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "Pub-Xel.py"
OUT = ROOT / "pubxel_core" / "ui"

COMMON_HEADER = '''# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>

import copy
import datetime
import os
import platform
import re
import shutil
import threading
import webbrowser

import concurrent.futures
import xlwings as xw
from pynput import keyboard
from PyQt6 import QtCore, uic
from PyQt6.QtCore import QEvent, QObject, QPropertyAnimation, QThread, QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QFontMetrics, QIcon, QKeySequence, QPixmap, QShortcut, QTextCursor
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpacerItem,
    QSystemTrayIcon,
    QVBoxLayout,
    QWidget,
)

from data.version import __version__
from pubxel_core import runtime as rt
from pubxel_core.excel_ops import check_file_exist, copy_list, files_name_to_path, process_ids
from pubxel_core.pubmed import input_pubmed_data
from pubxel_core.ids import list_to_string, string_to_list
from pubxel_core.pubmed import obtain_pubmed_data
from pubxel_core.settings import save_settings, save_settings_key
from pubxel_core.ui.helpers import dialog_onebutton, dirname, open_directory, try_open_directory

'''

SECTIONS = {
    "helpers.py": (233, 278),
    "preferences.py": (280, 718),
    "widgets.py": (720, 1347),
    "workers.py": (1349, 1511),
    "tray.py": (1514, 1542),
    "dialogs_extra.py": (1544, 1641),
    "main_window.py": (1643, 2208),
}

RUNTIME_NAMES = {
    "settings", "mainlibdir", "seclibdir", "outdir", "os_name",
    "mainlibdirdefault", "outdirdefault", "system_tray_notice_shown",
    "developerMode", "action_in_progress", "worksheetColumns_in_progress",
    "force_quit", "listeners", "app",
    "main_path", "inspect_path", "about_path", "worksheetColumns_path",
    "preferences_path", "icon_path", "questionmark_icon_path", "loading_image_path",
    "pubsheet_path", "pubsheetinitial_path",
}

def transform(body: str) -> str:
    body = re.sub(r"^\s*global (\w+)\s*$", r"# global \1 -> rt.\1", body, flags=re.MULTILINE)
    for name in sorted(RUNTIME_NAMES, key=len, reverse=True):
        body = re.sub(rf"(?<!\.)\b{name}\b", f"rt.{name}", body)
    return body


def main():
    lines = SRC.read_text(encoding="utf-8").splitlines(keepends=True)
    OUT.mkdir(parents=True, exist_ok=True)

    for fname, (start, end) in SECTIONS.items():
        body = "".join(lines[start - 1 : end])
        if fname == "helpers.py":
            header = COMMON_HEADER.replace(
                "from pubxel_core.ui.helpers import dialog_onebutton, dirname, open_directory, try_open_directory\n\n",
                "",
            )
        else:
            header = COMMON_HEADER
        content = header + transform(body)
        (OUT / fname).write_text(content, encoding="utf-8")
        print(f"wrote {fname} ({end - start + 1} lines)")


if __name__ == "__main__":
    main()
