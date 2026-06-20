# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>

import os
import platform
import shlex
import subprocess

from PyQt6 import QtCore
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QMessageBox, QPushButton, QWidget

from pubxel_core import runtime as rt
from pubxel_core.recent_worksheets import register_recent_worksheet
from pubxel_core.settings import save_settings_key


EXCEL_UNAVAILABLE_HINT = (
    "\n\nThis may be because a full version of Microsoft Excel is not available. "
    "Please check that Excel is installed and licensed."
)

# User-facing guidance from Excel selection/validation — not an Excel install issue.
_EXCEL_USER_GUIDANCE_PREFIXES = (
    "Please open the Excel Worksheet first.",
    "No selection made",
    "Please select ",
    "Error: Please ensure the following",
    "Error: Multiple 'ref' columns",
    "No worksheet columns selected",
    "No PubMed ID",
    "No PubMed metadata",
    "Template not found",
    "Tutorial worksheet template not found",
    "Template is missing expected column",
    "Template worksheet must contain",
)


def format_excel_operation_error(exc: BaseException | str) -> str:
    """Append Excel availability guidance for likely COM / xlwings failures."""
    message = str(exc).strip() or exc.__class__.__name__
    if isinstance(exc, FileNotFoundError):
        return message
    if any(message.startswith(prefix) for prefix in _EXCEL_USER_GUIDANCE_PREFIXES):
        return message
    if EXCEL_UNAVAILABLE_HINT.strip() in message:
        return message
    return message + EXCEL_UNAVAILABLE_HINT


def _get_sep(p):
    if isinstance(p, bytes):
        return b"\\" if os.name == "nt" else b"/"
    return "\\" if os.name == "nt" else "/"


def dirname(p):
    p = os.fspath(p)
    sep = _get_sep(p)
    i = p.rfind(sep) + 1
    head = p[:i]
    if head and head != sep * len(head):
        head = head.rstrip(sep)
    return head


def open_directory(dir):
    os_name = platform.system()
    if os_name == "Windows":
        os.startfile(dir)
    elif os_name == "Darwin":
        subprocess.call(
            f"open {shlex.quote(dir)}",
            shell=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,
        )


def try_open_directory(dir):
    try:
        open_directory(dir)
    except Exception as e:
        print(f"Cannot open: {e}")


def stop_listeners():
    for listener in rt.listeners:
        listener.stop()
    rt.listeners = []


def graceful_shutdown():
    try:
        stop_listeners()
    except Exception:
        pass
    try:
        if rt.lock_file:
            rt.lock_file.close()
    except Exception:
        pass
    try:
        if rt.lock_file_path and os.path.exists(rt.lock_file_path):
            os.remove(rt.lock_file_path)
    except Exception:
        pass


WORKSHEET_DEFAULT_BASENAME = "Pub-Xel Worksheet"
TSV_DEFAULT_BASENAME = "Pub-Xel data"


def default_worksheet_save_directory() -> str:
    """Return the default save/browse directory (~/Documents, fallback ~)."""
    save_dir = os.path.expanduser("~/Documents")
    if os.path.isdir(save_dir):
        return save_dir
    return os.path.expanduser("~")


def _default_save_name(directory: str, base_name: str, suffix: str) -> str:
    dir_path = os.path.abspath(directory)
    first = f"{base_name}{suffix}"
    if not os.path.exists(os.path.join(dir_path, first)):
        return first
    n = 2
    while True:
        candidate = f"{base_name} ({n}){suffix}"
        if not os.path.exists(os.path.join(dir_path, candidate)):
            return candidate
        n += 1


def default_worksheet_save_name(directory: str, base_name: str = WORKSHEET_DEFAULT_BASENAME) -> str:
    """Return the next unused default worksheet filename in ``directory``."""
    return _default_save_name(directory, base_name, ".xlsx")


def default_tsv_save_name(directory: str, base_name: str = TSV_DEFAULT_BASENAME) -> str:
    """Return the next unused default TSV filename in ``directory``."""
    return _default_save_name(directory, base_name, ".tsv")


def dialog_onebutton(parent, message, title="Confirmation"):
    msg_box = QMessageBox(parent)
    msg_box.setWindowTitle(title)
    msg_box.setText(message)
    msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg_box.setWindowModality(Qt.WindowModality.ApplicationModal)
    msg_box.exec()
    msg_box.setFocus(QtCore.Qt.FocusReason.PopupFocusReason)


def show_file_saved_dialog(
    parent: QWidget | None,
    file_path: str,
    *,
    open_label: str = "Open File",
    saved_message: str = "File saved.",
    register_recent: bool = False,
    increment_worksheet_count: bool = False,
) -> None:
    """Show a post-save dialog with optional open file/folder actions."""
    msg_box = QMessageBox(parent)
    msg_box.setIcon(QMessageBox.Icon.Information)
    msg_box.setText(saved_message)
    msg_box.setWindowTitle("Success")

    open_button = QPushButton(open_label)
    msg_box.addButton(open_button, QMessageBox.ButtonRole.ActionRole)

    open_folder_button = None
    if rt.settings.get("worksheet_count", 0) > 0:
        open_folder_button = QPushButton("Open Folder")
        msg_box.addButton(open_folder_button, QMessageBox.ButtonRole.ActionRole)
        msg_box.addButton(QMessageBox.StandardButton.Ok)

    msg_box.exec()

    if msg_box.clickedButton() == open_button:
        try_open_directory(file_path)

    if open_folder_button is not None and msg_box.clickedButton() == open_folder_button:
        folder_path = os.path.dirname(file_path)
        try_open_directory(folder_path)

    if increment_worksheet_count:
        rt.settings = save_settings_key(
            rt.settings, "worksheet_count", rt.settings.get("worksheet_count", 0) + 1
        )
    if register_recent:
        register_recent_worksheet(file_path)


def show_worksheet_saved_dialog(parent: QWidget | None, file_path: str) -> None:
    show_file_saved_dialog(
        parent,
        file_path,
        open_label="Open Worksheet",
        saved_message="Worksheet saved.",
        register_recent=True,
        increment_worksheet_count=True,
    )
