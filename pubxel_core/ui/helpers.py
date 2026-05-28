# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>

import os
import platform
import shlex
import subprocess

from PyQt6 import QtCore
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QMessageBox

from pubxel_core import runtime as rt


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


def dialog_onebutton(parent, message, title="Confirmation"):
    msg_box = QMessageBox(parent)
    msg_box.setWindowTitle(title)
    msg_box.setText(message)
    msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg_box.setWindowModality(Qt.WindowModality.ApplicationModal)
    msg_box.exec()
    msg_box.setFocus(QtCore.Qt.FocusReason.PopupFocusReason)
