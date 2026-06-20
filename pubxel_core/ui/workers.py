# Pub-Xel - A Biomedical Reference Management Tool
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

from pubxel_core import runtime as rt
from pubxel_core.excel_ops import check_file_exist, copy_list, files_name_to_path, process_ids
from pubxel_core.pubmed import input_pubmed_data
from pubxel_core.pubmed import obtain_pubmed_data
from pubxel_core.settings import save_settings, save_settings_key
from pubxel_core.ui.helpers import dialog_onebutton, dirname, open_directory, try_open_directory

class listenerWorker(QObject):
    signal_inspect = pyqtSignal()
    signal_open = pyqtSignal()
    signal_pubmed = pyqtSignal()
    def __init__(self):
        super().__init__()
    def run_inspect(self):
        self.signal_inspect.emit()
    def run_open(self):
        self.signal_open.emit()
    def run_pubmed(self):
        self.signal_pubmed.emit()

class excelWorker(QThread):
    excel_updated = pyqtSignal(str)
    def run(self):
        text = ""
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future = executor.submit(self.process_excel)
            try:
                text = future.result(timeout=5)  # 5 seconds timeout
                self.excel_updated.emit(text)
            except concurrent.futures.TimeoutError:
                print("Processing excel data timed out")
    def process_excel(self):
        try:
            wb, rng =None, None
            try:
                wb = xw.books.active
                ws = xw.sheets.active
                rng = wb.app.selection
                used_range = ws.used_range
                rng = ws.range((rng.row, rng.column), 
                            (rng.rows[-1].row, rng.columns[-1].column))
                if(rng.row < used_range.row):
                    rng = ws.range((used_range.row, rng.column), 
                            (rng.rows[-1].row, rng.columns[-1].column))
                if(rng.rows[-1].row > used_range.rows[-1].row):
                    rng = ws.range((rng.row, rng.column), 
                            (used_range.rows[-1].row, rng.columns[-1].column))
                if(rng.column < used_range.column):
                    rng = ws.range((rng.row, used_range.column), 
                            (rng.rows[-1].row, rng.columns[-1].column))
                if(rng.columns[-1].column > used_range.columns[-1].column):
                    rng = ws.range((rng.row, rng.column), 
                            (rng.rows[-1].row, used_range.columns[-1].column))
            except Exception as e:
                print(f"Error accessing workbook or range: {e}")
            if wb is not None and rng is not None:
                cellcount = rng.count
                text = f"{cellcount} Cell{'s' if cellcount > 1 else ''} in: {os.path.basename(wb.fullname)}."
            else: 
                text = ""
            return text
        except Exception as e:
            print(f"Error accessing clipboard: {e}")
        return ""


def check_shortcut(worker):

    hotkey_inspect_value = rt.settings.get("hotkey_inspect_value",0)
    hotkey_open_value = rt.settings.get("hotkey_open_value",0)
    hotkey_pubmed_value = rt.settings.get("hotkey_pubmed_value",0)

    # rt.action_in_progress is read here from the pynput listener thread to
    # gate hotkey dispatch. A bool read is atomic under the GIL; the actual
    # check-and-set happens on the Qt main thread inside the slot, via
    # rt.try_begin_action().
    def on_activate_inspect():
        if not rt.action_in_progress:
            print("on_activate_inspect activated")
            worker.run_inspect()
        else:
            print("on_activate_inspect but action in progress")

    def on_activate_open():
        if not rt.action_in_progress:
            print("on_activate_open activated")
            worker.run_open()
        else:
            print("on_activate_open but action in progress")

    def on_activate_pubmed():
        if not rt.action_in_progress:
            print("on_activate_pubmed activated")
            worker.run_pubmed()
        else:
            print("on_activate_pubmed but action in progress")

    if hotkey_inspect_value == "" and hotkey_open_value == "" and hotkey_pubmed_value == "":
        return
    if not hotkey_inspect_value == "":
        hotkey_j = keyboard.HotKey(keyboard.HotKey.parse(hotkey_inspect_value), on_activate_inspect)
    if not hotkey_open_value == "":
        hotkey_k = keyboard.HotKey(keyboard.HotKey.parse(hotkey_open_value), on_activate_open)
    if not hotkey_pubmed_value == "":
        hotkey_p = keyboard.HotKey(keyboard.HotKey.parse(hotkey_pubmed_value), on_activate_pubmed)


    def for_canonical(f):
        return lambda k: f(rt.listeners[0].canonical(k))

    def on_press(key):
        if not hotkey_inspect_value == "":
            hotkey_j.press(key)
        if not hotkey_open_value == "":
            hotkey_k.press(key)
        if not hotkey_pubmed_value == "":
            hotkey_p.press(key)
        return

    def on_release(key):
        if not hotkey_inspect_value == "":
            hotkey_j.release(key)
        if not hotkey_open_value == "":
            hotkey_k.release(key)
        if not hotkey_pubmed_value == "":
            hotkey_p.release(key)
        return

    def start_listener():
        try:
            listener = keyboard.Listener(
                on_press=for_canonical(on_press),
                on_release=for_canonical(on_release))
            rt.listeners.append(listener)
            listener.start()
            print("Listener started successfully")
        except Exception as e:
            print(f"Error starting listener: {e}")

    print("listener_thread = threading.Thread(target=start_listener)")
    listener_thread = threading.Thread(target=start_listener)
    print("listener_thread.daemon = True")
    listener_thread.daemon = True  # Ensure the thread exits when the main program exits
    print("listener_thread.start()")
    listener_thread.start()
