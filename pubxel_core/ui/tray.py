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
from PyQt6.QtCore import QEvent, QObject, QPropertyAnimation, QThread, QTimer, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtGui import QAction, QIcon
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
from pubxel_core.ids import list_to_string, string_to_list
from pubxel_core.pubmed import obtain_pubmed_data
from pubxel_core.settings import save_settings, save_settings_key
from pubxel_core.ui.helpers import dialog_onebutton, dirname, open_directory, try_open_directory

class SystemTrayIcon(QSystemTrayIcon):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setIcon(QIcon(rt.icon_path))
        self.setToolTip("Pub-Xel [DEV]" if rt.developerMode else "Pub-Xel")
        if rt.os_name == "Windows":
            self.activated.connect(self.on_activated)

        self.menu = QMenu(parent)
        self.restore_action = QAction("Restore", self)
        self.exit_action = QAction("Exit", self)
        self.restore_action.triggered.connect(self.on_restore)
        self.exit_action.triggered.connect(self.on_exit)
        self.menu.addAction(self.restore_action)
        self.menu.addAction(self.exit_action)
        self.setContextMenu(self.menu)

    @pyqtSlot(QSystemTrayIcon.ActivationReason)
    def on_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason in (
            QSystemTrayIcon.ActivationReason.Trigger,
            QSystemTrayIcon.ActivationReason.DoubleClick,
        ):
            self.on_restore()

    def on_restore(self):
        if self.parent().isMinimized():
            self.parent().showNormal()
        self.parent().show()
        self.parent().activateWindow()

    def on_exit(self):
        self.parent().close_application()
