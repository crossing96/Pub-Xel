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
from pubxel_core.ids import list_to_string, string_to_list
from pubxel_core.nbib import NBIB_EXCEL_MAX_ARTICLES
from pubxel_core.pubmed import input_pubmed_data, obtain_pubmed_data
from pubxel_core.settings import save_settings, save_settings_key
from pubxel_core.ui.helpers import dialog_onebutton, dirname, open_directory, try_open_directory

class PopupWidgettest(QWidget):
    popup_count = 0  # Class variable to keep track of the number of popups

    def __init__(self,texts_with_links=None):
        super().__init__()
        PopupWidgettest.popup_count += 1  # Increment the count each time a PopupWidgettest is created
        self.setWindowFlags(Qt.WindowType.Popup)
        self.setLayout(QVBoxLayout())
        self.layout().addWidget(QLabel(f"Total Popups: {PopupWidgettest.popup_count}"))
        self.layout().addWidget(QLabel("Label 1"))
        if texts_with_links:
            for text, link in texts_with_links:
                label = QLabel(f'This is a <a href="{link}">{text}</a> to Google')
                label.setOpenExternalLinks(True)
                self.layout().addWidget(label)

        self.counter_label = QLabel("0")
        self.layout().addWidget(self.counter_label)

        self.counter = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_counter)
        self.timer.start(1000)  # Update every 1 second

    def update_counter(self):
        self.counter += 1
        self.counter_label.setText(str(self.counter))

    def focusOutEvent(self, event):
        self.close()

class PopupInstructions(QWidget):
    def __init__(self,text=None):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.Popup)
        self.setLayout(QVBoxLayout())
        if text:
            label = QLabel(text)
            label.setWordWrap(True)
            self.layout().addWidget(label)
        self.setMaximumWidth(800)
    def focusOutEvent(self, event):
        self.close()


class NbibImportChoice:
    CANCEL = "cancel"
    EXCEL = "excel"
    TSV = "tsv"


TSV_IMPORT_HINT = (
    "You can open the TSV, copy the table, and paste it into a Pub-Xel worksheet."
)


class NbibImportDialog(QDialog):
    """Offer Excel or TSV export based on nbib article count."""

    def __init__(self, parent=None, article_count: int = 0):
        super().__init__(parent)
        self._choice = NbibImportChoice.CANCEL
        self.setWindowTitle("Import PubMed nbib")
        self.setModal(True)
        self.setMinimumWidth(280)
        self.setMaximumWidth(320)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        label = QLabel(f"{article_count} PubMed article(s) found in this file.")
        label.setWordWrap(True)
        layout.addWidget(label)

        is_tsv = article_count > NBIB_EXCEL_MAX_ARTICLES
        if is_tsv:
            hint = QLabel(TSV_IMPORT_HINT)
            hint.setWordWrap(True)
            hint_font = hint.font()
            hint_font.setItalic(True)
            hint.setFont(hint_font)
            layout.addWidget(hint)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(8)

        if is_tsv:
            primary_button = QPushButton("Make TSV File")
            primary_button.clicked.connect(self._choose_tsv)
        else:
            primary_button = QPushButton("Make Excel Worksheet")
            primary_button.clicked.connect(self._choose_excel)

        primary_button.setDefault(True)
        primary_button.setAutoDefault(True)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(primary_button)
        button_layout.addStretch()
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        primary_button.setFocus()

    def _choose_excel(self) -> None:
        self._choice = NbibImportChoice.EXCEL
        self.accept()

    def _choose_tsv(self) -> None:
        self._choice = NbibImportChoice.TSV
        self.accept()

    def choice(self) -> str:
        return self._choice


class RunningFunctionDialog(QDialog):
    def __init__(self, parent=None, message=None):
        # The action_in_progress flag and parent.setEnabled() are owned by
        # the caller (run_check_file_exist2 / run_input_pubmed_data2) —
        # see pubxel_core/runtime.py.
        super().__init__()
        self.pseudo_parent = parent

        self.setWindowTitle("Running Function")
        self.setModal(True)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setObjectName("widgettemp1")
        self.setStyleSheet('QDialog#widgettemp1 { border: 1px solid black; background-color: white; }')
        self.setFixedSize(200, 100)
        
        # Create and center the label
        self.label = QLabel(message, self)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setObjectName("widgettemp2")
        self.label.setStyleSheet('QLabel#widgettemp2 { color: black;  }')
        
        # Create the main layout and add the label to it
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.label)
        self.setLayout(main_layout)

        # Center the dialog in the parent window
        if parent:
            parent_rect = parent.geometry()
            self.move(
                parent_rect.center().x() - self.width() // 2,
                parent_rect.center().y() - self.height() // 2
            )
        # Set the cursor to waiting shape: current code wont work
        # self.setCursor(QCursor(Qt.CursorShape.WaitCursor))
        self.show()
        
    def closeEvent(self, event):
        # action_in_progress and parent.setEnabled() are owned by the caller
        # (run_check_file_exist2 / run_input_pubmed_data2).
        event.accept()

    def showEvent(self,  event):
        # Center the dialog in the parent window
        if self.parent():
            parent_rect = self.parent().geometry()
            self.move(
                parent_rect.center().x() - self.width() // 2,
                parent_rect.center().y() - self.height() // 2
            )
        super().showEvent(event)
