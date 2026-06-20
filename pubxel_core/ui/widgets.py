# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>

import copy
import datetime
import os
import platform
import re
import shutil
import tempfile
import threading
import webbrowser

import concurrent.futures
import xlwings as xw
from pynput import keyboard
from PyQt6 import QtCore, uic
from PyQt6.QtCore import QEvent, QObject, QPropertyAnimation, QThread, QTimer, Qt, QMimeData, QUrl, pyqtSignal
from PyQt6.QtGui import QFontMetrics, QGuiApplication, QIcon, QKeySequence, QMovie, QPixmap, QShortcut, QTextCursor
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
    QStyle,
    QSystemTrayIcon,
    QVBoxLayout,
    QWidget,
)

from data.version import __version__
from pubxel_core import runtime as rt
from pubxel_core.excel_ops import check_file_exist, copy_list, files_name_to_path, process_ids
from pubxel_core.clipboard import message_for_action, read_clipboard
from pubxel_core.ids import list_to_string
from pubxel_core.pubmed import input_pubmed_data, normalize_pmid, normalize_pmid_list, resolve_metadata_for_pmids
from pubxel_core.settings import save_settings, save_settings_key
from pubxel_core.ui.dialogs_extra import RunningFunctionDialog
from pubxel_core.ui.helpers import (
    default_worksheet_save_name,
    dialog_onebutton,
    dirname,
    format_excel_operation_error,
    open_directory,
    show_worksheet_saved_dialog,
    try_open_directory,
)
from pubxel_core.worksheet_builder import create_filled_worksheet

class PopupMessageFade(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background-color: rgba(0, 0, 0, 150); color: white; padding: 10px; border-radius: 5px;")
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.ToolTip)

    def show_popup(self, message, duration=3000):
        self.setText(message)
        self.adjustSize()
        parent = self.parent()
        parent_pos = parent.pos()
        self.move(parent_pos.x() + (parent.width() - self.width()) // 2, 
                  parent_pos.y() + (parent.height() - self.height()) // 2)
        self.show()

        # Fade out animation
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(duration)
        self.animation.setStartValue(1)
        self.animation.setEndValue(0)
        self.animation.finished.connect(self.deleteLater)
        self.animation.start()

class window_about(QDialog):
    def __init__(self, parent):
        # The action_in_progress flag and parent.setEnabled() are owned by
        # the caller (open_about_window) — see pubxel_core/runtime.py.
        super().__init__()
        uic.loadUi(rt.about_path, self)
        self.hide()
        self.pseudo_parent = parent

        self.setWindowTitle('About')
# global __version__ -> rt.__version__
        link = "https://pubxel.org/"
        self.findChild(QLabel, 'label_version').setText(f"Version: {__version__}")
        self.label_home = self.findChild(QLabel, 'label_home')
        self.label_home.setText(f'Home: <a href="{link}">{link}</a> ')
        self.label_home.setOpenExternalLinks(True)
        #question mark icons
        pixmap = QPixmap(rt.loading_image_path)
        if pixmap.width() > 500:
            pixmap = pixmap.scaledToWidth(500, Qt.TransformationMode.SmoothTransformation)
        layout = self.findChild(QGridLayout, "gridLayout_image")
        self.label_q1 = QLabel("")
        self.label_q1.setPixmap(pixmap)
        layout.addWidget(self.label_q1)

        button_ok = self.findChild(QPushButton, 'button_ok')
        button_ok.clicked.connect(self.close_about_window)
        shortcut_Esc = QShortcut(QKeySequence('Esc'), self)
        shortcut_Esc.activated.connect(self.close_about_window)
        self.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        self.show()

    def close_about_window(self):
        self.close()


class window_inspect(QWidget):
    _pubmed_metadata_ready = pyqtSignal(object, list)
    _pubmed_connection_status = pyqtSignal(str, str)
    _worksheet_built = pyqtSignal(str)
    _worksheet_build_failed = pyqtSignal(str)
    _pubmed_load_finished = pyqtSignal()

    def __init__(self, parent, data=None, clipboard_ids=None):
        # The action_in_progress flag is acquired by the caller
        # (open_inspect_window) and released by this widget's closeEvent —
        # see pubxel_core/runtime.py and main_window.open_inspect_window.
        # parent.setEnabled(False) is also done by the caller.
        super().__init__()
        uic.loadUi(rt.inspect_path, self)
        self.setWindowIcon(QIcon(rt.icon_path))
        self.hide()
        self.pseudo_parent = parent

        self.setWindowTitle('Inspect')
        self.data = data  # Store the passed data
        self.pubmeddata = None
        self._pubmed_metadata_ready.connect(self._on_pubmed_metadata_ready)
        self._pubmed_connection_status.connect(self._on_pubmed_connection_status)
        self._worksheet_built.connect(self._on_worksheet_built)
        self._worksheet_build_failed.connect(self._on_worksheet_build_failed)
        self._pubmed_load_finished.connect(self._on_pubmed_metadata_load_done)
        self._worksheet_buttons: list[QPushButton] = []
        self._worksheet_build_in_progress = False
        self._pubmed_load_in_progress = False
        self._worksheet_dialog = None
        self._spinner_frames = ["|", "/", "-", "\\"]
        self._spinner_index = 0
        self._spinner_timer = QTimer(self)
        self._spinner_timer.setInterval(120)
        self._spinner_timer.timeout.connect(self._advance_pubmed_spinner)
        assets_dir = os.path.dirname(rt.loading_image_path)
        repo_dir = os.path.dirname(os.path.dirname(os.path.dirname(rt.inspect_path)))
        spinner_candidates = [
            os.path.join(assets_dir, "spinner.gif"),
            os.path.join(repo_dir, "assets", "spinner.gif"),
            os.path.join(os.path.dirname(repo_dir), "assets", "spinner.gif"),
        ]
        self.spinner_gif_path = next((p for p in spinner_candidates if os.path.isfile(p)), "")
        self._spinner_movie = None
        self.pubmed_conn_state = "idle"
        self.pubmed_conn_error = ""
        self.pubmed_conn_enabled = True

        if clipboard_ids is None:
            clip = read_clipboard()
            # Inspect: ids only — never file_paths (see pubxel_core/clipboard.py).
            msg = message_for_action(clip)
            if msg:
                dialog_onebutton(parent, msg, "Clipboard")
                clipboard_ids = []
            else:
                clipboard_ids = clip.ids
        selected_cell_value = clipboard_ids
        output = process_ids(selected_cell_value, rt.mainlibdir, rt.seclibdir)
        valid_ids = output[0]
        pubmed_ids = output[1]
        non_pubmed_valid_ids = output[2]
        invalid_ids = output[3]
        valid_ids_with_m_files = output[4]
        valid_ids_without_m_files = output[5]
        pubmed_ids_with_m_files = output[6]
        pubmed_ids_without_m_files = output[7]
        pubmed_ids_with_s_files = output[8]
        pubmed_ids_without_s_files = output[9]
        all_m_files = output[10]
        all_s_files = output[11]
        nonpubmed_ids_with_m_files = output[12]
        nonpubmed_ids_without_m_files = output[13]
        self.file_map = output[14]

        self.all_ids = valid_ids + invalid_ids
        self.all_files = all_m_files + all_s_files
        
        self.label_title = self.findChild(QLabel, 'label_title')
        self.label_pubmed_connection_status = self.findChild(QLabel, 'label_pubmed_connection_status')
        self.gridLayout_title = self.findChild(QGridLayout, 'gridLayout_title')
        self.scrollArea_idlist = self.findChild(QScrollArea, 'scrollArea_idlist')
        self.scrollLayout_idlist = self.findChild(QWidget, 'scrollLayout_idlist')
        self.label_idlist = self.findChild(QLabel, 'label_idlist')
        if len(valid_ids) ==1 and len(pubmed_ids) == 1 and len(invalid_ids) ==0:           
            inspect_window_title=f"PMID {pubmed_ids[0]}.\n\n"
            self.onepubmedarticle = 1
            self.onepubmedarticle_id = pubmed_ids[0]
            self.scrollArea_idlist.deleteLater()
            if self.gridLayout_title is not None:
                # In single-PMID mode, let the title consume the center column too.
                self.gridLayout_title.addWidget(self.label_title, 0, 0, 1, 2)
                self.gridLayout_title.setColumnStretch(0, 1)
                self.gridLayout_title.setColumnStretch(1, 1)
            self.label_title.installEventFilter(self)
            self._enable_label_context_menu(self.label_title)
            self.label_title.setWordWrap(True)
        elif not valid_ids and not invalid_ids:
            inspect_window_title="No ID(s) selected."
        elif not valid_ids and invalid_ids:
            inspect_window_title="No valid ID(s) selected."
        else: 
            self.onepubmedarticle = 0
            inspect_window_title = f"\n{len(self.all_ids)} ID(s) selected:  \n"
            if len(valid_ids)+len(invalid_ids) > 3:
                inspect_window_title = "\n"+inspect_window_title+"\n"

            self.label_title.setSizePolicy(QSizePolicy.Policy.Preferred,QSizePolicy.Policy.Preferred)
            self.scrollArea_idlist.setSizePolicy(QSizePolicy.Policy.Expanding,QSizePolicy.Policy.Ignored)
            
            self.labels_all = []
            if self.all_ids:
                for file in self.all_ids:
                    idlabel = QLabel(file)
                    idlabel.installEventFilter(self)
                    self._enable_label_context_menu(idlabel)
                    self.labels_all.append(idlabel)

                    h_layout = QHBoxLayout()
                    h_layout.setContentsMargins(0, 0, 0, 0)
                    h_layout.addWidget(idlabel)
                    h_layout.addStretch()
                    container_widget = QWidget()
                    container_widget.setLayout(h_layout)
                    self.scrollLayout_idlist.layout().addWidget(container_widget)
        
        self.label_title.setText(inspect_window_title)
        if self.label_pubmed_connection_status is not None:
            if self.gridLayout_title is not None:
                self.gridLayout_title.setAlignment(
                    self.label_pubmed_connection_status,
                    Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight,
                )
            self.pubmed_conn_enabled = bool(pubmed_ids)
            # Keep a fixed-size reserved slot so title layout width is stable.
            self.label_pubmed_connection_status.setVisible(True)
            self.label_pubmed_connection_status.installEventFilter(self)
            self.label_pubmed_connection_status.setToolTip("")
        if self.pubmed_conn_enabled:
            self.set_pubmed_conn_idle()
        else:
            self.label_pubmed_connection_status.setText("")
            self.label_pubmed_connection_status.setPixmap(QPixmap())
        # self.label_title.installEventFilter(self)

        gridLayout_summary = self.findChild(QGridLayout, 'gridLayout_summary')
        if pubmed_ids:
            (
                label_pubmed,
                button_copypubmed,
                button_searchpubmed,
                button_worksheetpubmed,
            ) = self.create_elements(
                gridLayout_summary,
                0,
                'label_pubmed',
                'button_copypubmed',
                'button_searchpubmed',
                'button_worksheetpubmed',
            )
            label_pubmed.setText(f"PubMed ID(s): {len(pubmed_ids)}")
            button_copypubmed.clicked.connect(lambda: self.show_copy_id_and_show_popup(pubmed_ids))
            button_searchpubmed.clicked.connect(lambda: self.search_pubmed(pubmed_ids))
            button_searchpubmed.setText('View &PubMed')
            button_worksheetpubmed.clicked.connect(lambda: self.make_worksheet(pubmed_ids))
            button_worksheetpubmed.setText('Make &Worksheet')
            self._worksheet_buttons.append(button_worksheetpubmed)
            shortcut_p = QShortcut(QKeySequence('p'), self)
            shortcut_p.activated.connect(lambda: self.search_pubmed(pubmed_ids))
        if pubmed_ids_without_m_files:
            (
                label_pubmedna,
                button_copypubmedna,
                button_searchpubmedna,
                button_worksheetpubmedna,
            ) = self.create_elements(
                gridLayout_summary,
                1,
                'label_pubmedna',
                'button_copypubmedna',
                'button_searchpubmedna',
                'button_worksheetpubmedna',
            )
            label_pubmedna.setText(f"    PubMed ID(s) without main files: {len(pubmed_ids_without_m_files)}")
            button_copypubmedna.clicked.connect(lambda: self.show_copy_id_and_show_popup(pubmed_ids_without_m_files))
            button_searchpubmedna.clicked.connect(lambda: self.search_pubmed(pubmed_ids_without_m_files))
            button_searchpubmedna.setText('View &PubMed')
            button_worksheetpubmedna.clicked.connect(
                lambda: self.make_worksheet(pubmed_ids_without_m_files)
            )
            button_worksheetpubmedna.setText('Make &Worksheet')
            self._worksheet_buttons.append(button_worksheetpubmedna)
        if non_pubmed_valid_ids:
            label_nonpub, button_copynonpub, _, _ = self.create_elements(
                gridLayout_summary, 2, 'label_nonpub', 'button_copynonpub'
            )
            label_nonpub.setText(f"Non-PubMed ID(s): {len(non_pubmed_valid_ids)}")
            button_copynonpub.clicked.connect(lambda: self.show_copy_id_and_show_popup(non_pubmed_valid_ids))
        if nonpubmed_ids_without_m_files:
            label_nonpubna, button_copynonpubna, _, _ = self.create_elements(
                gridLayout_summary, 3, 'label_nonpubna', 'button_copynonpubna'
            )
            label_nonpubna.setText(f"    Non-PubMed ID(s) without main files: {len(nonpubmed_ids_without_m_files)}")
            button_copynonpubna.clicked.connect(lambda: self.show_copy_id_and_show_popup(nonpubmed_ids_without_m_files))
        if invalid_ids:
            label_na, button_copyna, _, _ = self.create_elements(
                gridLayout_summary, 4, 'label_na', 'button_copyna'
            )
            label_na.setText(f"Invalid ID(s): {len(invalid_ids)}")
            button_copyna.clicked.connect(lambda: self.show_copy_id_and_show_popup(invalid_ids))

        self._update_worksheet_buttons_state()
        
        self.scrollLayout_main = self.findChild(QWidget, 'scrollLayout_main')
        self.scrollLayout_suppl = self.findChild(QWidget, 'scrollLayout_suppl')
        self.frameupper_main = self.findChild(QWidget, 'frameupper_main')
        self.frameupper_suppl = self.findChild(QWidget, 'frameupper_suppl')
        
        checked_main = []
        checked_suppl = []
        
        #main files
        if not all_m_files:
            label = QLabel("No files available.")
            self.frameupper_main.layout().addWidget(label)
        else:
            n=len(all_m_files)
            self.checkbox_all_main = QCheckBox(f"All ({n})")
            self.checkbox_all_main.stateChanged.connect(self.on_checkbox_all_main)
            self.frameupper_main.layout().addWidget(self.checkbox_all_main)
        self.checkboxes_main = []
        if not all_m_files:
            pass
        else:
            for file in all_m_files:
                checkbox = QCheckBox(file)
                checkbox.stateChanged.connect(self.on_checkbox_main)
                checkbox.installEventFilter(self)
                self._enable_checkbox_text_copy(checkbox)
                self.checkboxes_main.append(checkbox)
                h_layout = QHBoxLayout()
                h_layout.setContentsMargins(0, 0, 0, 0)
                h_layout.addWidget(checkbox)
                h_layout.addStretch()
                container_widget = QWidget()
                container_widget.setLayout(h_layout)
                self.scrollLayout_main.layout().addWidget(container_widget)

        self.scrollLayout_main.layout().addStretch()

        #suppl files
        if not all_s_files:
            label = QLabel("No files available.")
            self.frameupper_suppl.layout().addWidget(label)
        else:
            n=len(all_s_files)
            self.checkbox_all_suppl = QCheckBox(f"All ({n})")
            self.checkbox_all_suppl.stateChanged.connect(self.on_checkbox_all_suppl)
            self.frameupper_suppl.layout().addWidget(self.checkbox_all_suppl)
        self.checkboxes_suppl = []
        if not all_s_files:
            pass
        else:
            for file in all_s_files:
                checkbox = QCheckBox(file)
                checkbox.stateChanged.connect(self.on_checkbox_suppl)
                checkbox.installEventFilter(self)
                self._enable_checkbox_text_copy(checkbox)
                self.checkboxes_suppl.append(checkbox)
                h_layout = QHBoxLayout()
                h_layout.setContentsMargins(0, 0, 0, 0)
                h_layout.addWidget(checkbox)
                h_layout.addStretch()
                container_widget = QWidget()
                container_widget.setLayout(h_layout)
                self.scrollLayout_suppl.layout().addWidget(container_widget)

        self.scrollLayout_suppl.layout().addStretch()

        thread = threading.Thread(target=self.load_pubmed_data, args=(pubmed_ids,))
        if pubmed_ids:
            self._pubmed_load_in_progress = True
            self._update_worksheet_buttons_state()
        thread.start()

        button_openall = self.findChild(QPushButton, 'button_openall')
        button_openall.setText('&All')
        button_opensel = self.findChild(QPushButton, 'button_opensel')
        button_opensel.setText('&Selected')
        button_exportall = self.findChild(QPushButton, 'button_exportall')
        button_exportsel = self.findChild(QPushButton, 'button_exportsel')        
        
        button_openall.clicked.connect(lambda: self.callback(self.open_all_files()))
        button_opensel.clicked.connect(lambda: self.callback(self.open_sel_files()))
        button_exportall.clicked.connect(self.export_all_files)
        button_exportsel.clicked.connect(self.export_sel_files)

        shortcut_a = QShortcut(QKeySequence('a'), self)
        shortcut_a.activated.connect(lambda: self.callback(self.open_all_files()))
        shortcut_s = QShortcut(QKeySequence('s'), self)
        shortcut_s.activated.connect(lambda: self.callback(self.open_sel_files()))

        button_exit = self.findChild(QPushButton, 'button_exit')
        button_exit.clicked.connect(self.close_inspect_window)
        button_exit.setText('Exit')
        shortcut_Esc = QShortcut(QKeySequence('Esc'), self)
        shortcut_Esc.activated.connect(self.close_inspect_window)

        self.popup_message = None

        self.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        self.show()
        self.activateWindow()
        self.raise_()
        self._make_window_labels_selectable()

        # inspect_window.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        # # After a short delay, remove the always on top attribute
        # QTimer.singleShot(1000, lambda: (inspect_window.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, False), inspect_window.show()))
    
    def show_copy_id_and_show_popup(self,pubmed_ids):
        copy_list(pubmed_ids)
        self.hide_popup_message()
        self.popup_message = PopupMessageFade(self)  # Create a new instance each time
        self.popup_message.show_popup("ID(s) copied to clipboard.")
        

    @staticmethod
    def _article_field(pubmeddata, pmid: str, key: str, default: str = "") -> str:
        if not pubmeddata:
            return default
        article = pubmeddata.get(pmid, {}).get("article") or {}
        value = article.get(key, default)
        return value if value is not None else default

    def _on_pubmed_metadata_ready(self, data, pubmed_ids):
        """Main-thread slot: apply metadata loaded from background thread."""
        self.pubmeddata = data
        self._apply_pubmed_metadata_to_ui(pubmed_ids)
        self._update_worksheet_buttons_state()

    def _on_pubmed_connection_status(self, state: str, detail: str = ""):
        if state == "loading":
            self.set_pubmed_conn_loading()
        elif state == "success":
            self.set_pubmed_conn_success()
        elif state == "error":
            self.set_pubmed_conn_error(detail)
        else:
            self.set_pubmed_conn_idle()
        self._update_worksheet_buttons_state()

    def _on_pubmed_metadata_load_done(self) -> None:
        self._pubmed_load_in_progress = False
        self._update_worksheet_buttons_state()

    def _worksheet_ready(self) -> bool:
        return bool(
            self.pubmeddata
            and not self._pubmed_load_in_progress
            and self.pubmed_conn_state in ("idle", "success")
            and not self._worksheet_build_in_progress
        )

    def _update_worksheet_buttons_state(self) -> None:
        ready = self._worksheet_ready()
        for button in self._worksheet_buttons:
            button.setEnabled(ready)

    def _set_pubmed_conn_state(self, state: str, error: str = ""):
        if not self.pubmed_conn_enabled:
            return
        self.pubmed_conn_state = state
        self.pubmed_conn_error = error
        self.render_pubmed_conn_icon()

    def set_pubmed_conn_idle(self):
        self._set_pubmed_conn_state("idle")

    def set_pubmed_conn_loading(self):
        self._set_pubmed_conn_state("loading")

    def set_pubmed_conn_success(self):
        self._set_pubmed_conn_state("success")

    def set_pubmed_conn_error(self, message: str | None = None):
        self._set_pubmed_conn_state("error", message or "")

    def _advance_pubmed_spinner(self):
        if (
            not self.pubmed_conn_enabled
            or self.pubmed_conn_state != "loading"
            or self.label_pubmed_connection_status is None
        ):
            self._spinner_timer.stop()
            return
        frame = self._spinner_frames[self._spinner_index]
        self._spinner_index = (self._spinner_index + 1) % len(self._spinner_frames)
        self.label_pubmed_connection_status.setPixmap(QPixmap())
        self.label_pubmed_connection_status.setStyleSheet("color: #1f6feb; font-weight: bold;")
        self.label_pubmed_connection_status.setText(frame)

    def render_pubmed_conn_icon(self):
        if not self.pubmed_conn_enabled or self.label_pubmed_connection_status is None:
            return
        self.label_pubmed_connection_status.setStyleSheet("")
        if self.pubmed_conn_state == "loading":
            if self.spinner_gif_path:
                self._spinner_timer.stop()
                if self._spinner_movie is None:
                    self._spinner_movie = QMovie(self.spinner_gif_path)
                    self._spinner_movie.setScaledSize(QtCore.QSize(16, 16))
                self.label_pubmed_connection_status.setText("")
                self.label_pubmed_connection_status.setMovie(self._spinner_movie)
                if self._spinner_movie.state() != QMovie.MovieState.Running:
                    self._spinner_movie.start()
            else:
                self._spinner_index = 0
                if not self._spinner_timer.isActive():
                    self._spinner_timer.start()
                self._advance_pubmed_spinner()
            return
        if self._spinner_movie is not None:
            self._spinner_movie.stop()
            self.label_pubmed_connection_status.setMovie(None)
        self._spinner_timer.stop()
        self.label_pubmed_connection_status.setText("")
        if self.pubmed_conn_state in ("idle", "success"):
            icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton)
        else:
            icon = self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxCritical)
        self.label_pubmed_connection_status.setPixmap(icon.pixmap(16, 16))

    def _apply_pubmed_metadata_to_ui(self, pubmed_ids):
        """Update title and checkbox labels from self.pubmeddata (main thread)."""
        if not self.pubmeddata:
            return

        if getattr(self, "onepubmedarticle", 0) and pubmed_ids:
            try:
                pmid = normalize_pmid(pubmed_ids[0])
                cite = self._article_field(self.pubmeddata, pmid, "cite")
                if cite:
                    self.label_title.setText(f"PMID {pmid} :\n{cite}")
                else:
                    self.label_title.setText(f"PMID {pmid}.")
            except Exception as e:
                print(f"Failed to set label_title: {e}")

        if self.checkboxes_main:
            for checkbox in self.checkboxes_main:
                article_id = checkbox.text()
                if article_id and article_id[0].isdigit():
                    match = re.match(r"^\d+", article_id)
                    if match:
                        pubmed_id = normalize_pmid(match.group(0))
                        if pubmed_id in self.pubmeddata:
                            authoryear = self._article_field(self.pubmeddata, pubmed_id, "authoryear")
                            if authoryear and " : " not in article_id:
                                checkbox.setText(article_id + " : " + authoryear)

    def _make_window_labels_selectable(self):
        """Allow mouse selection/copy for inspect window label text."""
        for label in self.findChildren(QLabel):
            flags = label.textInteractionFlags()
            label.setTextInteractionFlags(flags | Qt.TextInteractionFlag.TextSelectableByMouse)

    def _enable_checkbox_text_copy(self, checkbox: QCheckBox):
        checkbox.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        checkbox.customContextMenuRequested.connect(
            lambda pos, cb=checkbox: self._show_checkbox_copy_menu(cb, pos)
        )

    @staticmethod
    def _basename_from_display_text(text: str) -> str:
        return text.split(" : ")[0].strip() if text else ""

    def _article_popup_text_for_display_text(self, display_text: str, *, checkbox: bool) -> str:
        if not display_text or not display_text[0].isdigit() or not self.pubmeddata:
            return ""
        match = re.match(r"^\d+", display_text)
        if not match:
            return ""
        pubmed_id = normalize_pmid(match.group(0))
        if pubmed_id not in self.pubmeddata:
            return ""
        if checkbox:
            return self._article_field(self.pubmeddata, pubmed_id, "cite_maincheckbox") or display_text
        return self._article_field(self.pubmeddata, pubmed_id, "cite") or display_text

    def _copy_file_to_clipboard(self, full_path: str):
        mime = QMimeData()
        mime.setUrls([QUrl.fromLocalFile(full_path)])
        clipboard = QGuiApplication.clipboard()
        if clipboard is not None:
            clipboard.setMimeData(mime)

    def _show_checkbox_copy_menu(self, checkbox: QCheckBox, pos):
        menu = QMenu(self)
        action_copy = menu.addAction("Copy text")
        article_text = self._article_popup_text_for_display_text(checkbox.text(), checkbox=True)
        action_copy_article = menu.addAction("Copy article data") if article_text else None
        file_key = self._basename_from_display_text(checkbox.text())
        file_path = getattr(self, "file_map", {}).get(file_key)
        action_copy_file = menu.addAction("Copy file") if file_path else None
        selected = menu.exec(checkbox.mapToGlobal(pos))
        if selected == action_copy:
            copy_list(checkbox.text())
        elif action_copy_article is not None and selected == action_copy_article:
            copy_list(article_text)
        elif action_copy_file is not None and selected == action_copy_file and file_path:
            self._copy_file_to_clipboard(file_path)

    def _enable_label_context_menu(self, label: QLabel):
        label.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        label.customContextMenuRequested.connect(
            lambda pos, lb=label: self._show_label_copy_menu(lb, pos)
        )

    def _show_label_copy_menu(self, label: QLabel, pos):
        menu = QMenu(self)
        action_copy = menu.addAction("Copy text")
        article_text = self._article_popup_text_for_display_text(label.text(), checkbox=False)
        action_copy_article = menu.addAction("Copy article data") if article_text else None
        selected = menu.exec(label.mapToGlobal(pos))
        if selected == action_copy:
            copy_list(label.text())
        elif action_copy_article is not None and selected == action_copy_article:
            copy_list(article_text)

    def load_pubmed_data(self, pubmed_ids):
        if not pubmed_ids:
            return

        fetch_started = False

        def on_partial(data):
            self._pubmed_metadata_ready.emit(data, pubmed_ids)

        def on_fetch_start():
            nonlocal fetch_started
            fetch_started = True
            self._pubmed_connection_status.emit("loading", "")

        try:
            data = resolve_metadata_for_pmids(
                pubmed_ids,
                on_partial=on_partial,
                on_fetch_start=on_fetch_start,
            )
            if isinstance(data, dict):
                self.pubmeddata = data
            print("Data loaded: " + str(len(data)) + " items")
            self._pubmed_connection_status.emit("success" if fetch_started else "idle", "")
        except Exception as e:
            print(f"Failed to load data: {e}")
            self._pubmed_connection_status.emit("error", str(e))
        finally:
            self._pubmed_load_finished.emit()

    def eventFilter(self, source, event):
        if isinstance(source, QCheckBox) or isinstance(source, QLabel):
            if event.type() == QEvent.Type.Enter:
                QTimer.singleShot(0, lambda: self.show_popup(source))
            elif event.type() == QEvent.Type.Leave or event.type() == QEvent.Type.HoverLeave:
                self.hide_popup_tooltip()
        return super().eventFilter(source, event)

    def show_popup(self, source):
        self.hide_popup_tooltip()
        source.setAttribute(Qt.WidgetAttribute.WA_Hover, True)
        if source is getattr(self, "label_pubmed_connection_status", None):
            if not self.pubmed_conn_enabled:
                return
            state_text = {
                "idle": "PubMed data brought locally.",
                "loading": "Connecting to PubMed...",
                "success": "PubMed connection finished.",
                "error": "PubMed connection failed.",
            }
            popuptext = state_text.get(self.pubmed_conn_state, "PubMed connection status unknown.")
            if self.pubmed_conn_state == "error" and self.pubmed_conn_error:
                popuptext += "\n" + self.pubmed_conn_error
            self.popup_toolip = QLabel(popuptext, self)
            self.popup_toolip.setWindowFlags(Qt.WindowType.ToolTip)
            self.popup_toolip.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            self.popup_toolip.setWordWrap(True)
            self.popup_toolip.setStyleSheet(
                "QLabel {"
                "  background-color: #e0e0e0;"
                "  color: #222;"
                "  padding: 6px;"
                "  border: 1px solid #c8c8c8;"
                "  border-radius: 4px;"
                "}"
            )
            lines = popuptext.split("\n")
            largest_width = 0
            for line in lines:
                line_width = self.fontMetrics().boundingRect(line).width()
                if line_width > largest_width:
                    largest_width = line_width
            max_width = 420
            self.popup_toolip.setFixedWidth(min(max_width, largest_width + 30))
            anchor_bottom_left = source.mapToGlobal(source.rect().bottomLeft())
            anchor_bottom_center = source.mapToGlobal(source.rect().center()) + QtCore.QPoint(
                0, source.rect().height() // 2
            )
            popup_w = self.popup_toolip.width()
            popup_h = self.popup_toolip.sizeHint().height()
            available = QApplication.primaryScreen().availableGeometry()
            x = anchor_bottom_center.x() - (popup_w // 2)
            y = anchor_bottom_center.y() + 4
            if x + popup_w > available.right() - 8:
                x = anchor_bottom_left.x() - popup_w - 20
            if x < available.left() + 8:
                x = available.left() + 8
            if y + popup_h > available.bottom() - 8:
                y = max(available.top() + 8, available.bottom() - popup_h - 8)
            self.popup_toolip.move(x, y)
            self.popup_toolip.show()
            return
        if isinstance(source, QCheckBox):
            article_id = source.text()
        elif isinstance(source, QLabel) and self.onepubmedarticle:
            article_id = self.onepubmedarticle_id
        elif isinstance(source, QLabel) and not self.onepubmedarticle:
            article_id = source.text()
        else: article_id=""
        popuptext = ""
        if article_id and article_id[0].isdigit() and self.pubmeddata:
            match = re.match(r"^\d+", article_id) # remove characters after numerics
            if match:
                pubmed_id = normalize_pmid(match.group(0))
                if pubmed_id in self.pubmeddata:
                    if isinstance(source, QLabel):
                        cite = self._article_field(self.pubmeddata, pubmed_id, "cite")
                        popuptext = f"PMID {article_id} : {cite}"
                    if isinstance(source, QCheckBox):
                        cite_maincheckbox = self._article_field(
                            self.pubmeddata, pubmed_id, "cite_maincheckbox"
                        )
                        popuptext = cite_maincheckbox or article_id
                    if len(popuptext) > 1000:
                        popuptext = popuptext[:1000] + "..."
        elif article_id:
            popuptext = article_id
            if len(popuptext) > 1000:
                popuptext = popuptext[:1000] + "..."
        if popuptext:
            self.popup_toolip = QLabel(popuptext, self)
            self.popup_toolip.setWindowFlags(Qt.WindowType.ToolTip)
            self.popup_toolip.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            self.popup_toolip.setWordWrap(True)
            self.popup_toolip.setStyleSheet(
                "QLabel {"
                "  background-color: #e0e0e0;"
                "  color: #222;"
                "  padding: 6px;"
                "  border: 1px solid #c8c8c8;"
                "  border-radius: 4px;"
                "}"
            )
            # Calculate the width of each line in popuptext and set width
            lines = popuptext.split('\n')
            if isinstance (source, QCheckBox):
                max_width = 500
            if isinstance (source, QLabel) and not self.onepubmedarticle:
                max_width = 500
            if isinstance (source, QLabel) and self.onepubmedarticle:
                max_width = 800
            largest_width = 0
            for line in lines:
                line_width = self.fontMetrics().boundingRect(line).width()
                if line_width > largest_width:
                    largest_width = line_width
            if largest_width + 30 > max_width:
                self.popup_toolip.setFixedWidth(max_width)
            else:
                self.popup_toolip.setFixedWidth(largest_width + 30)
            # position and show popup_toolip
            self.popup_toolip.move(source.mapToGlobal(source.rect().bottomLeft()) + QtCore.QPoint(20, 0))
            self.popup_toolip.show()

    def hide_popup_tooltip(self):
        if hasattr(self, 'popup_toolip') and self.popup_toolip:
            try:
                self.popup_toolip.hide()
                self.popup_toolip.deleteLater()
                del self.popup_toolip
            except RuntimeError as e:
                print(f"Error hiding popup_toolip: {e}")

    def hide_popup_message(self):
        if hasattr(self, 'popup_message') and self.popup_message:
            try:
                self.popup_message.hide()
                self.popup_message.deleteLater()
                del self.popup_message
            except RuntimeError as e:
                print(f"Error hiding popup_message: {e}")

    def close_inspect_window(self):
        self.close()

    def closeEvent(self, event):
        # window_inspect owns action_in_progress for its lifetime — release
        # it here so the main window can start a new action.
        self.hide_popup_message()
        self.hide_popup_tooltip()
        if hasattr(self, "_spinner_timer") and self._spinner_timer is not None:
            self._spinner_timer.stop()
        if hasattr(self, "_spinner_movie") and self._spinner_movie is not None:
            self._spinner_movie.stop()
        self.pseudo_parent.setEnabled(True)
        rt.end_action()
        event.accept()

    def search_pubmed(self,lst):
        webbrowser.open("https://pubmed.ncbi.nlm.nih.gov/?term="+"%5Buid%5D+OR+".join(lst)+"%5Buid%5D&sort=date")
        self.close_inspect_window()

    def make_worksheet(self, pubmed_ids):
        if not self._worksheet_ready():
            return

        pmids = normalize_pmid_list(pubmed_ids)
        if not pmids:
            dialog_onebutton(self, "No PubMed ID(s) to include in the worksheet.", "Make Worksheet")
            return

        self._worksheet_build_in_progress = True
        self._update_worksheet_buttons_state()
        self._worksheet_dialog = RunningFunctionDialog(
            parent=self,
            message="Creating worksheet...\nPlease wait.",
        )
        thread = threading.Thread(
            target=self._build_worksheet_worker,
            args=(pmids,),
            daemon=True,
        )
        thread.start()

    def _build_worksheet_worker(self, pmids: list[str]) -> None:
        temp_path = None
        try:
            fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)
            create_filled_worksheet(
                temp_path,
                pmids,
                self.pubmeddata or {},
                settings=rt.settings,
            )
            self._worksheet_built.emit(temp_path)
        except Exception as e:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
            self._worksheet_build_failed.emit(str(e))

    def _close_worksheet_dialog(self) -> None:
        if self._worksheet_dialog is not None:
            try:
                self._worksheet_dialog.close()
            except Exception:
                pass
            self._worksheet_dialog = None

    def _on_worksheet_built(self, temp_path: str) -> None:
        self._close_worksheet_dialog()
        self._worksheet_build_in_progress = False
        self._update_worksheet_buttons_state()

        file_dialog = QFileDialog(self, "Save Worksheet")
        file_dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        file_dialog.setDefaultSuffix("xlsx")
        file_dialog.setNameFilters(["Excel Files (*.xlsx)", "All Files (*)"])
        save_dir = os.path.expanduser("~/Documents")
        try:
            if not os.path.isdir(save_dir):
                save_dir = os.path.expanduser("~")
            file_dialog.setDirectory(save_dir)
        except Exception:
            save_dir = os.path.expanduser("~")
        file_dialog.selectFile(default_worksheet_save_name(save_dir))

        dest_path = None
        if file_dialog.exec() == QFileDialog.DialogCode.Accepted:
            dest_path = file_dialog.selectedFiles()[0]

        try:
            if dest_path:
                shutil.copy2(temp_path, dest_path)
                show_worksheet_saved_dialog(self, dest_path)
        except Exception as e:
            dialog_onebutton(self, f"Failed to save worksheet:\n{e}", "Error")
        finally:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except OSError:
                pass

    def _on_worksheet_build_failed(self, message: str) -> None:
        self._close_worksheet_dialog()
        self._worksheet_build_in_progress = False
        self._update_worksheet_buttons_state()
        dialog_onebutton(
            self,
            f"Failed to create worksheet:\n{format_excel_operation_error(message)}",
            "Error",
        )

    def callback(self,result): #open files and quit new_window
        if result == []:
            return
        # Handle the result in your higher-level code
        filepathList = files_name_to_path(
            result,
            rt.mainlibdir,
            rt.seclibdir,
            file_map=getattr(self, "file_map", None),
        )

        failed_files = []  # List to keep track of files that could not be opened
        for filepath in filepathList:
            try:
                open_directory(filepath)
            except Exception as e:
                failed_files.append(filepath)  # Add the failed file to the list
                continue

        # If there were any failed files, show a messagebox after all file opening attempts are completed
        if failed_files:
            dialog_onebutton(self,"The following files could not be opened:\n" + "\n".join(failed_files),"Error")

        print("Result:", result)
        self.close_inspect_window()

    def open_files(self,files):
        try:
            if not files:
                dialog_onebutton(self,"No files to open.","Error")
                return []

            if len(files) > 50:
                dialog_onebutton(self,"Cannot open more than 50 files.","Error")
                return []

            # Check if there are 6 or more files to open
            if len(files) >= 6:
                msg_box = QMessageBox()
                msg_box.setWindowTitle("Warning")
                msg_box.setText("You are about to open " + str(len(files)) + " files. Proceed?")
                msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                msg_box.setWindowModality(Qt.WindowModality.ApplicationModal)
                proceed = msg_box.exec() == QMessageBox.StandardButton.Yes
                if not proceed:
                    return []
            return(files)
        except Exception as e:
            dialog_onebutton(self,str(e),"Error")
            return []

    def open_sel_files(self):
        checked_main = [checkbox.text() for checkbox in self.checkboxes_main if checkbox.isChecked()]
        checked_suppl = [checkbox.text() for checkbox in self.checkboxes_suppl if checkbox.isChecked()]
        checked = checked_main + checked_suppl
        if checked:
            checked = [item.split(" : ")[0] for item in checked]
        files = self.open_files(checked)
        return(files)

    def open_all_files(self):
        files = self.open_files(self.all_files)
        return(files)

    def export_sel_files(self):
        checked_main = [checkbox.text() for checkbox in self.checkboxes_main if checkbox.isChecked()]
        checked_suppl = [checkbox.text() for checkbox in self.checkboxes_suppl if checkbox.isChecked()]
        checked = checked_main + checked_suppl
        if checked:
            checked = [item.split(" : ")[0] for item in checked]
        result = self.export_files(checked)
        if result:
            self.close_inspect_window()     

    def export_all_files(self):
        result=self.export_files(self.all_files)
        if result:
            self.close_inspect_window()     

    def export_files(self,files):
        # Check if "requestfiles" list is empty 
        if not files:
            dialog_onebutton(self,"No files to export.","Error")
            return False
        
        def create_dated_folder(base_dir):
            today = datetime.datetime.today().strftime('%Y-%m-%d')
            base_name = f"Pub-Xel Export {today}"
            folder_name = base_name
            folder_path = os.path.join(base_dir, folder_name)
            counter = 2

            while os.path.exists(folder_path):
                folder_name = f"{base_name} ({counter})"
                folder_path = os.path.join(base_dir, folder_name)
                counter += 1

            os.makedirs(folder_path)
            return folder_name, folder_path
        
        newfolder, newdir = create_dated_folder(rt.outdir)
        print(f"New folder created at: {newdir}")

        paths = files_name_to_path(
            files,
            rt.mainlibdir,
            rt.seclibdir,
            file_map=getattr(self, "file_map", None),
        )
        
        # Copy wanted files
        copied_files = 0
        for path in paths:
            if os.path.isfile(path):
                shutil.copy(path, newdir)
                copied_files += 1

        def open_newdir():
            try_open_directory(os.path.realpath(newdir))

        total = len(files)
        exported = copied_files
        if exported == total:
            completion_text = f"{exported} exported successfully"
        else:
            completion_text = f"{exported} out of {total} files exported"

        def show_message_box():
            msg_box = QMessageBox()
            msg_box.setWindowTitle("Completion")
            msg_box.setText(completion_text)
            open_folder_button = msg_box.addButton("Open Folder", QMessageBox.ButtonRole.AcceptRole)
            open_folder_button.clicked.connect(lambda: open_newdir())
            msg_box.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
            msg_box.exec()
        show_message_box()
        return True

    def create_elements(
        self,
        layout: QGridLayout,
        row: int,
        label: str,
        copy_button: str,
        view_pubmed_button: str | None = None,
        worksheet_button: str | None = None,
    ):
        """
        Summary row layout (columns):
          0 label | 1 spacer | 2 View PubMed | 3 Make Worksheet | 4 Copy ID(s)

        PubMed rows pass view_pubmed_button and worksheet_button.
        Non-PubMed rows pass copy only (column 4).
        """
        label_widget = QLabel(label)
        layout.addWidget(label_widget, row, 0)
        spacer_item = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        layout.addItem(spacer_item, row, 1)

        copy_widget = QPushButton(copy_button)
        copy_widget.setText("Copy ID(s)")
        layout.addWidget(copy_widget, row, 4)

        view_widget = None
        worksheet_widget = None
        if view_pubmed_button:
            view_widget = QPushButton(view_pubmed_button)
            view_widget.setText("View PubMed")
            layout.addWidget(view_widget, row, 2)
        if worksheet_button:
            worksheet_widget = QPushButton(worksheet_button)
            worksheet_widget.setText("Make Worksheet")
            layout.addWidget(worksheet_widget, row, 3)

        return label_widget, copy_widget, view_widget, worksheet_widget

    def sayHello(self):
        print("Hello")
    
    def on_checkbox_all_main(self):
        if self.checkbox_all_main.testAttribute(Qt.WidgetAttribute.WA_UnderMouse):
            state = self.checkbox_all_main.checkState()
            print(state)
            for checkbox in self.checkboxes_main:
                checkbox.setChecked(state == Qt.CheckState.Checked)

    def on_checkbox_main(self):
        clicked = False
        for checkbox in self.checkboxes_main:
            if checkbox.testAttribute(Qt.WidgetAttribute.WA_UnderMouse):
                clicked = True
                break
        if clicked:
            for checkbox in self.checkboxes_main:
                print(f"{checkbox.text()} is checked: {checkbox.isChecked()}")
            all_checked = all(checkbox.isChecked() for checkbox in self.checkboxes_main)
            print(all_checked)
            self.checkbox_all_main.setChecked(all_checked)

    def on_checkbox_all_suppl(self):
        if self.checkbox_all_suppl.testAttribute(Qt.WidgetAttribute.WA_UnderMouse):
            state = self.checkbox_all_suppl.checkState()
            print(state)
            for checkbox in self.checkboxes_suppl:
                checkbox.setChecked(state == Qt.CheckState.Checked)
                
    def on_checkbox_suppl(self):
        clicked = False
        for checkbox in self.checkboxes_suppl:
            if checkbox.testAttribute(Qt.WidgetAttribute.WA_UnderMouse):
                clicked = True
                break
        if clicked:
            for checkbox in self.checkboxes_suppl:
                print(f"{checkbox.text()} is checked: {checkbox.isChecked()}")
            all_checked = all(checkbox.isChecked() for checkbox in self.checkboxes_suppl)
            print(all_checked)
            self.checkbox_all_suppl.setChecked(all_checked)

