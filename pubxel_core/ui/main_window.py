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
from PyQt6.QtCore import QEvent, QObject, QPropertyAnimation, QThread, QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QAction, QFontMetrics, QIcon, QKeySequence, QPixmap, QShortcut, QTextCursor
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
from pubxel_core.clipboard import message_for_action, read_clipboard
from pubxel_core.excel_ops import check_file_exist, copy_list, files_name_to_path, process_ids
from pubxel_core.nbib import load_nbib_file
from pubxel_core.pubmed import import_nbib_to_metadata, input_pubmed_data
from pubxel_core.pubmed import obtain_pubmed_data
from pubxel_core.recent_worksheets import (
    format_recent_menu_label,
    get_recent_worksheets,
    register_recent_worksheet,
    remove_recent_worksheet,
    set_recent_menu_rebuild_callback,
)
from pubxel_core.settings import save_settings, save_settings_key
from pubxel_core.worksheet_builder import create_filled_worksheet, create_worksheet as create_pubsheet_worksheet
from pubxel_core.worksheet_export import write_worksheet_tsv
from pubxel_core.ui.dialogs_extra import (
    NbibImportChoice,
    NbibImportDialog,
    PopupInstructions,
    PopupWidgettest,
    RunningFunctionDialog,
)
from pubxel_core.ui.helpers import (
    default_tsv_save_name,
    default_worksheet_save_directory,
    default_worksheet_save_name,
    dialog_onebutton,
    graceful_shutdown,
    open_directory,
    show_file_saved_dialog,
    show_worksheet_saved_dialog,
    try_open_directory,
)
from pubxel_core.ui.preferences import window_preferences
from pubxel_core.ui.tray import SystemTrayIcon
from pubxel_core.ui.widgets import PopupMessageFade, window_about, window_inspect, window_worksheetColumns
from pubxel_core.ui.workers import excelWorker

_OPEN_FILE_EXTENSIONS = (".xlsx", ".nbib")


class main_window(QMainWindow):
    _nbib_export_built = pyqtSignal(str, str)
    _nbib_export_failed = pyqtSignal(str)

    def __init__(self):
# global rt.settings -> rt.settings
        super().__init__()
        uic.loadUi(rt.main_path, self)
        self.setWindowTitle('Pub-Xel')
        self.setWindowIcon(QIcon(rt.icon_path))
        self.exitstatus = False
        self.is_closing = False

        if rt.os_name == "Darwin":
            appdock = QApplication.instance() # for dock click event
            appdock.applicationStateChanged.connect(self.on_application_state_changed) # for dock click event

        if rt.developerMode:
            self.layout_developer = self.findChild(QHBoxLayout, "layout_developer")
            self.button1 = QPushButton('Button 1')
            self.button2 = QPushButton('Button 2')
            self.button3 = QPushButton('Button 3')
            self.layout_developer.addWidget(self.button1)
            self.layout_developer.addWidget(self.button2)
            self.layout_developer.addWidget(self.button3)
            self.button1.clicked.connect(self.action_in_progress_switch)
            self.button2.clicked.connect(self.crash)
            self.button3.clicked.connect(self.button3_clicked)

        self.button_checkfiles.clicked.connect(self.run_check_file_exist2)
        self.button_import.clicked.connect(self.run_input_pubmed_data2)
        self.button_open.clicked.connect(lambda: self.main_openfile())
        self.button_inspect.clicked.connect(lambda: self.open_inspect_window())
        
        self.findChild(QLabel, 'label_inspect').setText("  "+rt.settings.get('hotkey_inspect_value', 0).replace('<', '').replace('>', ''))
        self.findChild(QLabel, 'label_open').setText("  "+rt.settings.get('hotkey_open_value', 0).replace('<', '').replace('>', ''))

        #question mark icons
        pixmap = QPixmap(rt.questionmark_icon_path)
        layout_q1 = self.findChild(QGridLayout, "layout_q1")
        self.label_q1 = QLabel("")
        self.label_q1.setPixmap(pixmap)
        layout_q1.addWidget(self.label_q1)
        self.label_q1.setCursor(Qt.CursorShape.PointingHandCursor)
        layout_q2 = self.findChild(QGridLayout, "layout_q2")
        self.label_q2 = QLabel("")
        self.label_q2.setPixmap(pixmap)
        layout_q2.addWidget(self.label_q2)
        self.label_q2.setCursor(Qt.CursorShape.PointingHandCursor)

        #question mark instructions
        self.instructionsExcel = """Perform actions on the currently selected cell(s) containing ID(s) in the Pub-Xel Excel Worksheet.

- Import PubMed Data: Fetch article data from PubMed and fill the worksheet.

- Check if Files Exist: Verify if article files (e.g., PDFs) are present in the library folder. Cells of IDs that have corresponding files will be colored blue. For example, cell "33301246" will be colored blue if a file "33301246.pdf" exists in the library folder. Cells that do not have files will remain unchanged."""

        self.instructionsClipboard = """After copying (Ctrl+c) article ID(s), perform actions on them.

- Open Files: From the library folder, open saved article files (e.g., PDFs) corresponding to the copied IDs.

- Inspect Files: Open a window that provides an overview and functionalities for the copied IDs.

*Using global hotkeys, you can open saved article files while Pub-Xel is running in the background. One usage is to open files directly from a Pub-Xel Excel Worksheet. For example, select cells containing IDs (e.g., "33301246"), copy the cells (Ctrl+c), and press the global hotkey (Alt+Shift+k by default) to open the corresponding PDF files (e.g., "33301246.pdf") in the library folder.
*The same steps can be used to open files directly from PubMed IDs on platforms other than Microsoft Excel, including Microsoft Word, Microsoft PowerPoint, and the PubMed website."""

        self.label_q1.mousePressEvent = self.show_popup_instructionsExcel
        self.label_q2.mousePressEvent = self.show_popup_instructionsClipboard


        self.plainTextEdit_excel_current = self.findChild(QPlainTextEdit, 'plainTextEdit_excel_current')
        self.plainTextEdit_clipboard_current = self.findChild(QPlainTextEdit, 'plainTextEdit_clipboard_current')

        self.excelWorker = excelWorker()
        self.excelWorker.excel_updated.connect(self.update_excel_text)
        self.excelWorker.finished.connect(self.excelWorker_finished)
        self.excelWorker_running = False

        self.findChild(QAction, 'actionExit').triggered.connect(self.close_application)
        self.findChild(QAction, 'actionMinimize').triggered.connect(self.minimize_to_tray)
        if rt.settings.get('esc_to_system_tray',0):
            shortcut_Esc = QShortcut(QKeySequence('Esc'), self)
            shortcut_Esc.activated.connect(self.minimize_to_tray) 
        self.findChild(QAction, 'actionOpen_Library_Folder').triggered.connect(lambda: try_open_directory(rt.mainlibdir))
        self.findChild(QAction, 'actionOpen_Output_Folder').triggered.connect(lambda: try_open_directory(rt.outdir))
        self.findChild(QAction, 'actionNew_Excel_Template').triggered.connect(self.save_pubsheet)
        open_action = self.findChild(QAction, 'actionOpen')
        if open_action is not None:
            open_action.triggered.connect(self.open_file_dialog)
        self.findChild(QAction, 'actionPreferences').triggered.connect(self.open_preferences)
        self.findChild(QAction, 'actionAbout').triggered.connect(self.open_about_window)
        self.findChild(QAction, 'actionWorksheetColumns').triggered.connect(self.open_worksheetColumns_window)

        self.menu_file = self.findChild(QMenu, "menuFile")
        set_recent_menu_rebuild_callback(self.rebuild_recent_worksheet_menu)
        self.rebuild_recent_worksheet_menu()

        self.setStyleSheet("""QGroupBox#groupBoxExcel {font-size: 14px;}
                           QGroupBox#groupBoxClipboard {font-size: 14px;}""")


        self.tray_icon = SystemTrayIcon(self)
        self.tray_icon.show()

        self._nbib_dialog = None
        self._nbib_export_built.connect(self._on_nbib_export_built)
        self._nbib_export_failed.connect(self._on_nbib_export_failed)
        self._install_file_drop_targets()

    def button3_clicked(self):
        def on_press(key):
            try:
                if key.char == 'a':  # Replace 'a' with the key you want to listen for
                    print("Key 'a' pressed")
            except AttributeError:
                pass
        def on_release(key):
            if key == keyboard.Key.esc:
                # Stop listener
                return False
        # Collect events until released
        listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        listener.start()
        # Stop the listener after 1 second
        def stop_listener():
            listener.stop()
            print("Listener stopped after 1 second")
        timer = threading.Timer(1.0, stop_listener)
        timer.start()

    def update_excel_current(self):
        if not self.excelWorker_running:
            # print("excelWorker called, starting now")
            self.excelWorker_running = True
            self.excelWorker.start()
        else:
            print("excelWorker called but already running")

    def excelWorker_finished(self):
        # print("excelWorker_finished")
        self.excelWorker_running = False
        
    def update_excel_text(self, text):
        self.plainTextEdit_excel_current.setPlainText(text)

    def update_clipboard_current(self):
        clip = read_clipboard()
        self.plainTextEdit_clipboard_current.setPlainText(clip.preview)

    def showEvent(self, event):
        super().showEvent(event)
        self.update_excel_current()
        self.update_clipboard_current()
    def changeEvent(self, event):
        super().changeEvent(event)
        if event.type() == QtCore.QEvent.Type.WindowStateChange:
            if self.windowState() & QtCore.Qt.WindowState.WindowMaximized:
                self.update_excel_current()
                self.update_clipboard_current()
    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_excel_current()
        self.update_clipboard_current()
    def focusInEvent(self, event):
        super().focusInEvent(event)
        self.update_excel_current()
        self.update_clipboard_current()
    def event(self, event):
        if event.type() == QtCore.QEvent.Type.WindowActivate:
            print("WindowActivate event detected")
            self.update_excel_current()
            self.update_clipboard_current()
        return super().event(event)

    def on_application_state_changed(self, state):# Handle application state changes (e.g., dock/taskbar icon clicked)
        if rt.os_name == "Darwin":
            if state == Qt.ApplicationState.ApplicationActive:
                self.show()
                self.activateWindow()
                print("Application activated from Dock")
        else:
            return

    def open_about_window(self):
        if not rt.try_begin_action():
            return
        self.setEnabled(False)
        try:
            dialog = window_about(self)
            dialog.exec()
        finally:
            self.setEnabled(True)
            rt.end_action()

    def open_worksheetColumns_window(self):
        # Uses a separate guard so it can coexist with other actions.
        if rt.worksheetColumns_in_progress:
            return
        rt.worksheetColumns_in_progress = True
        try:
            dialog = window_worksheetColumns(self)
            dialog.exec()
        finally:
            rt.worksheetColumns_in_progress = False

    def open_preferences(self):
        if not rt.try_begin_action():
            return
        self.setEnabled(False)
        try:
            dialog = window_preferences(self)
            dialog.exit_main_window.connect(self.handle_exit_main_window)
            dialog.exec()
            if self.exitstatus:
                self.close_application()
        finally:
            self.setEnabled(True)
            rt.end_action()

    def handle_exit_main_window(self):
        self.exitstatus = True

    def _recent_menu_anchor(self) -> QAction | None:
        """Second separator below the recent-worksheet slot (above Minimize)."""
        menu = self.menu_file
        if menu is None:
            return None
        found_output = False
        sep_count = 0
        for action in menu.actions():
            if action.objectName() == "actionOpen_Output_Folder":
                found_output = True
                continue
            if found_output and action.isSeparator():
                sep_count += 1
                if sep_count == 2:
                    return action
        return None

    def rebuild_recent_worksheet_menu(self) -> None:
        menu = self.menu_file
        if menu is None:
            return
        for action in list(menu.actions()):
            if action.property("pubxel_recent_worksheet"):
                menu.removeAction(action)
        anchor = self._recent_menu_anchor()
        if anchor is None:
            return

        # Newest first in settings; each insertAction(anchor, …) stacks above prior inserts.
        for path in get_recent_worksheets():
            label = format_recent_menu_label(path)
            recent_action = QAction(label, self)
            recent_action.setProperty("pubxel_recent_worksheet", True)
            recent_action.setToolTip(path)
            recent_action.triggered.connect(
                lambda checked=False, p=path: self._open_recent_worksheet(p)
            )
            menu.insertAction(anchor, recent_action)

    def _open_recent_worksheet(self, path: str) -> None:
        if os.path.isfile(path):
            try_open_directory(path)
            register_recent_worksheet(path)
            return
        msg = QMessageBox(self)
        msg.setWindowTitle("Recent Worksheet")
        msg.setText("No file located. Remove history?")
        msg.setStandardButtons(
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        msg.setDefaultButton(QMessageBox.StandardButton.Yes)
        if msg.exec() == QMessageBox.StandardButton.Yes:
            remove_recent_worksheet(path)

    def save_pubsheet(self):
        if not rt.try_begin_action():
            return
        self.setEnabled(False)
        try:
            file_dialog = QFileDialog(self, "Save Worksheet")
            file_dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
            file_dialog.setDefaultSuffix("xlsx")
            file_dialog.setNameFilters(["Excel Files (*.xlsx)", "All Files (*)"])

            save_dir = os.path.expanduser("~/Documents")
            try:
                if not os.path.isdir(save_dir):
                    save_dir = os.path.expanduser("~")
                file_dialog.setDirectory(save_dir)
            except Exception as e:
                print(f"Failed to set initial directory: {e}")
                save_dir = os.path.expanduser("~")

            file_dialog.selectFile(default_worksheet_save_name(save_dir))

            if file_dialog.exec() == QFileDialog.DialogCode.Accepted:
                file_path = file_dialog.selectedFiles()[0]
                try:
                    create_pubsheet_worksheet(file_path, settings=rt.settings)
                    self.save_success_dialog(file_path)
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Failed to save file: {str(e)}")
        finally:
            self.setEnabled(True)
            rt.end_action()

    def crash(self):
        raise Exception("Crash")
    
    def save_success_dialog(self, file_path):
        show_worksheet_saved_dialog(self, file_path)

    def _local_paths_from_mime(self, mime) -> list[str]:
        if not mime.hasUrls():
            return []
        paths: list[str] = []
        for url in mime.urls():
            if url.isLocalFile():
                paths.append(url.toLocalFile())
        return paths

    def _first_openable_path(self, paths: list[str]) -> str | None:
        for path in paths:
            if os.path.isfile(path) and os.path.splitext(path)[1].lower() in _OPEN_FILE_EXTENSIONS:
                return path
        return None

    def _first_droppable_path(self, mime) -> str | None:
        return self._first_openable_path(self._local_paths_from_mime(mime))

    def _install_file_drop_targets(self) -> None:
        self.setAcceptDrops(True)
        for widget in [self, *self.findChildren(QWidget)]:
            widget.installEventFilter(self)

    def eventFilter(self, watched, event) -> bool:
        if event.type() in (QEvent.Type.DragEnter, QEvent.Type.DragMove):
            if self._first_droppable_path(event.mimeData()):
                event.acceptProposedAction()
                return True
            return False
        if event.type() == QEvent.Type.Drop:
            file_path = self._first_droppable_path(event.mimeData())
            if not file_path:
                return False
            if not rt.try_begin_action():
                return True
            event.acceptProposedAction()
            self._process_open_file_path(file_path)
            return True
        return super().eventFilter(watched, event)

    def open_file_dialog(self):
        if not rt.try_begin_action():
            return
        save_dir = default_worksheet_save_directory()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open File",
            save_dir,
            "Excel Worksheets & nbib (*.xlsx *.nbib);;"
            "Excel Worksheets (*.xlsx);;"
            "PubMed nbib (*.nbib);;"
            "All Files (*)",
        )
        if not file_path:
            rt.end_action()
            return
        self._process_open_file_path(file_path)

    def _process_open_file_path(self, file_path: str) -> None:
        """Handle Open or nbib drag-drop. Caller must hold rt action lock."""
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".xlsx":
            try:
                try_open_directory(file_path)
                register_recent_worksheet(file_path)
            finally:
                rt.end_action()
            return

        if ext == ".nbib":
            self._start_nbib_import(file_path)
            return

        dialog_onebutton(self, "Unsupported file type.", "Open File")
        rt.end_action()

    def _start_nbib_import(self, file_path: str) -> None:
        try:
            records, pmids, count = load_nbib_file(file_path)
        except ValueError as e:
            dialog_onebutton(self, str(e), "Open File")
            rt.end_action()
            return

        choice_dialog = NbibImportDialog(self, count)
        if choice_dialog.exec() != QDialog.DialogCode.Accepted:
            rt.end_action()
            return

        choice = choice_dialog.choice()
        if choice == NbibImportChoice.CANCEL:
            rt.end_action()
            return

        self.setEnabled(False)
        self._nbib_dialog = RunningFunctionDialog(
            parent=self,
            message="Importing PubMed data...\nPlease wait.",
        )
        thread = threading.Thread(
            target=self._nbib_import_worker,
            args=(records, pmids, choice),
            daemon=True,
        )
        thread.start()

    def _close_nbib_dialog(self) -> None:
        if self._nbib_dialog is not None:
            try:
                self._nbib_dialog.close()
            except Exception:
                pass
            self._nbib_dialog = None

    def _nbib_import_worker(
        self,
        records: list[str],
        pmids: list[str],
        choice: str,
    ) -> None:
        temp_path = None
        try:
            metadata = import_nbib_to_metadata(records)
            suffix = ".xlsx" if choice == NbibImportChoice.EXCEL else ".tsv"
            fd, temp_path = tempfile.mkstemp(suffix=suffix)
            os.close(fd)
            if choice == NbibImportChoice.EXCEL:
                create_filled_worksheet(
                    temp_path,
                    pmids,
                    metadata,
                    settings=rt.settings,
                )
            else:
                write_worksheet_tsv(temp_path, pmids, metadata, rt.settings)
            self._nbib_export_built.emit(temp_path, choice)
        except Exception as e:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
            self._nbib_export_failed.emit(str(e))

    def _on_nbib_export_built(self, temp_path: str, choice: str) -> None:
        self._close_nbib_dialog()
        self.setEnabled(True)

        save_dir = default_worksheet_save_directory()
        if choice == NbibImportChoice.EXCEL:
            file_dialog = QFileDialog(self, "Save Worksheet")
            file_dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
            file_dialog.setDefaultSuffix("xlsx")
            file_dialog.setNameFilters(["Excel Files (*.xlsx)", "All Files (*)"])
            file_dialog.setDirectory(save_dir)
            file_dialog.selectFile(default_worksheet_save_name(save_dir))
        else:
            file_dialog = QFileDialog(self, "Save TSV File")
            file_dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
            file_dialog.setDefaultSuffix("tsv")
            file_dialog.setNameFilters(["TSV Files (*.tsv)", "All Files (*)"])
            file_dialog.setDirectory(save_dir)
            file_dialog.selectFile(default_tsv_save_name(save_dir))

        dest_path = None
        if file_dialog.exec() == QFileDialog.DialogCode.Accepted:
            dest_path = file_dialog.selectedFiles()[0]

        try:
            if dest_path:
                shutil.copy2(temp_path, dest_path)
                if choice == NbibImportChoice.EXCEL:
                    show_worksheet_saved_dialog(self, dest_path)
                else:
                    show_file_saved_dialog(self, dest_path, open_label="Open File")
        except Exception as e:
            dialog_onebutton(self, f"Failed to save file:\n{e}", "Error")
        finally:
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except OSError:
                pass
            rt.end_action()

    def _on_nbib_export_failed(self, message: str) -> None:
        self._close_nbib_dialog()
        self.setEnabled(True)
        rt.end_action()
        dialog_onebutton(self, f"Failed to import nbib file:\n{message}", "Error")

    def run_check_file_exist2(self):
        if not rt.try_begin_action():
            return
        self.setEnabled(False)
        self.dialog = RunningFunctionDialog(
            parent=self,
            message="Checking file existence...\nPlease do not interact with Excel.",
        )
        try:
            def proceed():
                print("0.5 seconds have passed!")
            timer = QTimer()
            timer.setSingleShot(True)
            timer.timeout.connect(proceed)
            timer.start(500)
            message = self.check_file_exist2()
            if message:
                dialog_onebutton(self.dialog, str(message), "Confirmation")
        finally:
            self.dialog.close()
            self.setEnabled(True)
            rt.end_action()

    def run_input_pubmed_data2(self):
        if not rt.try_begin_action():
            return
        self.setEnabled(False)
        self.dialog = RunningFunctionDialog(
            parent=self,
            message="Importing PubMed data...\nPlease do not interact with Excel.",
        )
        try:
            def proceed():
                print("0.5 seconds have passed!")
            timer = QTimer()
            timer.setSingleShot(True)
            timer.timeout.connect(proceed)
            timer.start(500)
            message = self.input_pubmed_data2()
            if message:
                dialog_onebutton(self.dialog, str(message), "Confirmation")
        finally:
            self.dialog.close()
            self.setEnabled(True)
            rt.end_action()

    def check_file_exist2(self):
        try: 
            message = check_file_exist(rt.mainlibdir,rt.seclibdir)
        except Exception as e:
            return e
        return message

    def input_pubmed_data2(self):
        try: 
            check_file_exist(rt.mainlibdir,rt.seclibdir)
        except Exception as e:
            return e
        try:
            message = input_pubmed_data()
        except Exception as e:
            return e
        return message

    def action_in_progress_switch(self):
        # Developer-mode toggle: flip the flag in place. Bypasses the lock
        # because this is a debugging affordance only available in developer
        # mode and is not safe to combine with real actions.
        if rt.action_in_progress:
            rt.end_action()
        else:
            rt.try_begin_action()
        print(rt.action_in_progress)

    def mousePressEvent(self, event):
        if rt.developerMode:
            if event.button() == Qt.MouseButton.RightButton:
                texts_with_links = [
                    ("link", "https://www.google.com")
                ]
                self.show_popuptest(event.globalPosition().toPoint(),texts_with_links)
        else:
            pass

    def show_popuptest(self, pos, texts_with_links):
        self.popup = PopupWidgettest(texts_with_links)
        self.popup.move(pos)
        self.popup.show()

    def show_popup(self, source, text):
        self.popup = PopupInstructions(text)
        self.popup.move(source.mapToGlobal(source.rect().bottomLeft()) + QtCore.QPoint(-200, 0))
        self.popup.show()

    def show_popup_instructionsExcel(self, event):
        self.show_popup(self.label_q1, self.instructionsExcel)

    def show_popup_instructionsClipboard(self, event):
        self.show_popup(self.label_q2,self.instructionsClipboard)

    def minimize_to_tray(self):# dont change self.is_closing and rt.system_tray_notice_shown orders
# global rt.system_tray_notice_shown -> rt.system_tray_notice_shown
# global rt.settings -> rt.settings
        self.hide()
        if self.is_closing:
            return
        if rt.system_tray_notice_shown:
            return
        else:
            try:
                self.tray_icon.showMessage(
                    "Notice",
                    "Pub-Xel will continue running in the background. This behavior can be changed in the Preferences menu.",
                    QSystemTrayIcon.MessageIcon.Information,
                    5000  # Duration in milliseconds
                    )
            except Exception as e:
                print(f"Failed to show system tray notice: {e}")
            rt.system_tray_notice_shown += 1
            rt.settings = save_settings_key(rt.settings, 'system_tray_notice_shown', 1)



    # if rt.settings.get('close_to_system_tray', 0):
    #     def closeEvent(self, event):
    #         event.ignore()
    #         self.minimize_to_tray()
    # else:
    #     def closeEvent(self, event):
    #         print("closeEvent")
    #         QApplication.quit()
    #         event.accept()

    def closeEvent(self, event):
        # If Qt or the installer is closing us (not the user clicking X), accept and quit.
        # event.spontaneous() is True for user-driven events.
        # We accept when:
        #   - rt.force_quit is True (we called close_application / QApplication.quit)
        #   - not event.spontaneous() (programmatic close like WM_CLOSE from installer/OS)
        if rt.force_quit or not event.spontaneous():
            # QApplication.quit()
            event.accept()
            return

        # Otherwise apply "close to tray" behavior only for user-initiated closes
        if rt.settings.get('close_to_system_tray', 0):
            event.ignore()
            self.minimize_to_tray()
        else:
            # QApplication.quit()
            event.accept()


    # def close_application(self):
    #     self.is_closing = True
    #     print("close_application")
    #     QApplication.quit()

    def close_application(self):
# global rt.force_quit -> rt.force_quit
        self.is_closing = True
        rt.force_quit = True
        try:
            graceful_shutdown()
        except Exception:
            pass
        QApplication.quit()
                                          
    def sayHello(self,event=None):
        print("Hello")

    def main_openfile(self):
        if not rt.try_begin_action():
            return
        print("opening files from main window...")
        self.setEnabled(False)
        try:
            clip = read_clipboard()
            # Open Files: file_paths only for Explorer copy; otherwise ids -> library.
            # See pubxel_core/clipboard.py module docstring.
            if clip.kind == "files":
                self.open_paths_from_list(clip.file_paths)
            else:
                msg = message_for_action(clip)
                if msg:
                    dialog_onebutton(self, msg, "Clipboard")
                    return
                output = process_ids(clip.ids, rt.mainlibdir, rt.seclibdir)
                self.openfile_from_list(output[10], output[14])
        finally:
            self.setEnabled(True)
            rt.end_action()

    def main_openpubmed(self):
        if not rt.try_begin_action():
            return
        print("opening pubmed link from main window...")
        self.setEnabled(False)
        try:
            clip = read_clipboard()
            # PubMed: ids only — never file_paths (see pubxel_core/clipboard.py).
            msg = message_for_action(clip)
            if msg:
                dialog_onebutton(self, msg, "Clipboard")
                return
            output = process_ids(clip.ids, rt.mainlibdir, rt.seclibdir)
            self.openpubmed_from_list(output[1])  # pubmed ids
        finally:
            self.setEnabled(True)
            rt.end_action()

    def open_paths_from_list(self, paths):
        """Open absolute file paths from a file-manager clipboard copy."""
        if not paths:
            dialog_onebutton(self, "No files to open.", "Error")
            return
        filepath_list = [p for p in paths if os.path.isfile(p)]
        if not filepath_list:
            dialog_onebutton(self, "No existing files to open on clipboard.", "Error")
            return
        if len(filepath_list) > 50:
            dialog_onebutton(self, "Cannot open more than 50 files.", "Error")
            return
        if len(filepath_list) > 5:
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("Warning")
            msg_box.setText(
                "You are about to open " + str(len(filepath_list)) + " files. Proceed?"
            )
            msg_box.setStandardButtons(
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            msg_box.setWindowModality(Qt.WindowModality.ApplicationModal)
            if msg_box.exec() != QMessageBox.StandardButton.Yes:
                return
        failed_files = []
        for filepath in filepath_list:
            try:
                open_directory(filepath)
            except Exception:
                failed_files.append(filepath)
        if failed_files:
            dialog_onebutton(
                self,
                "The following files could not be opened:\n" + "\n".join(failed_files),
                "Error",
            )

    def openfile_from_list(self, filelist, file_map=None):
        if filelist == []:
            dialog_onebutton(self,"No files to open.","Error")
            return
        filepathList = files_name_to_path(
            filelist,
            rt.mainlibdir,
            rt.seclibdir,
            file_map=file_map,
        )
        # If there are more than 5 files, show a warning and ask for confirmation

        if len(filepathList) > 50:
            dialog_onebutton(self,"Cannot open more than 50 files.","Error")
            return
        
        if len(filepathList) > 5:
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("Warning")
            msg_box.setText("You are about to open " + str(len(filepathList)) + " files. Proceed?")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg_box.setWindowModality(Qt.WindowModality.ApplicationModal)
            proceed = msg_box.exec() == QMessageBox.StandardButton.Yes
            if not proceed:
                return

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



    def openpubmed_from_list(self, filelist): #filelist: list of pubmed ids
        if filelist == []: 
            dialog_onebutton(self,"No PubMed articles to open.","Error")
            return
        
        if len(filelist) > 500:
            dialog_onebutton(self,"Cannot open more than 500 files.","Error")
            return
        
        url = "https://pubmed.ncbi.nlm.nih.gov/?term=" + "%5Buid%5D+OR+".join(filelist) + "%5Buid%5D&sort=date"
        webbrowser.open(url, new=2)

    



    def open_inspect_window(self):
        # window_inspect is a *non-modal* QWidget — it returns from __init__
        # immediately. Once it's on screen, the inspect window owns the
        # action_in_progress flag and re-enables the main window itself in
        # its closeEvent. Releasing here would let the user open multiple
        # inspect windows at once.
        if not rt.try_begin_action():
            return
        self.setEnabled(False)
        transferred_ownership = False
        try:
            clip = read_clipboard()
            # Inspect: ids only — never file_paths (see pubxel_core/clipboard.py).
            msg = message_for_action(clip)
            if msg:
                dialog_onebutton(self, msg, "Clipboard")
                return
            selected_cell_value = clip.ids
            if len(selected_cell_value) == 0:
                dialog_onebutton(self, "No ID(s) selected.", "Error")
                return
            if len(selected_cell_value) > 200:
                dialog_onebutton(self, "Please select 200 or fewer ID(s).", "Error")
                return
            print("opening inspect window...")
            window_inspect(self, data=None, clipboard_ids=selected_cell_value)
            transferred_ownership = True
        finally:
            if not transferred_ownership:
                self.setEnabled(True)
                rt.end_action()
