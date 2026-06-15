# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>

import copy
import logging
import os
from PyQt6 import QtCore, uic
from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QTextCursor
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QTabBar,
    QWidget,
)

from pubxel_core import runtime as rt
from pubxel_core.clipboard import write_clipboard
from pubxel_core.settings import save_settings
from pubxel_core.ui.helpers import dialog_onebutton

class window_preferences(QDialog):
    exit_main_window = pyqtSignal()
    def __init__(self, parent):
        # The action_in_progress flag and parent.setEnabled() are owned by
        # the caller (open_preferences) — see pubxel_core/runtime.py.
        super().__init__(parent)
        uic.loadUi(rt.preferences_path, self)
        self.pseudo_parent = parent

        self.setWindowTitle('Preferences')
        self.settings = rt.settings
        self.tabWidget.setCurrentIndex(0)
        self.tab_hot = self.findChild(QWidget, 'tab_hot')
        if self.tab_hot and rt.os_name == "Darwin":
            logging.getLogger(__name__).debug("Hotkeys tab disabled in MacOS")
            index_to_hide = self.tabWidget.indexOf(self.tab_hot)
            self.tabWidget.removeTab(index_to_hide)

        self.tab_other = self.findChild(QWidget, "tab_other")
        self.tab_drive = self.findChild(QWidget, "tab_drive")
        self.checkBox_developerMode = self.findChild(QCheckBox, "checkBox_developerMode")
        self._dev_unlock_clicks = 0
        self._dev_unlock_timer = QTimer(self)
        self._dev_unlock_timer.setSingleShot(True)
        self._dev_unlock_timer.timeout.connect(self._reset_dev_unlock_clicks)
        self._configure_developer_mode_ui()
        self.tabWidget.tabBarClicked.connect(self._on_preferences_tab_clicked)

        # library tab
        self.plainTextEdit_mainlib = self.findChild(QPlainTextEdit, 'plainTextEdit_mainlib')
        self.plainTextEdit_mainlib.setPlainText(self.settings.get('mainlib_path', ""))
        self.button_selectmainlib = self.findChild(QPushButton, 'button_selectmainlib')
        self.button_selectmainlib.clicked.connect(self.setmainlib)

        self.checkbox_mainlib_include_subfolders = self.findChild(
            QCheckBox, 'checkbox_mainlib_include_subfolders'
        )
        if self.checkbox_mainlib_include_subfolders is not None:
            self.checkbox_mainlib_include_subfolders.setChecked(
                bool(self.settings.get('mainlib_include_subfolders', 0))
            )

        self.plainTextEdit_output = self.findChild(QPlainTextEdit, 'plainTextEdit_output')
        self.plainTextEdit_output.setPlainText(self.settings.get('output_path', ""))
        self.button_selectoutput = self.findChild(QPushButton, 'button_selectoutput')
        self.button_selectoutput.clicked.connect(self.setoutput)

        self.plainTextEdit_seclib = self.findChild(QPlainTextEdit, 'plainTextEdit_seclib')
        seclib_paths = self.settings.get('seclib_path', [])
        self.plainTextEdit_seclib.setPlainText('\n'.join(seclib_paths))
        self.button_selectseclib = self.findChild(QPushButton, 'button_selectseclib')
        self.button_selectseclib.clicked.connect(self.setseclib)

        self.groupBox_seclib = self.findChild(QGroupBox, 'groupBox_seclib')
        self.groupBox_seclib.setChecked(self.settings.get('seclib_enable',0))

        self.button_libdefault = self.findChild(QPushButton, 'button_libdefault')
        self.button_libdefault.clicked.connect(self.libdefault)

        # worksheet tab
        self.gridLayout_worksheet_columns = self.findChild(
            QGridLayout, "gridLayout_worksheet_columns"
        )
        if self.gridLayout_worksheet_columns is not None:
            self._populate_worksheet_columns_grid()

        #hotkey tab
        self.hotkey_strings = [
            # "<Ctrl>+",
            "<Ctrl>+<Shift>+",
            # "<Alt>+",
            "<Alt>+<Shift>+",
            "<Ctrl>+<Alt>+",
            "<Ctrl>+<Alt>+<Shift>+"
        ]
        self.textboxcurrent = False

        #hotkey tab - open
        self.hotkeyvalue_open = ""
        self.layout_open = self.findChild(QGridLayout, "layout_open")
        self.checkboxes_open = []
        self.textboxes_open = []
        self.groupBox_open = self.findChild(QGroupBox, "groupBox_open")
        for i, hotkey_string in enumerate(self.hotkey_strings):
            checkbox = QCheckBox(hotkey_string)
            checkbox.setContentsMargins(6, 0, 6, 0)
            checkbox.stateChanged.connect(lambda: self.update_hotkeyvalue_open())
            self.layout_open.addWidget(checkbox, i, 0)
            self.checkboxes_open.append(checkbox)
            textbox = QPlainTextEdit()
            textbox.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            textbox.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            textbox.setContentsMargins(6, 0, 6, 0)
            textbox.setFixedHeight(25)
            textbox.setFixedWidth(50)
            textbox.setEnabled(False)
            textbox.textChanged.connect(lambda: self.update_hotkeyvalue_open())
            self.layout_open.addWidget(textbox, i, 1)
            self.textboxes_open.append(textbox)
        hotkey_open_value = rt.settings.get('hotkey_open_value', "")
        if not hotkey_open_value == "":
            self.groupBox_open.setChecked(True)
            for i, hotkey_string in enumerate(self.hotkey_strings):
                if hotkey_open_value[:-1] == hotkey_string:
                    self.checkboxes_open[i].setChecked(True)
                    self.textboxes_open[i].setEnabled(True)
                    self.textboxes_open[i].setPlainText(hotkey_open_value[len(hotkey_string):])
                    break

        #hotkey tab - inspect
        self.hotkeyvalue_inspect = ""
        self.layout_inspect = self.findChild(QGridLayout, "layout_inspect")
        self.checkboxes_inspect = []
        self.textboxes_inspect = []
        self.groupBox_inspect = self.findChild(QGroupBox, "groupBox_inspect")
        for i, hotkey_string in enumerate(self.hotkey_strings):
            checkbox = QCheckBox(hotkey_string)
            checkbox.setContentsMargins(6, 0, 6, 0)
            checkbox.stateChanged.connect(lambda: self.update_hotkeyvalue_inspect())
            self.layout_inspect.addWidget(checkbox, i, 0)
            self.checkboxes_inspect.append(checkbox)
            textbox = QPlainTextEdit()
            textbox.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            textbox.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            textbox.setContentsMargins(6, 0, 6, 0)
            textbox.setFixedHeight(25)
            textbox.setFixedWidth(50)
            textbox.setEnabled(False)
            textbox.textChanged.connect(lambda: self.update_hotkeyvalue_inspect())
            self.layout_inspect.addWidget(textbox, i, 1)
            self.textboxes_inspect.append(textbox)
        hotkey_inspect_value = rt.settings.get('hotkey_inspect_value', "")
        if not hotkey_inspect_value == "":
            self.groupBox_inspect.setChecked(True)
            for i, hotkey_string in enumerate(self.hotkey_strings):
                if hotkey_inspect_value[:-1] == hotkey_string:
                    self.checkboxes_inspect[i].setChecked(True)
                    self.textboxes_inspect[i].setEnabled(True)
                    self.textboxes_inspect[i].setPlainText(hotkey_inspect_value[len(hotkey_string):])
                    break


        #hotkey tab - pubmed
        self.hotkeyvalue_pubmed = ""
        self.layout_pubmed = self.findChild(QGridLayout, "layout_pubmed")
        self.checkboxes_pubmed = []
        self.textboxes_pubmed = []
        self.groupBox_pubmed = self.findChild(QGroupBox, "groupBox_pubmed")
        for i, hotkey_string in enumerate(self.hotkey_strings):
            checkbox = QCheckBox(hotkey_string)
            checkbox.setContentsMargins(6, 0, 6, 0)
            checkbox.stateChanged.connect(lambda: self.update_hotkeyvalue_pubmed())
            self.layout_pubmed.addWidget(checkbox, i, 0)
            self.checkboxes_pubmed.append(checkbox)
            textbox = QPlainTextEdit()
            textbox.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            textbox.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            textbox.setContentsMargins(6, 0, 6, 0)
            textbox.setFixedHeight(25)
            textbox.setFixedWidth(50)
            textbox.setEnabled(False)
            textbox.textChanged.connect(lambda: self.update_hotkeyvalue_pubmed())
            self.layout_pubmed.addWidget(textbox, i, 1)
            self.textboxes_pubmed.append(textbox)
        hotkey_pubmed_value = rt.settings.get('hotkey_pubmed_value', "")
        if not hotkey_pubmed_value == "":
            self.groupBox_pubmed.setChecked(True)
            for i, hotkey_string in enumerate(self.hotkey_strings):
                if hotkey_pubmed_value[:-1] == hotkey_string:
                    self.checkboxes_pubmed[i].setChecked(True)
                    self.textboxes_pubmed[i].setEnabled(True)
                    self.textboxes_pubmed[i].setPlainText(hotkey_pubmed_value[len(hotkey_string):])
                    break

        # hotkey tab - restore default
        self.button_hotdefault = self.findChild(QPushButton, 'button_hotdefault')
        self.button_hotdefault.clicked.connect(self.hotdefault)

        #exit tab and exit
        self.button_saveexit = self.findChild(QPushButton, 'button_saveexit')
        self.button_saveexit.clicked.connect(self.saveexit)
        self.button_cancel = self.findChild(QPushButton, 'button_cancel')
        self.button_cancel.clicked.connect(self.close_preferences_window)
        self.show()

        #others tab
        # self.checkBox_launch_at_startup = self.findChild(QCheckBox, 'checkBox_launch_at_startup')
        # if rt.settings.get('launch_at_startup',0):
        #     self.checkBox_launch_at_startup.setChecked(True)
        # # self.settings['launch_at_startup'] = 1 if self.checkBox_launch_at_startup.isChecked() else 0
        self.checkBox_close_to_system_tray = self.findChild(QCheckBox, 'checkBox_close_to_system_tray')
        if rt.settings.get('close_to_system_tray',0):
            self.checkBox_close_to_system_tray.setChecked(True)
        self.checkBox_esc_to_system_tray = self.findChild(QCheckBox, 'checkBox_esc_to_system_tray')
        if rt.settings.get('esc_to_system_tray',0):
            self.checkBox_esc_to_system_tray.setChecked(True)

    def _configure_developer_mode_ui(self) -> None:
        if self.tab_drive is not None and not rt.developerMode:
            drive_index = self.tabWidget.indexOf(self.tab_drive)
            if drive_index >= 0:
                self.tabWidget.removeTab(drive_index)

        if self.checkBox_developerMode is not None:
            self.checkBox_developerMode.setChecked(bool(self.settings.get("developerMode", 0)))
            self.checkBox_developerMode.setVisible(bool(rt.developerMode))

    def _reset_dev_unlock_clicks(self) -> None:
        self._dev_unlock_clicks = 0

    def _reveal_developer_mode_checkbox(self) -> None:
        if self.checkBox_developerMode is None:
            return
        self.checkBox_developerMode.setVisible(True)
        if self.tab_other is not None:
            self.tabWidget.setCurrentWidget(self.tab_other)
        logging.getLogger(__name__).debug("Developer mode checkbox revealed via unlock gesture")

    def _on_preferences_tab_clicked(self, index: int) -> None:
        if self.tab_other is None or self.tabWidget.widget(index) is not self.tab_other:
            return
        mods = QApplication.keyboardModifiers()
        if not (
            mods & Qt.KeyboardModifier.ControlModifier
            and mods & Qt.KeyboardModifier.ShiftModifier
        ):
            return

        self._dev_unlock_clicks += 1
        self._dev_unlock_timer.start(3000)
        if self._dev_unlock_clicks >= 5:
            self._dev_unlock_clicks = 0
            self._dev_unlock_timer.stop()
            self._reveal_developer_mode_checkbox()

    def setmainlib(self):  
        folder_path = self.select_folder()
        if folder_path:
            self.plainTextEdit_mainlib.setPlainText(folder_path)

    def setseclib(self):  
        folder_path = self.select_folder()
        if folder_path:
            current_text = self.plainTextEdit_seclib.toPlainText()
            current_text = current_text.replace('\n\n', '\n')
            current_text = current_text.rstrip('\n') + '\n'
            current_text = current_text.lstrip('\n')
            new_text = current_text + folder_path + '\n'
            self.plainTextEdit_seclib.setPlainText(new_text)
            self.plainTextEdit_seclib.moveCursor(QTextCursor.MoveOperation.End)

    def setoutput(self):  
        folder_path = self.select_folder()
        if folder_path:
            self.plainTextEdit_output.setPlainText(folder_path)

    def select_folder(self):
        file_dialog = QFileDialog(self, "Select Folder")
        file_dialog.setFileMode(QFileDialog.FileMode.Directory)
        file_dialog.setOption(QFileDialog.Option.ShowDirsOnly, True)

        if file_dialog.exec() == QFileDialog.DialogCode.Accepted:
            folder_path = file_dialog.selectedFiles()[0]
            if rt.os_name == 'Windows':
                folder_path = folder_path.replace("/", "\\")
            logging.getLogger(__name__).debug("Selected folder: %s", folder_path)
            return folder_path
        else:
            return None

    def update_hotkeyvalue_open(self):
        sender = self.sender()
        if isinstance(sender, QCheckBox) and not sender.testAttribute(Qt.WidgetAttribute.WA_UnderMouse):
            return
        if self.textboxcurrent:
            return
        sender = self.sender()
        if isinstance(sender, QCheckBox):
            for checkbox in self.checkboxes_open:
                if checkbox != sender:
                    checkbox.setChecked(False)
            index = self.checkboxes_open.index(sender)
            logging.getLogger(__name__).debug("open checkbox index=%s", index)
            self.textboxes_open[index].setEnabled(sender.isChecked())
            for textbox in self.textboxes_open:
                if textbox != self.textboxes_open[index]:
                    textbox.setEnabled(False)
        elif isinstance(sender, QPlainTextEdit):
            self.textboxcurrent = True
            logging.getLogger(__name__).debug("open textboxcurrent=True")
            index = self.textboxes_open.index(sender)
            text = sender.toPlainText()
            logging.getLogger(__name__).debug("open textbox raw=%r", text)
            if len(text) > 1:
                sender.setPlainText(text[-1])
            text = sender.toPlainText()
            if not text.isalpha():
                sender.setPlainText("")
            sender.setPlainText(sender.toPlainText().lower())
            self.textboxcurrent = False
            logging.getLogger(__name__).debug("open textboxcurrent=False")

    def update_hotkeyvalue_inspect(self):
        sender = self.sender()
        if isinstance(sender, QCheckBox) and not sender.testAttribute(Qt.WidgetAttribute.WA_UnderMouse):
            return
        if self.textboxcurrent:
            return
        sender = self.sender()
        if isinstance(sender, QCheckBox):
            for checkbox in self.checkboxes_inspect:
                if checkbox != sender:
                    checkbox.setChecked(False)
            index = self.checkboxes_inspect.index(sender)
            logging.getLogger(__name__).debug("inspect checkbox index=%s", index)
            self.textboxes_inspect[index].setEnabled(sender.isChecked())
            for textbox in self.textboxes_inspect:
                if textbox != self.textboxes_inspect[index]:
                    textbox.setEnabled(False)
            
        elif isinstance(sender, QPlainTextEdit):
            self.textboxcurrent = True
            logging.getLogger(__name__).debug("inspect textboxcurrent=True")
            index = self.textboxes_inspect.index(sender)
            text = sender.toPlainText()
            logging.getLogger(__name__).debug("inspect textbox raw=%r", text)
            if len(text) > 1:
                sender.setPlainText(text[-1])
            text = sender.toPlainText()
            if not text.isalpha():
                sender.setPlainText("")
            sender.setPlainText(sender.toPlainText().lower())
            self.textboxcurrent = False
            logging.getLogger(__name__).debug("inspect textboxcurrent=False")

    def update_hotkeyvalue_pubmed(self):
        sender = self.sender()
        if isinstance(sender, QCheckBox) and not sender.testAttribute(Qt.WidgetAttribute.WA_UnderMouse):
            return
        if self.textboxcurrent:
            return
        sender = self.sender()
        if isinstance(sender, QCheckBox):
            for checkbox in self.checkboxes_pubmed:
                if checkbox != sender:
                    checkbox.setChecked(False)
            index = self.checkboxes_pubmed.index(sender)
            logging.getLogger(__name__).debug("pubmed checkbox index=%s", index)
            self.textboxes_pubmed[index].setEnabled(sender.isChecked())
            for textbox in self.textboxes_pubmed:
                if textbox != self.textboxes_pubmed[index]:
                    textbox.setEnabled(False)
        elif isinstance(sender, QPlainTextEdit):
            self.textboxcurrent = True
            logging.getLogger(__name__).debug("pubmed textboxcurrent=True")
            index = self.textboxes_pubmed.index(sender)
            text = sender.toPlainText()
            logging.getLogger(__name__).debug("pubmed textbox raw=%r", text)
            if len(text) > 1:
                sender.setPlainText(text[-1])
            text = sender.toPlainText()
            if not text.isalpha():
                sender.setPlainText("")
            sender.setPlainText(sender.toPlainText().lower())
            self.textboxcurrent = False
            logging.getLogger(__name__).debug("pubmed textboxcurrent=False")

    def libdefault(self):
        self.plainTextEdit_mainlib.setPlainText(rt.mainlibdirdefault)
        self.plainTextEdit_output.setPlainText(rt.outdirdefault)
        self.groupBox_seclib.setChecked(False)
        if self.checkbox_mainlib_include_subfolders is not None:
            self.checkbox_mainlib_include_subfolders.setChecked(False)

    def _populate_worksheet_columns_grid(self):
        headers = ("Column", "", "Definition", "Example", "Use")
        for col, text in enumerate(headers):
            lbl = QLabel(text)
            self.gridLayout_worksheet_columns.addWidget(lbl, 0, col)

        rows = [
            ("Ref", "PMID", "33301246"),
            ("DOI", "DOI", "10.1056/NEJMoa2034577"),
            ("AuthorYear", "Author, year", "Polack et al., 2020"),
            ("Authors", "Author list", "Polack FP, Thomas SJ, Kitchin N, Absalon J, Gurtman A, Lockhart S, ..."),
            ("Year", "Year", "2020"),
            ("Journal", "Journal", "N Engl J Med"),
            ("IF2024", "2024 Journal Impact Factor", "78.5"),
            ("Title", "Title", "Safety and Efficacy of the BNT162b2 mRNA Covid-19 Vaccine."),
            ("Abstract", "Abstract", "BACKGROUND: Severe acute respiratory syndrome coronavirus 2 ..."),
            ("Citation", "NLM-style citation", "Polack et al. Safety and Efficacy of the BNT162b2 mRNA Covid-19 Vaccine. N Engl J Med. 2020 Dec 31;383(27):2603-2615. PMID: 33301246."),
            ("Citation2024", "NLM-style citation with 2024 IF", "Polack et al. Safety and Efficacy of the BNT162b2 mRNA Covid-19 Vaccine. N Engl J Med (IF: 78.5). 2020 Dec 31;383(27):2603-2615. PMID: 33301246."),
            ("Q2024", "2024 Journal Quartile", "Q1"),
            ("Identifier", "Secondary identifier", "ClinicalTrials.gov/NCT04368728"),
            ("Funding", "Grant/funding", "—"),
        ]

        default_enabled = {
            "Ref": True,
            "Title": True,
            "AuthorYear": True,
            "Journal": True,
            "Abstract": True,
            "Citation": True,
            "IF2024": True,
        }
        locked_on = {"Ref", "Title"}
        settings_enabled = self.settings.get("worksheet_column_enabled", {})
        self.worksheet_column_checks = {}

        for row_idx, (col_name, definition, example) in enumerate(rows, start=1):
            label_col = QLabel(col_name)
            label_col.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            button_copy = QPushButton("Copy")
            button_copy.clicked.connect(lambda _, s=col_name: write_clipboard(s))
            label_def = QLabel(definition)
            label_ex = QLabel(example)
            label_ex.setWordWrap(True)
            label_ex.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            checkbox_use = QCheckBox()
            enabled_default = default_enabled.get(col_name, False)
            checked = bool(settings_enabled.get(col_name, 1 if enabled_default else 0))
            if col_name in locked_on:
                checked = True
                checkbox_use.setEnabled(False)
            checkbox_use.setChecked(checked)
            self.worksheet_column_checks[col_name] = checkbox_use
            self.gridLayout_worksheet_columns.addWidget(label_col, row_idx, 0)
            self.gridLayout_worksheet_columns.addWidget(button_copy, row_idx, 1)
            self.gridLayout_worksheet_columns.addWidget(label_def, row_idx, 2)
            self.gridLayout_worksheet_columns.addWidget(label_ex, row_idx, 3)
            self.gridLayout_worksheet_columns.addWidget(checkbox_use, row_idx, 4)

    def hotdefault(self):
        self.groupBox_open.setChecked(True)
        self.groupBox_inspect.setChecked(True)
        self.groupBox_pubmed.setChecked(True)
        for checkbox in self.checkboxes_open:
            checkbox.setChecked(False)
        for checkbox in self.checkboxes_inspect:
            checkbox.setChecked(False)
        for checkbox in self.checkboxes_pubmed:
            checkbox.setChecked(False)
        for textbox in self.textboxes_open:
            textbox.setPlainText("")
            textbox.setEnabled(False)
        for textbox in self.textboxes_inspect:
            textbox.setPlainText("")
            textbox.setEnabled(False)
        for textbox in self.textboxes_pubmed:
            textbox.setPlainText("")
            textbox.setEnabled(False)

        index = self.hotkey_strings.index("<Alt>+<Shift>+")
        self.checkboxes_open[index].setChecked(True)
        self.textboxes_open[index].setEnabled(True)
        self.textboxes_open[index].setPlainText("k")
        self.checkboxes_inspect[index].setChecked(True)
        self.textboxes_inspect[index].setEnabled(True)
        self.textboxes_inspect[index].setPlainText("j")
        self.checkboxes_pubmed[index].setChecked(True)
        self.textboxes_pubmed[index].setEnabled(True)
        self.textboxes_pubmed[index].setPlainText("p")

    def saveexit(self):
        # Never mutate rt.settings until validation succeeds.
        new_settings = copy.deepcopy(rt.settings)

        def _norm_path(p: str) -> str:
            return os.path.normcase(os.path.normpath(p.strip()))

        mainlib_path = self.plainTextEdit_mainlib.toPlainText()
        new_settings['mainlib_path'] = mainlib_path
        if self.checkbox_mainlib_include_subfolders is not None:
            new_settings['mainlib_include_subfolders'] = (
                1 if self.checkbox_mainlib_include_subfolders.isChecked() else 0
            )
        if not mainlib_path or not os.path.isdir(mainlib_path):
            dialog_onebutton(self, "Error in Library Settings.\nMain library path does not exist.", "Error")
            return

        output_path = self.plainTextEdit_output.toPlainText()
        new_settings['output_path'] = output_path
        if not output_path or not os.path.isdir(output_path):
            dialog_onebutton(self, "Error in Library Settings.\nOutput path does not exist.", "Error")
            return

        if _norm_path(output_path) == _norm_path(mainlib_path):
            dialog_onebutton(
                self,
                "Error in Library Settings.\nOutput path cannot be identical to main library path.",
                "Error",
            )
            return

        new_settings['seclib_path'] = sorted(
            set(
                line
                for line in self.plainTextEdit_seclib.toPlainText().split('\n')
                if line.strip()
            )
        )
        new_settings['seclib_enable'] = 1 if self.groupBox_seclib.isChecked() else 0

        # Validate secondary library directories only when enabled.
        if new_settings['seclib_enable'] and new_settings['seclib_path']:
            nopaths = [p for p in new_settings['seclib_path'] if not os.path.isdir(p)]
            if nopaths:
                dialog_onebutton(
                    self,
                    "Error in Library Settings.\nSecondary library path(s) do not exist:\n" + '\n'.join(nopaths),
                    "Error",
                )
                return

            main_norm = _norm_path(mainlib_path)
            out_norm = _norm_path(output_path)
            for p in new_settings['seclib_path']:
                p_norm = _norm_path(p)
                if p_norm == main_norm:
                    dialog_onebutton(
                        self,
                        "Error in Library Settings.\nSecondary library path(s) cannot be identical to main library path.",
                        "Error",
                    )
                    return
                if p_norm == out_norm:
                    dialog_onebutton(
                        self,
                        "Error in Library Settings.\nSecondary library path(s) cannot be identical to output path.",
                        "Error",
                    )
                    return

        new_settings['close_to_system_tray'] = 1 if self.checkBox_close_to_system_tray.isChecked() else 0
        new_settings['esc_to_system_tray'] = 1 if self.checkBox_esc_to_system_tray.isChecked() else 0
        if hasattr(self, "worksheet_column_checks"):
            worksheet_column_enabled = {
                name: 1 if checkbox.isChecked() else 0
                for name, checkbox in self.worksheet_column_checks.items()
            }
            worksheet_column_enabled["Ref"] = 1
            worksheet_column_enabled["Title"] = 1
            new_settings["worksheet_column_enabled"] = worksheet_column_enabled

        # Developer mode (restart required). Only persist when the checkbox is visible.
        if (
            hasattr(self, "checkBox_developerMode")
            and self.checkBox_developerMode is not None
            and self.checkBox_developerMode.isVisible()
        ):
            new_settings["developerMode"] = 1 if self.checkBox_developerMode.isChecked() else 0

        # Hotkeys (keep existing behaviour)
        self.hotkey_open = ""
        if self.groupBox_open.isChecked():
            for checkbox, textbox in zip(self.checkboxes_open, self.textboxes_open):
                if checkbox.isChecked():
                    text = textbox.toPlainText()
                    if len(text) == 1 and text.isalpha():
                        self.hotkey_open = checkbox.text() + text.lower()

        self.hotkey_inspect = ""
        if self.groupBox_inspect.isChecked():
            for checkbox, textbox in zip(self.checkboxes_inspect, self.textboxes_inspect):
                if checkbox.isChecked():
                    text = textbox.toPlainText()
                    if len(text) == 1 and text.isalpha():
                        self.hotkey_inspect = checkbox.text() + text.lower()

        self.hotkey_pubmed = ""
        if self.groupBox_pubmed.isChecked():
            for checkbox, textbox in zip(self.checkboxes_pubmed, self.textboxes_pubmed):
                if checkbox.isChecked():
                    text = textbox.toPlainText()
                    if len(text) == 1 and text.isalpha():
                        self.hotkey_pubmed = checkbox.text() + text.lower()

        # Validate duplicate trailing alphabet across any enabled hotkeys.
        active_hotkeys = [h for h in (self.hotkey_open, self.hotkey_inspect, self.hotkey_pubmed) if h]
        active_letters = [h[-1] for h in active_hotkeys]
        if len(active_letters) != len(set(active_letters)):
            dialog_onebutton(
                self,
                "Error in Hotkeys Settings.\nThe alphabets for the two hotkeys cannot be the same.",
                "Error",
            )
            return

        new_settings['hotkey_open_value'] = self.hotkey_open
        new_settings['hotkey_inspect_value'] = self.hotkey_inspect
        new_settings['hotkey_pubmed_value'] = self.hotkey_pubmed

        save_settings(new_settings)
        rt.settings = new_settings
        logging.getLogger(__name__).debug("Settings saved")
        dialog_onebutton(self, "Settings saved! Please restart the software.", "Settings saved")
        self.exit_main_window.emit()
        self.close()

    def close_preferences_window(self):
        self.close()

    def closeEvent(self, event):
        # action_in_progress and parent.setEnabled() are owned by the caller
        # (open_preferences) — nothing to release here.
        event.accept()
