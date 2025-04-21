# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

import sys
import os
import platform
from PyQt6.QtWidgets import QMessageBox, QApplication
from PyQt6.QtGui import QPixmap
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QSplashScreen

app = QApplication(sys.argv)

# Get the name of the operating system
os_name = platform.system() #"Windows" or "Darwin"
if os_name == "Windows" or os_name == "Darwin":
    pass
else:
    raise Exception("Unsupported operating system")

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
assets_dir = os.path.join(script_dir, 'assets')
src_dir = os.path.join(script_dir, 'src')
ui_dir = os.path.join(script_dir, 'ui')
data_dir = os.path.join(script_dir, 'data')

# Read the version from the version.txt file
version_file_path = os.path.join(data_dir, 'version.txt')
with open(version_file_path, 'r') as version_file:
    version = version_file.read().strip()

#identify and make appdata directory
if os_name == "Windows":
    appdatadir = os.path.join(os.getenv('APPDATA'), 'pubxel')
elif os_name == "Darwin":
    appdatadir = os.path.expanduser("~/Library/Application Support/pubxel")
os.makedirs(appdatadir, exist_ok=True)


if os_name == "Windows":
    import msvcrt
    lock_file_path = os.path.join(appdatadir, 'my_script.lock')
    def create_lock_file():
        global lock_file
        lock_file = open(lock_file_path, 'w')
        try:
            msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        except IOError:
            QMessageBox.critical(None, "Error", "Pub-Xel is already running.")
            sys.exit(0)
    create_lock_file()
elif os_name == "Darwin":
    import fcntl
    lock_file_path = os.path.join(appdatadir, 'my_script.lock')
    def create_lock_file():
        global lock_file
        lock_file = open(lock_file_path, 'w')
        try:
            fcntl.flock(lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except IOError:
            QMessageBox.critical(None, "Error", "Pub-Xel is already running.")
            sys.exit(0)
    create_lock_file()
else :
    raise Exception("Unsupported operating system")


# Function to create and display the loading screen
def show_loading_screen():
    global splash
    loading_image_path = os.path.join(assets_dir, 'loading.png')
    pixmap = QPixmap(loading_image_path)
    pixmap = pixmap.scaled(500, 500, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
    splash = QSplashScreen(pixmap, Qt.WindowType.WindowStaysOnTopHint)
    splash.setWindowFlag(Qt.WindowType.FramelessWindowHint)
    splash.show()
    splash.activateWindow()
    splash.raise_()

# Show the loading screen
show_loading_screen()
print("Loading screen shown")

# Function to close the loading screen
def close_loading_screen():
    global splash
    if splash is not None:
        splash.close()
        splash = None
app.setQuitOnLastWindowClosed(False)

#do stuff
import time
from pynput import keyboard
import threading
import pyperclip
import webbrowser
import subprocess
import shlex
from mainfunctions import *
import datetime
import shutil
import copy
import xlwings as xw
import concurrent.futures

from PyQt6.QtWidgets import QScrollArea,QGroupBox,QPlainTextEdit,QFileDialog, QDialog, QFrame, QSystemTrayIcon,QSizePolicy,QGridLayout,QSpacerItem,QPushButton,QLabel,QCheckBox, QWidget, QMainWindow,QHBoxLayout, QVBoxLayout, QMenu
from PyQt6.QtGui import QFontMetrics,QTextDocument,QKeySequence,QShortcut,QAction, QIcon, QTextCursor
from PyQt6 import QtCore
from PyQt6.QtCore import QPropertyAnimation, QEventLoop,QThread, QUrl, QEvent, QObject, pyqtSignal, QTimer
from PyQt6 import uic

# Build the full path to the file
main_path = os.path.join(ui_dir, 'main.ui')
inspect_path = os.path.join(ui_dir, 'inspect.ui')
about_path = os.path.join(ui_dir, 'about.ui')
preferences_path = os.path.join(ui_dir, 'preferences.ui')
icon_path = os.path.join(assets_dir, 'logo64.ico')
questionmark_icon_path = os.path.join(assets_dir, 'questionmark.png')
loading_image_path = os.path.join(assets_dir, 'loading.png')
settingsdefault_path = os.path.join(data_dir, 'settingsdefault.json')
pubsheetdefault_path = os.path.join(data_dir, 'pubsheet.xlsx')
pubsheetinitialdefault_path = os.path.join(data_dir, 'pubsheetinitial.xlsx')
#settings
settings_path = os.path.join(appdatadir,"settings.json")
#worksheet
pubsheet_path = os.path.join(appdatadir, 'pubsheet.xlsx')
pubsheetinitial_path = os.path.join(appdatadir, 'pubsheetinitial.xlsx')


#Copy files to Application Folder.
files_to_copy = {
    settingsdefault_path: settings_path,  # Copy to settings.json
    pubsheetdefault_path: pubsheet_path,  # Copy to pubsheet.xlsx
    pubsheetinitialdefault_path: pubsheetinitial_path  # Copy to pubsheet.xlsx
}
for source_path, dest_path in files_to_copy.items():
    if dest_path == pubsheet_path or dest_path == pubsheetinitial_path or not os.path.exists(dest_path):
        shutil.copy(source_path, dest_path)
        print(f"Copied {source_path} as {dest_path}")

#update the settings
def update_settings(settingsdefault_path, settings_path):
    with open(settingsdefault_path, 'r') as default_file:
        default_settings = json.load(default_file)
    with open(settings_path, 'r') as settings_file:
        settings = json.load(settings_file)
    updated = False
    for key, value in default_settings.items():
        if key not in settings:
            settings[key] = value
            updated = True
            print("Updated settings:", key)
    if updated:
        with open(settings_path, 'w') as settings_file:
            json.dump(settings, settings_file, indent=4)
    else:
        print("No settings updates needed.")

# update settings and then load settings
update_settings(settingsdefault_path, settings_path)
settings = load_settings()

#desktop
if os_name == "Windows":
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
elif os_name == "Darwin":
    desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
#documents_path
if os_name == "Windows":
    documents_path = os.path.join(os.environ['USERPROFILE'], 'Documents')
elif os_name == "Darwin":  # MacOS
    documents_path = os.path.join(os.environ['HOME'], 'Documents')

mainlibdirdefault = os.path.join(documents_path, "Pub-Xel Library")
outdirdefault = documents_path

#initialize library and output path
#review and update settings
if settings.get('mainlib_path', 0) == "":
    settings['mainlib_path'] = mainlibdirdefault
    os.makedirs(mainlibdirdefault, exist_ok=True)
    save_settings(settings)
mainlibdir = settings.get('mainlib_path', 0)

if settings.get('output_path', 0) == "":
    settings['output_path'] = documents_path
    save_settings(settings)
outdir = settings.get('output_path', 0)

if settings.get('seclib_enable',0):
    seclibdir = settings.get('seclib_path', [])
else:
    seclibdir = []

#disable hotkeys for now in Mac
if os_name == "Darwin":
    settings = save_settings_key(settings,"hotkey_inspect_value","")
    settings = save_settings_key(settings,"hotkey_open_value","")

#other settings
system_tray_notice_shown = settings.get('system_tray_notice_shown', 0)
developerMode = settings.get('developerMode',0)

# Flag to indicate whether an action is in progress
action_in_progress = False

# _get_sep, dirname: functions to get the path separator and directory component, for both windows and mac
def _get_sep(p):
    """Returns the appropriate path separator for the given path"""
    if isinstance(p, bytes):
        return b'\\' if os.name == 'nt' else b'/'
    else:
        return '\\' if os.name == 'nt' else '/'
def dirname(p):
    """Returns the directory component of a pathname"""
    p = os.fspath(p)
    sep = _get_sep(p)
    i = p.rfind(sep) + 1
    head = p[:i]
    if head and head != sep*len(head):
        head = head.rstrip(sep)
    return head

def open_directory(dir):
    os_name = platform.system()
    if os_name == 'Windows':  # Windows
        os.startfile(dir)
        # subprocess.Popen(f'start "" "{dir}"', start_new_session=True, shell=True,creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)
    elif os_name == 'Darwin':  # macOS
        print(f'open {shlex.quote(dir)}')
        subprocess.call(f'open {shlex.quote(dir)}', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)

def try_open_directory(dir):
    try:
        open_directory(dir)
    except Exception as e:
        print(f"Cannot open: {e}")

def stop_listeners():
    global listeners
    for listener in listeners:
        listener.stop()
    listeners = []

def dialog_onebutton(parent, message,title="Confirmation"):
    msg_box = QMessageBox(parent)
    msg_box.setWindowTitle(title)
    msg_box.setText(message)
    msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
    msg_box.setWindowModality(Qt.WindowModality.ApplicationModal)
    msg_box.exec()
    msg_box.setFocus(QtCore.Qt.FocusReason.PopupFocusReason)

class window_preferences(QDialog):
    exit_main_window = pyqtSignal()
    def __init__(self, parent):
        super().__init__()
        uic.loadUi(preferences_path, self)
        global action_in_progress
        action_in_progress = True
        print("action_in_progress set to True")
        parent.setEnabled(False)  # Disable the main window
        self.pseudo_parent = parent

        self.setWindowTitle('Preferences')
        global settings
        self.settings = settings
        self.tabWidget.setCurrentIndex(0)
        self.tab_hot = self.findChild(QWidget, 'tab_hot')
        if self.tab_hot and os_name == "Darwin":
            print("Hotkeys tab disabled in MacOS")
            index_to_hide = self.tabWidget.indexOf(self.tab_hot)
            self.tabWidget.removeTab(index_to_hide)
        
        # library tab
        self.plainTextEdit_mainlib = self.findChild(QPlainTextEdit, 'plainTextEdit_mainlib')
        self.plainTextEdit_mainlib.setPlainText(self.settings.get('mainlib_path', 0))
        self.button_selectmainlib = self.findChild(QPushButton, 'button_selectmainlib')
        self.button_selectmainlib.clicked.connect(self.setmainlib)

        self.plainTextEdit_output = self.findChild(QPlainTextEdit, 'plainTextEdit_output')
        self.plainTextEdit_output.setPlainText(self.settings.get('output_path', 0))
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
        hotkey_open_value = settings.get('hotkey_open_value', 0)
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
        hotkey_inspect_value = settings.get('hotkey_inspect_value', 0)
        if not hotkey_inspect_value == "":
            self.groupBox_inspect.setChecked(True)
            for i, hotkey_string in enumerate(self.hotkey_strings):
                if hotkey_inspect_value[:-1] == hotkey_string:
                    self.checkboxes_inspect[i].setChecked(True)
                    self.textboxes_inspect[i].setEnabled(True)
                    self.textboxes_inspect[i].setPlainText(hotkey_inspect_value[len(hotkey_string):])
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
        # if settings.get('launch_at_startup',0):
        #     self.checkBox_launch_at_startup.setChecked(True)
        # # self.settings['launch_at_startup'] = 1 if self.checkBox_launch_at_startup.isChecked() else 0
        self.checkBox_close_to_system_tray = self.findChild(QCheckBox, 'checkBox_close_to_system_tray')
        if settings.get('close_to_system_tray',0):
            self.checkBox_close_to_system_tray.setChecked(True)
        self.checkBox_esc_to_system_tray = self.findChild(QCheckBox, 'checkBox_esc_to_system_tray')
        if settings.get('esc_to_system_tray',0):
            self.checkBox_esc_to_system_tray.setChecked(True)

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
            if os_name == 'Windows':
                folder_path = folder_path.replace("/", "\\")
            print(folder_path)
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
            print(f"Checkbox index: {index}")
            self.textboxes_open[index].setEnabled(sender.isChecked())
            for textbox in self.textboxes_open:
                if textbox != self.textboxes_open[index]:
                    textbox.setEnabled(False)
        elif isinstance(sender, QPlainTextEdit):
            self.textboxcurrent = True
            print(self.textboxcurrent)
            index = self.textboxes_open.index(sender)
            text = sender.toPlainText()
            print(f"Textbox text: {text}")
            if len(text) > 1:
                sender.setPlainText(text[-1])
            text = sender.toPlainText()
            if not text.isalpha():
                sender.setPlainText("")
            sender.setPlainText(sender.toPlainText().lower())
            self.textboxcurrent = False
            print(self.textboxcurrent)

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
            print(f"Checkbox index: {index}")
            self.textboxes_inspect[index].setEnabled(sender.isChecked())
            for textbox in self.textboxes_inspect:
                if textbox != self.textboxes_inspect[index]:
                    textbox.setEnabled(False)
            
        elif isinstance(sender, QPlainTextEdit):
            self.textboxcurrent = True
            print(self.textboxcurrent)
            index = self.textboxes_inspect.index(sender)
            text = sender.toPlainText()
            print(f"Textbox text: {text}")
            if len(text) > 1:
                sender.setPlainText(text[-1])
            text = sender.toPlainText()
            if not text.isalpha():
                sender.setPlainText("")
            sender.setPlainText(sender.toPlainText().lower())
            self.textboxcurrent = False
            print(self.textboxcurrent)

    def libdefault(self):
        self.plainTextEdit_mainlib.setPlainText(mainlibdirdefault)
        self.plainTextEdit_output.setPlainText(outdirdefault)
        self.groupBox_seclib.setChecked(False)

    def hotdefault(self):
        self.groupBox_open.setChecked(True)
        self.groupBox_inspect.setChecked(True)
        for checkbox in self.checkboxes_open:
            checkbox.setChecked(False)
        for checkbox in self.checkboxes_inspect:
            checkbox.setChecked(False)
        for textbox in self.textboxes_open:
            textbox.setPlainText("")
            textbox.setEnabled(False)
        for textbox in self.textboxes_inspect:
            textbox.setPlainText("")
            textbox.setEnabled(False)

        index = self.hotkey_strings.index("<Alt>+<Shift>+")
        self.checkboxes_open[index].setChecked(True)
        self.textboxes_open[index].setEnabled(True)
        self.textboxes_open[index].setPlainText("k")
        self.checkboxes_inspect[index].setChecked(True)
        self.textboxes_inspect[index].setEnabled(True)
        self.textboxes_inspect[index].setPlainText("j")

    def saveexit(self):
        backup_settings = copy.deepcopy(self.settings)
        self.settings['mainlib_path'] = self.plainTextEdit_mainlib.toPlainText()
        if not os.path.exists(self.settings['mainlib_path']):
            self.settings = copy.deepcopy(backup_settings)
            dialog_onebutton(self,"Error in Library Settings.\nMain library path does not exist.","Error")
            return
        self.settings['output_path'] = self.plainTextEdit_output.toPlainText()
        if not os.path.exists(self.settings['output_path']):
            self.settings = copy.deepcopy(backup_settings)
            dialog_onebutton(self,"Error in Library Settings.\nOutput path does not exist.","Error")
            return
        if dirname(self.settings['output_path']) == dirname(self.settings['mainlib_path']):
            self.settings = copy.deepcopy(backup_settings)
            dialog_onebutton(self,"Error in Library Settings.\nOutput path cannot be identical to main library path.","Error")
            return
        self.settings['seclib_path'] = sorted(set(line for line in self.plainTextEdit_seclib.toPlainText().split('\n') if line.strip()))
        self.settings['seclib_enable'] = 1 if self.groupBox_seclib.isChecked() else 0
        # if seclib_enable is 1 and any of seclib_path does not exist, show error, list the error directories, and return
        if self.settings['seclib_enable'] and self.settings['seclib_path']:
            nopaths = []
            for path in self.settings['seclib_path']:
                if not os.path.exists(path):
                    nopaths.append(path)
            if nopaths:
                self.settings = copy.deepcopy(backup_settings)
                dialog_onebutton(self,f"Error in Library Settings.\nSecondary library path(s) do not exist:\n" + '\n'.join(nopaths),"Error")
                return
        if self.settings['seclib_enable'] and self.settings['seclib_path']:
            # if any of seclib_path is identical to mainlib_path in terms of directory, show error, and return
            for path in self.settings['seclib_path']:
                if dirname(path) == dirname(self.settings['mainlib_path']):
                    self.settings = copy.deepcopy(backup_settings)
                    dialog_onebutton(self,"Error in Library Settings.\nSecondary library path(s) cannot be identical to main library path.","Error")
                    return
                if dirname(path) == dirname(self.settings['output_path']):
                    self.settings = copy.deepcopy(backup_settings)
                    dialog_onebutton(self,"Error in Library Settings.\nSecondary library path(s) cannot be identical to output path.","Error")
                    return
        self.settings['close_to_system_tray'] = 1 if self.checkBox_close_to_system_tray.isChecked() else 0
        self.settings['esc_to_system_tray'] = 1 if self.checkBox_esc_to_system_tray.isChecked() else 0
        
        self.hotkey_open = ""
        if not self.groupBox_open.isChecked():
            pass
        else:
            for checkbox, textbox in zip(self.checkboxes_open, self.textboxes_open):
                if checkbox.isChecked():
                    text = textbox.toPlainText()
                    if len(text) == 1 and text.isalpha():
                        self.hotkey_open = checkbox.text() + text.lower()
        self.hotkey_inspect = ""
        if not self.groupBox_inspect.isChecked():
            pass
        else:
            for checkbox, textbox in zip(self.checkboxes_inspect, self.textboxes_inspect):
                if checkbox.isChecked():
                    text = textbox.toPlainText()
                    if len(text) == 1 and text.isalpha():
                        self.hotkey_inspect = checkbox.text() + text.lower()
        if not self.hotkey_open == "" and not self.hotkey_inspect == "":
            if self.hotkey_open[-1] == self.hotkey_inspect[-1]:
                self.settings = copy.deepcopy(backup_settings)
                dialog_onebutton(self,"Error in Hotkeys Settings.\nThe alphabets for the two hotkeys cannot be the same.","Error")
                return

        self.settings['hotkey_open_value'] = self.hotkey_open
        self.settings['hotkey_inspect_value'] = self.hotkey_inspect
        save_settings(self.settings)
        print("Settings saved")
        self.exit_main_window.emit()
        self.close()

    def close_preferences_window(self):
        self.close()

    def closeEvent(self, event):
        global action_in_progress
        action_in_progress = False
        print("action_in_progress False")
        self.pseudo_parent.setEnabled(True)  # Re-enable the main window
        event.accept()

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
        super().__init__()
        uic.loadUi(about_path, self)
        global action_in_progress
        action_in_progress = True
        print("action_in_progress True")
        self.hide()
        parent.setEnabled(False)  # Disable the main window
        self.pseudo_parent = parent

        self.setWindowTitle('About')

        global version
        link = "https://pubxel.org/"
        self.findChild(QLabel, 'label_version').setText(f"Version: {version}")
        self.label_home = self.findChild(QLabel, 'label_home')
        self.label_home.setText(f'Home: <a href="{link}">{link}</a> ')
        self.label_home.setOpenExternalLinks(True)
        #question mark icons
        pixmap = QPixmap(loading_image_path)
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
    def __init__(self, parent, data=None):
        super().__init__()
        uic.loadUi(inspect_path, self)
        global action_in_progress
        action_in_progress = True
        print("action_in_progress True")
        self.setWindowIcon(QIcon(icon_path))
        self.hide()
        parent.setEnabled(False)  # Disable the main window
        self.pseudo_parent = parent

        self.setWindowTitle('Inspect')
        self.data = data  # Store the passed data
        self.pubmeddata = None

        # Get the selected cell value from clipboard
        selected_cell_value = pyperclip.paste()
        selected_cell_value = string_to_list(selected_cell_value)
        output = process_ids(selected_cell_value,mainlibdir,seclibdir)
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

        self.all_ids = valid_ids + invalid_ids
        self.all_files = all_m_files + all_s_files
        
        self.label_title = self.findChild(QLabel, 'label_title')
        self.scrollArea_idlist = self.findChild(QScrollArea, 'scrollArea_idlist')
        self.scrollLayout_idlist = self.findChild(QWidget, 'scrollLayout_idlist')
        self.label_idlist = self.findChild(QLabel, 'label_idlist')
        if len(valid_ids) ==1 and len(pubmed_ids) == 1 and len(invalid_ids) ==0:           
            inspect_window_title=f"PMID {pubmed_ids[0]}.\n\n"
            self.onepubmedarticle = 1
            self.onepubmedarticle_id = pubmed_ids[0]
            self.scrollArea_idlist.deleteLater()
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
                    self.labels_all.append(idlabel)

                    h_layout = QHBoxLayout()
                    h_layout.setContentsMargins(0, 0, 0, 0)
                    h_layout.addWidget(idlabel)
                    h_layout.addStretch()
                    container_widget = QWidget()
                    container_widget.setLayout(h_layout)
                    self.scrollLayout_idlist.layout().addWidget(container_widget)
        
        self.label_title.setText(inspect_window_title)
        # self.label_title.installEventFilter(self)

        gridLayout_summary = self.findChild(QGridLayout, 'gridLayout_summary')
        if pubmed_ids:
            label_pubmed,button_copypubmed,button_searchpubmed = self.create_elements(gridLayout_summary,0,'label_pubmed','button_copypubmed', 'button_searchpubmed')
            label_pubmed.setText(f"PubMed ID(s): {len(pubmed_ids)}")
            button_copypubmed.clicked.connect(lambda: self.show_copy_id_and_show_popup(pubmed_ids))
            button_searchpubmed.clicked.connect(lambda: self.search_pubmed(pubmed_ids))
            button_searchpubmed.setText('View &PubMed')
            shortcut_p = QShortcut(QKeySequence('p'), self)
            shortcut_p.activated.connect(lambda: self.search_pubmed(pubmed_ids))
        if pubmed_ids_without_m_files:
            label_pubmedna,button_copypubmedna,button_searchpubmedna = self.create_elements(gridLayout_summary,1,'label_pubmedna','button_copypubmedna', 'button_searchpubmedna')
            label_pubmedna.setText(f"    PubMed ID(s) without main files: {len(pubmed_ids_without_m_files)}")
            button_copypubmedna.clicked.connect(lambda: self.show_copy_id_and_show_popup(pubmed_ids_without_m_files))
            button_searchpubmedna.clicked.connect(lambda: self.search_pubmed(pubmed_ids_without_m_files))
        if non_pubmed_valid_ids:
            label_nonpub,button_copynonpub,temp = self.create_elements(gridLayout_summary,2,'label_nonpub','button_copynonpub')
            label_nonpub.setText(f"Non-PubMed ID(s): {len(non_pubmed_valid_ids)}")
            button_copynonpub.clicked.connect(lambda: self.show_copy_id_and_show_popup(non_pubmed_valid_ids))
        if nonpubmed_ids_without_m_files:
            label_nonpubna,button_copynonpubna,temp = self.create_elements(gridLayout_summary,3,'label_nonpubna','button_copynonpubna')
            label_nonpubna.setText(f"    Non-PubMed ID(s) without main files: {len(nonpubmed_ids_without_m_files)}")
            button_copynonpubna.clicked.connect(lambda: self.show_copy_id_and_show_popup(nonpubmed_ids_without_m_files))
        if invalid_ids:
            label_na,button_copyna,temp = self.create_elements(gridLayout_summary,4,'label_na','button_copyna')
            label_na.setText(f"Invalid ID(s): {len(invalid_ids)}")
            button_copyna.clicked.connect(lambda: self.show_copy_id_and_show_popup(invalid_ids))
        
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

        # inspect_window.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        # # After a short delay, remove the always on top attribute
        # QTimer.singleShot(1000, lambda: (inspect_window.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, False), inspect_window.show()))
    
    def show_copy_id_and_show_popup(self,pubmed_ids):
        copy_list(pubmed_ids)
        self.hide_popup_message()
        self.popup_message = PopupMessageFade(self)  # Create a new instance each time
        self.popup_message.show_popup("ID(s) copied to clipboard.")
        

    def load_pubmed_data(self, pubmed_ids):
        if pubmed_ids:
            try:
                data = obtain_pubmed_data(pubmed_ids)
                if isinstance(data, dict):
                    self.pubmeddata = data
                print("Data loaded: "+str(len(data))+" items")
            except Exception as e:
                print(f"Failed to load data: {e}")
        if self.pubmeddata and self.onepubmedarticle:
            try:
                inspect_window_title = f"PMID {pubmed_ids[0]} : " + self.pubmeddata[pubmed_ids[0]].get("cite","")
                self.label_title.setText(inspect_window_title)
            except Exception as e:
                print(f"Failed to set label_title: {e}")
        if self.pubmeddata and self.checkboxes_main:
            for checkbox in self.checkboxes_main:
                article_id = checkbox.text()  
                if article_id and article_id[0].isdigit():
                    match = re.match(r"^\d+", article_id) # remove characters after numerics
                    if match:  
                        pubmed_id = match.group(0)
                        if pubmed_id in self.pubmeddata:
                            authoryear = self.pubmeddata[pubmed_id].get("authoryear","")
                            if authoryear:
                                checkbox.setText(article_id+" : "+authoryear)

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
                pubmed_id = match.group(0)
                PMID = self.pubmeddata.get(pubmed_id, {}).get("PMID", "")
                if PMID:
                    if isinstance(source, QLabel):
                        popuptext = f"PMID {article_id} : " + self.pubmeddata[pubmed_id].get("cite","")
                    if isinstance(source, QCheckBox):
                        cite_maincheckbox = self.pubmeddata[pubmed_id].get("cite_maincheckbox","")
                        if cite_maincheckbox:
                            popuptext = cite_maincheckbox
                        else: popuptext = article_id
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
            self.popup_toolip.setStyleSheet("QLabel { padding: 1px; }")
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
        global action_in_progress
        action_in_progress = False
        print("action_in_progress False")
        self.hide_popup_message()
        self.hide_popup_tooltip()
        self.pseudo_parent.setEnabled(True) # Re-enable the main window
        event.accept()

    def search_pubmed(self,lst):
        webbrowser.open("https://pubmed.ncbi.nlm.nih.gov/?term="+"%5Buid%5D+OR+".join(lst)+"%5Buid%5D&sort=date")
        self.close_inspect_window()

    def callback(self,result): #open files and quit new_window
        global action_in_progress
        if result == []:
            return
        # Handle the result in your higher-level code
        filepathList = files_name_to_path(result,mainlibdir,seclibdir)

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
            folder_name = f"{today}"
            folder_path = os.path.join(base_dir, folder_name)
            counter = 2

            while os.path.exists(folder_path):
                folder_name = f"{today} ({counter})"
                folder_path = os.path.join(base_dir, folder_name)
                counter += 1

            os.makedirs(folder_path)
            return folder_name, folder_path
        
        newfolder, newdir = create_dated_folder(outdir)
        print(f"New folder created at: {newdir}")

        paths = files_name_to_path(files,mainlibdir,seclibdir)
        
        # Copy wanted files
        copied_files = 0
        for path in paths:
            if os.path.isfile(path):
                shutil.copy(path, newdir)
                copied_files += 1

        def open_newdir():
            try_open_directory(os.path.realpath(newdir))

        def show_message_box():
            msg_box = QMessageBox()
            msg_box.setWindowTitle("Completion")
            msg_box.setText(f"A total of {copied_files} file(s) were successfully copied in folder: {newfolder}")
            open_folder_button = msg_box.addButton("Open Folder", QMessageBox.ButtonRole.AcceptRole)
            open_folder_button.clicked.connect(lambda: open_newdir())
            msg_box.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
            msg_box.exec()
        show_message_box()
        return True

    def create_elements(self, layout: QGridLayout, row: int, label: str, button1: str, button2: str=None):
        # Create a QLabel and add it to the layout
        label_widget = QLabel(label)
        layout.addWidget(label_widget, row, 0)
        # Create a spacer item and add it to the layout
        spacer_item = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        layout.addItem(spacer_item, row, 1)
        button2_widget = None
        # Create the first button and add it to the layout
        button1_widget = QPushButton(button1)
        layout.addWidget(button1_widget, row, 3)
        button1_widget.setText("Copy ID(s)")
        # If a second button is provided, create it and add it to the layout
        if button2:
            button2_widget = QPushButton(button2)
            layout.addWidget(button2_widget, row, 2)
            button2_widget.setText("View PubMed")
        return label_widget,button1_widget, button2_widget

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


class listenerWorker(QObject):
    open_inspect_signal = pyqtSignal()
    open_file_signal = pyqtSignal()
    def __init__(self):
        super().__init__()
    def run_inspect(self):
        self.open_inspect_signal.emit()
    def run_file(self):
        self.open_file_signal.emit()

listeners = []

class clipboardWorker(QThread):
    clipboard_updated = pyqtSignal(str)
    def run(self):
        text = ""
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future = executor.submit(self.process_clipboard)
            try:
                text = future.result(timeout=5)  # 5 seconds timeout
                self.clipboard_updated.emit(text)
            except concurrent.futures.TimeoutError:
                print("Processing clipboard data timed out")

    def process_clipboard(self):
        try:
            clipboardstring = ""
            clipboardstring = pyperclip.paste()
            if clipboardstring:
                return list_to_string(string_to_list(clipboardstring))
        except Exception as e:
            print(f"Error accessing clipboard: {e}")
        return ""

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
                text = f"{cellcount} Cell{'s' if cellcount > 1 else ''} in {os.path.basename(wb.fullname)}."
            else: 
                text = ""
            return text
        except Exception as e:
            print(f"Error accessing clipboard: {e}")
        return ""


def check_shortcut(worker):

    hotkey_inspect_value = settings.get("hotkey_inspect_value",0)
    hotkey_open_value = settings.get("hotkey_open_value",0)

    global action_in_progress, listeners

    def on_activate_inspect():
        global action_in_progress
        if not action_in_progress:
            print("on_activate_inspect activated")
            worker.run_inspect()
        else: 
            print("on_activate_inspect but action in progress")

    def on_activate_open():
        global action_in_progress
        if not action_in_progress:
            print("on_activate_open activated")
            worker.run_file()
        else:  
            print("on_activate_open but action in progress")

    if hotkey_inspect_value == "" and hotkey_open_value == "":
        return
    if not hotkey_inspect_value == "":
        hotkey_j = keyboard.HotKey(keyboard.HotKey.parse(hotkey_inspect_value), on_activate_inspect)
    if not hotkey_open_value == "":
        hotkey_k = keyboard.HotKey(keyboard.HotKey.parse(hotkey_open_value), on_activate_open)
    
    def for_canonical(f):
        return lambda k: f(listeners[0].canonical(k))

    def on_press(key):
        if not hotkey_inspect_value == "":
            hotkey_j.press(key)
        if not hotkey_open_value == "":
            hotkey_k.press(key)
        return

    def on_release(key):
        if not hotkey_inspect_value == "":
            hotkey_j.release(key)
        if not hotkey_open_value == "":
            hotkey_k.release(key)
        return

    def start_listener():
        try:
            listener = keyboard.Listener(
                on_press=for_canonical(on_press),
                on_release=for_canonical(on_release))
            listeners.append(listener)
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


class SystemTrayIcon(QSystemTrayIcon):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setIcon(QIcon(icon_path))
        self.setToolTip("Pub-Xel")
        if os_name == "Windows":
            self.activated.connect(self.on_activated)

        self.menu = QMenu(parent)
        self.restore_action = QAction("Restore", self)
        self.exit_action = QAction("Exit", self)
        self.restore_action.triggered.connect(self.on_restore)
        self.exit_action.triggered.connect(self.on_exit)
        self.menu.addAction(self.restore_action)
        self.menu.addAction(self.exit_action)
        self.setContextMenu(self.menu)

    def on_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.on_restore()

    def on_restore(self):
        if self.parent().isMinimized():
            self.parent().showNormal()
        self.parent().show()
        self.parent().activateWindow()

    def on_exit(self):
        self.parent().close_application()

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

class RunningFunctionDialog(QDialog):
    def __init__(self, parent=None, message=None):
        global action_in_progress
        action_in_progress = True
        print("action_in_progress True")
        parent.setEnabled(False)  # Disable the main window
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
        global action_in_progress
        action_in_progress = False
        print("action_in_progress False")
        self.pseudo_parent.setEnabled(True)  # Re-enable the main window
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

class main_window(QMainWindow):

    def __init__(self):
        global settings
        super().__init__()
        uic.loadUi(main_path, self)
        self.setWindowTitle('Pub-Xel')
        self.setWindowIcon(QIcon(icon_path))
        self.exitstatus = False
        self.is_closing = False

        if os_name == "Darwin":
            appdock = QApplication.instance() # for dock click event
            appdock.applicationStateChanged.connect(self.on_application_state_changed) # for dock click event

        if developerMode:
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
        
        self.findChild(QLabel, 'label_inspect').setText("  "+settings.get('hotkey_inspect_value', 0).replace('<', '').replace('>', ''))
        self.findChild(QLabel, 'label_open').setText("  "+settings.get('hotkey_open_value', 0).replace('<', '').replace('>', ''))

        #question mark icons
        pixmap = QPixmap(questionmark_icon_path)
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

        self.clipboardWorker = clipboardWorker()
        self.clipboardWorker.clipboard_updated.connect(self.update_clipboard_text)
        self.clipboardWorker.finished.connect(self.clipboardWorker_finished)
        self.clipboardWorker_running = False

        self.excelWorker = excelWorker()
        self.excelWorker.excel_updated.connect(self.update_excel_text)
        self.excelWorker.finished.connect(self.excelWorker_finished)
        self.excelWorker_running = False

        self.findChild(QAction, 'actionExit').triggered.connect(self.close_application)
        self.findChild(QAction, 'actionMinimize').triggered.connect(self.minimize_to_tray)
        if settings.get('esc_to_system_tray',0):
            shortcut_Esc = QShortcut(QKeySequence('Esc'), self)
            shortcut_Esc.activated.connect(self.minimize_to_tray) 
        self.findChild(QAction, 'actionOpen_Library_Folder').triggered.connect(lambda: try_open_directory(mainlibdir))
        self.findChild(QAction, 'actionOpen_Output_Folder').triggered.connect(lambda: try_open_directory(outdir))
        self.findChild(QAction, 'actionNew_Excel_Template').triggered.connect(self.save_pubsheet)
        self.findChild(QAction, 'actionPreferences').triggered.connect(self.open_preferences)
        self.findChild(QAction, 'actionAbout').triggered.connect(self.open_about_window)

        self.setStyleSheet("""QGroupBox#groupBoxExcel {font-size: 14px;}
                           QGroupBox#groupBoxClipboard {font-size: 14px;}""")


        self.tray_icon = SystemTrayIcon(self)
        self.tray_icon.show()

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
        if not self.clipboardWorker_running:
            # print("clipboardWorker called, starting now")
            self.clipboardWorker_running = True
            self.clipboardWorker.start()
        else:
            print("clipboardWorker called but already running")

    def clipboardWorker_finished(self):
        # print("clipboardWorker_finished")
        self.clipboardWorker_running = False
        
    def update_clipboard_text(self, text):
        self.plainTextEdit_clipboard_current.setPlainText(text)

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
        if os_name == "Darwin":
            if state == Qt.ApplicationState.ApplicationActive:
                self.show()
                self.activateWindow()
                print("Application activated from Dock")
        else:
            return

    def open_about_window(self):
        global action_in_progress
        if action_in_progress:
            return
        action_in_progress=True
        print("action_in_progress True")
        self.setEnabled(False)  # Disable the main window
        dialog = window_about(self)
        dialog.exec()  # This will block until the dialog is closed
        action_in_progress = False
        print("action_in_progress False")
        self.setEnabled(True)  # Re-enable the main window

    def open_preferences(self):
        global action_in_progress
        if action_in_progress:
            return
        action_in_progress=True
        print("action_in_progress True")
        self.setEnabled(False)  # Disable the main window
        dialog = window_preferences(self)
        dialog.exit_main_window.connect(self.handle_exit_main_window)  # Connect the signal before exec()
        dialog.exec()  # This will block until the dialog is closed
        if self.exitstatus:
            self.close_application()
        action_in_progress = False
        print("action_in_progress False")
        self.setEnabled(True)  # Re-enable the main window

    def handle_exit_main_window(self):
        self.exitstatus = True

    def save_pubsheet(self):
        global action_in_progress
        if action_in_progress:
            return
        action_in_progress = True
        print("action_in_progress True")
        self.setEnabled(False)  # Disable the main window

        file_dialog = QFileDialog(self, "Save Worksheet")
        file_dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        file_dialog.setDefaultSuffix("xlsx")
        file_dialog.setNameFilters(["Excel Files (*.xlsx)", "All Files (*)"])

        try:
            # Set the initial directory to documents_path
            documents_path = os.path.expanduser("~/Documents")
            file_dialog.setDirectory(documents_path)
        except Exception as e:
            print(f"Failed to set initial directory: {e}")

        file_dialog.selectFile("Pub-Xel Worksheet.xlsx")

        if file_dialog.exec() == QFileDialog.DialogCode.Accepted:
            file_path = file_dialog.selectedFiles()[0]

            try:
                if settings["worksheet_count"] > 0:
                    shutil.copyfile(pubsheet_path, file_path)
                else:
                    shutil.copyfile(pubsheetinitial_path, file_path)
                self.save_success_dialog(file_path)

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file: {str(e)}")

        action_in_progress = False
        print("action_in_progress False")
        self.setEnabled(True)  # Re-enable the main window

    def crash(self):
        raise Exception("Crash")
    
    def save_success_dialog(self, file_path):
        global settings
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.setText("Worksheet saved.")
        msg_box.setWindowTitle("Success")

        open_button = QPushButton("Open Worksheet")
        msg_box.addButton(open_button, QMessageBox.ButtonRole.ActionRole)

        if settings["worksheet_count"] > 0:
            open_folder_button = QPushButton("Open Folder")
            msg_box.addButton(open_folder_button, QMessageBox.ButtonRole.ActionRole)
            msg_box.addButton(QMessageBox.StandardButton.Ok)
        
        msg_box.exec()

        if msg_box.clickedButton() == open_button:
            try_open_directory(file_path)
        
        if settings["worksheet_count"] > 0:
            if msg_box.clickedButton() == open_folder_button:
                folder_path = os.path.dirname(file_path)
                try_open_directory(folder_path)

        settings = save_settings_key(settings,"worksheet_count",settings['worksheet_count']+1)

    def run_check_file_exist2(self):
        global action_in_progress
        if action_in_progress:
            return
        self.setEnabled(False)
        self.dialog = RunningFunctionDialog(parent=self,message="Checking file existence...\nPlease do not interact with Excel.")
        def proceed():
            print("0.5 seconds have passed!")  
        timer = QTimer()
        timer.setSingleShot(True)
        timer.timeout.connect(proceed)
        timer.start(500)
        message = self.check_file_exist2()
        if message:
            dialog_onebutton(self.dialog,str(message),"Confirmation")
        self.dialog.close()

    def run_input_pubmed_data2(self):
        global action_in_progress
        if action_in_progress:
            return
        action_in_progress = True
        print("action_in_progress True")
        self.setEnabled(False) 
        self.dialog = RunningFunctionDialog(parent=self,message="Importing PubMed data...\nPlease do not interact with Excel.")
        def proceed():
            print("0.5 seconds have passed!")  
        timer = QTimer()
        timer.setSingleShot(True)
        timer.timeout.connect(proceed)
        timer.start(500)
        message = self.input_pubmed_data2()
        if message:
            dialog_onebutton(self.dialog,str(message),"Confirmation")
        self.dialog.close()

    def check_file_exist2(self):
        try: 
            message = check_file_exist(mainlibdir,seclibdir)
        except Exception as e:
            return e
        return message

    def input_pubmed_data2(self):
        try: 
            check_file_exist(mainlibdir,seclibdir)
        except Exception as e:
            return e
        try:
            message = input_pubmed_data()
        except Exception as e:
            return e
        return message

    def action_in_progress_switch(self):
        global action_in_progress
        action_in_progress = not action_in_progress
        print(action_in_progress)

    def mousePressEvent(self, event):
        if developerMode:
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

    def minimize_to_tray(self):# dont change self.is_closing and system_tray_notice_shown orders
        global system_tray_notice_shown
        global settings
        self.hide()
        if self.is_closing:
            return
        if system_tray_notice_shown:
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
            system_tray_notice_shown += 1
            settings = save_settings_key(settings,'system_tray_notice_shown',1)

    if settings.get('close_to_system_tray', 0):
        def closeEvent(self, event):
            event.ignore()
            self.minimize_to_tray()
    else:
        def closeEvent(self, event):
            print("closeEvent")
            QApplication.quit()
            event.accept()

    def close_application(self):
        self.is_closing = True
        print("close_application")
        QApplication.quit()
                                          
    def sayHello(self,event=None):
        print("Hello")

    def main_openfile(self):
        global action_in_progress
        if action_in_progress:
            return
        print("opening files from main window...")
        action_in_progress = True  # Set the action in progress
        print("action_in_progress True")
        self.setEnabled(False)  # Disable the main window
        clipboardstring = pyperclip.paste()
        idlist = string_to_list(clipboardstring)
        output = process_ids(idlist,mainlibdir,seclibdir)
        self.openfile_from_list(output[10])
        action_in_progress = False
        print("action_in_progress False")
        self.setEnabled(True)  # Disable the main window

    def openfile_from_list(self, filelist):
        if filelist == []:
            dialog_onebutton(self,"No files to open.","Error")
            return
        filepathList = files_name_to_path(filelist,mainlibdir,seclibdir)
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

    def open_inspect_window(self):
        global action_in_progress
        if action_in_progress:
            return
        action_in_progress=True
        print("action_in_progress True")
        self.setEnabled(False)  # Disable the main window
        selected_cell_value = pyperclip.paste()
        selected_cell_value = string_to_list(selected_cell_value)
        if selected_cell_value is None:
            selected_cell_value = []
        if len(selected_cell_value) == 0:
            dialog_onebutton(self,"No ID(s) selected.","Error")
            action_in_progress=False
            print("action_in_progress False")
            self.setEnabled(True)  # Disable the main window
            return
        if len(selected_cell_value) >= 200:
            dialog_onebutton(self,"Please select 200 or fewer ID(s).","Error")
            action_in_progress=False
            print("action_in_progress False")
            self.setEnabled(True)  # Disable the main window
            return
        print("opening inspect window...")
        data_to_pass = 123  # Replace with actual data
        inspect_window = None
        inspect_window = window_inspect(self,data_to_pass)

if __name__ == '__main__':

    main_window = main_window()

    if settings.get('hotkey_open_value', 0) or settings.get('hotkey_inspect_value', 0):
        worker = listenerWorker()
        worker.open_inspect_signal.connect(main_window.open_inspect_window)
        worker.open_file_signal.connect(main_window.main_openfile)
        # Start the shortcut detection in the main thread
        print("starting shortcut thread")
        try:
            check_shortcut(worker)
            # shortcut_thread = threading.Thread(target=check_shortcut, args=(worker,))
            # shortcut_thread.start()
            print("shortcut_thread started successfully")
        except Exception as e:
            print(f"Error starting shortcut_thread: {e}")
    
    print("close loading screen")
    close_loading_screen()
    print("main_window show")
    main_window.show()


    if settings['run_count'] == 0:
        print("welcomewindow")
        from welcome import WelcomeDialog
        main_window.welcome_dialog = WelcomeDialog()
        main_window.welcome_dialog.exec()
        
    main_window.activateWindow()

    # increase run count
    settings = save_settings_key(settings, 'run_count', settings['run_count']+1)

    try:
        sys.exit(app.exec())
    except SystemExit:
        stop_listeners()
        print('Closing Application')