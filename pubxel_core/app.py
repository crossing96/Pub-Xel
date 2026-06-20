# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# Application bootstrap and main event loop.

import atexit
import json
import os
import shutil
import sys

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import QApplication, QMessageBox, QSplashScreen

from pubxel_core import runtime as rt
from pubxel_core.paths import appdatadir, assets_dir, os_name, settings_path
from pubxel_core.settings import load_settings, save_settings, save_settings_key
from pubxel_core.ui.helpers import graceful_shutdown, stop_listeners
from pubxel_core.ui.main_window import main_window
from pubxel_core.ui.workers import check_shortcut, listenerWorker
from pubxel_core.update import check_for_update
from pubxel_core.worksheet_builder import create_worksheet

_bootstrapped = False


def _merge_default_settings(settingsdefault_path, settings_path):
    with open(settingsdefault_path, "r") as default_file:
        default_settings = json.load(default_file)
    with open(settings_path, "r") as settings_file:
        settings = json.load(settings_file)
    updated = False
    for key, value in default_settings.items():
        if key not in settings:
            settings[key] = value
            updated = True
            print("Updated settings:", key)
    if updated:
        with open(settings_path, "w") as settings_file:
            json.dump(settings, settings_file, indent=4)
    else:
        print("No settings updates needed.")


def _create_lock_file():
    if os_name == "Windows":
        import msvcrt

        rt.lock_file_path = os.path.join(appdatadir, "my_script.lock")
        rt.lock_file = open(rt.lock_file_path, "w")
        try:
            msvcrt.locking(rt.lock_file.fileno(), msvcrt.LK_NBLCK, 1)
        except IOError:
            QMessageBox.critical(None, "Error", "Pub-Xel is already running.")
            sys.exit(0)
    elif os_name == "Darwin":
        import fcntl

        rt.lock_file_path = os.path.join(appdatadir, "my_script.lock")
        rt.lock_file = open(rt.lock_file_path, "w")
        try:
            fcntl.flock(rt.lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except IOError:
            QMessageBox.critical(None, "Error", "Pub-Xel is already running.")
            sys.exit(0)
    else:
        raise Exception("Unsupported operating system")


def show_loading_screen():
    loading_image_path = os.path.join(assets_dir, "loading.png")
    pixmap = QPixmap(loading_image_path)
    pixmap = pixmap.scaled(500, 500, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
    rt.splash = QSplashScreen(pixmap, Qt.WindowType.WindowStaysOnTopHint)
    rt.splash.setWindowFlag(Qt.WindowType.FramelessWindowHint)
    rt.splash.show()
    rt.splash.activateWindow()
    rt.splash.raise_()


def close_loading_screen():
    if rt.splash is not None:
        rt.splash.close()
        rt.splash = None


def bootstrap():
    """Initialize QApplication, lock file, splash, and settings. Idempotent."""
    global _bootstrapped
    if _bootstrapped:
        return
    _bootstrapped = True

    if os_name not in ("Windows", "Darwin"):
        raise Exception("Unsupported operating system")

    rt.app = QApplication.instance() or QApplication(sys.argv)
    rt.app.setQuitOnLastWindowClosed(False)

    _create_lock_file()
    show_loading_screen()
    print("Loading screen shown")

    files_to_copy = {
        rt.settingsdefault_path: settings_path,
        rt.pubsheetinitialdefault_path: rt.pubsheetinitial_path,
    }
    for source_path, dest_path in files_to_copy.items():
        if dest_path == rt.pubsheetinitial_path or not os.path.exists(dest_path):
            shutil.copy(source_path, dest_path)
            print(f"Copied {source_path} as {dest_path}")

    _merge_default_settings(rt.settingsdefault_path, settings_path)
    rt.settings = load_settings()

    if not os.path.exists(rt.pubsheet_path):
        create_worksheet(rt.pubsheet_path, settings=rt.settings)
        print(f"Created {rt.pubsheet_path}")

    if os_name == "Windows":
        documents_path = os.path.join(os.environ["USERPROFILE"], "Documents")
    elif os_name == "Darwin":
        documents_path = os.path.join(os.environ["HOME"], "Documents")

    rt.mainlibdirdefault = os.path.join(documents_path, "Pub-Xel Library")
    rt.outdirdefault = documents_path

    if rt.settings.get("mainlib_path", 0) == "":
        rt.settings["mainlib_path"] = rt.mainlibdirdefault
        os.makedirs(rt.mainlibdirdefault, exist_ok=True)
        save_settings(rt.settings)
    rt.mainlibdir = rt.settings.get("mainlib_path", 0)

    if rt.settings.get("output_path", 0) == "":
        rt.settings["output_path"] = documents_path
        save_settings(rt.settings)
    rt.outdir = rt.settings.get("output_path", 0)

    if rt.settings.get("seclib_enable", 0):
        rt.seclibdir = rt.settings.get("seclib_path", [])
    else:
        rt.seclibdir = []

    if os_name == "Darwin":
        rt.settings = save_settings_key(rt.settings, "hotkey_inspect_value", "")
        rt.settings = save_settings_key(rt.settings, "hotkey_open_value", "")
        rt.settings = save_settings_key(rt.settings, "hotkey_pubmed_value", "")

    rt.system_tray_notice_shown = rt.settings.get("system_tray_notice_shown", 0)
    # PUBXEL_DEV=1 in the environment forces developer mode on for this run
    # without modifying settings.json. Useful for ad-hoc testing.
    _env_dev = os.environ.get("PUBXEL_DEV", "").strip().lower()
    if _env_dev in ("1", "true", "yes", "on"):
        rt.developerMode = 1
        print("Developer mode forced on by PUBXEL_DEV environment variable")
    else:
        rt.developerMode = rt.settings.get("developerMode", 0)

    atexit.register(graceful_shutdown)


def run():
    """Start Pub-Xel: bootstrap, show UI, run the Qt event loop."""
    bootstrap()

    window = main_window()

    if (
        rt.settings.get("hotkey_open_value", 0)
        or rt.settings.get("hotkey_inspect_value", 0)
        or rt.settings.get("hotkey_pubmed_value", 0)
    ):
        worker = listenerWorker()
        worker.signal_inspect.connect(window.open_inspect_window)
        worker.signal_open.connect(window.main_openfile)
        worker.signal_pubmed.connect(window.main_openpubmed)
        print("starting shortcut thread")
        try:
            check_shortcut(worker)
            print("shortcut_thread started successfully")
        except Exception as e:
            print(f"Error starting shortcut_thread: {e}")

    print("close loading screen")
    close_loading_screen()
    print("main_window show")

    check_for_update(parent=window)

    window.show()

    if rt.settings["run_count"] == 0:
        print("welcomewindow")
        from pubxel_core.welcome import WelcomeDialog

        window.welcome_dialog = WelcomeDialog()
        window.welcome_dialog.exec()

    window.activateWindow()

    rt.settings = save_settings_key(rt.settings, "run_count", rt.settings["run_count"] + 1)

    try:
        sys.exit(rt.app.exec())
    except SystemExit:
        stop_listeners()
        print("Closing Application")
