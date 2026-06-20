# Shared runtime paths and mutable application state (set during app bootstrap).

import logging
import os
import threading
from contextlib import contextmanager

from pubxel_core.paths import appdatadir, assets_dir, data_dir, os_name, settings_path, ui_dir

logger = logging.getLogger(__name__)

# Bundled / install UI asset paths
main_path = os.path.join(ui_dir, "main.ui")
inspect_path = os.path.join(ui_dir, "inspect.ui")
about_path = os.path.join(ui_dir, "about.ui")
preferences_path = os.path.join(ui_dir, "preferences.ui")
icon_path = os.path.join(assets_dir, "logo64.ico")
questionmark_icon_path = os.path.join(assets_dir, "questionmark.png")
loading_image_path = os.path.join(assets_dir, "loading.png")
settingsdefault_path = os.path.join(data_dir, "settingsdefault.json")
pubsheet_no_columns_path = os.path.join(data_dir, "pubsheet_no_columns.xlsx")

# Populated by pubxel_core.app.bootstrap() before the main window is shown
app = None
settings = {}
mainlibdir = ""
seclibdir = []
outdir = ""
mainlibdirdefault = ""
outdirdefault = ""
system_tray_notice_shown = 0
developerMode = 0

action_in_progress = False
force_quit = False

listeners = []
splash = None
lock_file = None
lock_file_path = None


# --- Mutual-exclusion helpers for `action_in_progress` ----------------------
#
# All writes to `action_in_progress` should go through these helpers so the
# check-and-set is atomic across threads (the Qt main thread plus the pynput
# hotkey listener thread). Reads remain free — a single bool read is atomic
# under the GIL and read-only callers (e.g. hotkey gating in
# pubxel_core/ui/workers.py) don't need the lock.
#
# Two patterns are supported:
#
#   1. Modal/synchronous owners (caller holds the flag for the whole action):
#
#        if not rt.try_begin_action():
#            return
#        try:
#            ... do work / open modal dialog ...
#        finally:
#            rt.end_action()
#
#      Or equivalently with the context manager:
#
#        with rt.action_guard() as acquired:
#            if not acquired:
#                return
#            ... do work ...
#
#   2. Non-modal owners (a long-lived window owns the flag past the
#      caller's stack frame, e.g. window_inspect): the caller acquires
#      with try_begin_action() and only releases on the early-exit paths
#      where the window was never actually shown. The window itself
#      releases via end_action() in its closeEvent.

_action_lock = threading.Lock()


def try_begin_action() -> bool:
    """Atomically claim ``action_in_progress``.

    Returns True if the flag was free and is now held by the caller; False
    if another action is already in progress (caller should bail out).
    """
    global action_in_progress
    with _action_lock:
        if action_in_progress:
            print("action_in_progress: blocked (already True)")
            logger.debug("action_in_progress: blocked (already True)")
            return False
        action_in_progress = True
        print("action_in_progress -> True")
        logger.debug("action_in_progress -> True")
        return True


def end_action() -> None:
    """Release ``action_in_progress``. Safe to call when already False."""
    global action_in_progress
    with _action_lock:
        if action_in_progress:
            print("action_in_progress -> False")
            logger.debug("action_in_progress -> False")
        else:
            print("action_in_progress: end_action called but already False")
            logger.debug("action_in_progress: end_action called but already False")
        action_in_progress = False


@contextmanager
def action_guard():
    """Context manager that acquires ``action_in_progress`` for its block.

    Yields True if the flag was acquired (caller should proceed) or False if
    another action is in progress (caller should return). When True, the
    flag is released automatically on exit — including via exception.
    """
    acquired = try_begin_action()
    try:
        yield acquired
    finally:
        if acquired:
            end_action()
