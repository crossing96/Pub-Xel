# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>

import datetime
import json
import os

import requests
import webbrowser
from PyQt6 import QtCore
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QMessageBox

from data.version import __version__
from pubxel_core.paths import appdatadir

CHECKFILE = os.path.join(appdatadir, "pubxel_check.json")


def should_check_for_update() -> bool:
    if not os.path.exists(CHECKFILE):
        print("CHECKFILE does not exist; should_check_for_update = TRUE")
        return True
    try:
        with open(CHECKFILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        last_check = datetime.datetime.fromisoformat(data.get("last_check", "1970-01-01"))
        print("current time: ", datetime.datetime.now().isoformat())
        print(f"Last update check was on {last_check.isoformat()}")
        print("Days since last check: ", (datetime.datetime.now() - last_check).days)
        print("should_check_for_update = ", (datetime.datetime.now() - last_check).days >= 7)
        return (datetime.datetime.now() - last_check).days >= 7
    except Exception:
        print("CHECKFILE open error; should_check_for_update = TRUE")
        return True


def save_check_timestamp() -> None:
    try:
        with open(CHECKFILE, "w", encoding="utf-8") as f:
            json.dump({"last_check": datetime.datetime.now().isoformat()}, f)
    except Exception:
        pass


def parse_version(v: str):
    try:
        parts = [int(x) for x in v.split(".")]
        while len(parts) < 3:
            parts.append(0)
        return tuple(parts[:3])
    except Exception:
        return (0, 0, 0)


def dialog_update(parent, message, title="Update Available"):
    msg_box = QMessageBox(parent)
    msg_box.setWindowTitle(title)
    msg_box.setText(message)
    msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
    msg_box.button(QMessageBox.StandardButton.Yes).setText("Update now")
    msg_box.button(QMessageBox.StandardButton.No).setText("Remind me in a week")
    msg_box.setWindowModality(Qt.WindowModality.ApplicationModal)
    msg_box.exec()
    msg_box.setFocus(QtCore.Qt.FocusReason.PopupFocusReason)
    return msg_box.clickedButton() == msg_box.button(QMessageBox.StandardButton.Yes)


def check_for_update(parent=None):
    if not should_check_for_update():
        return

    try:
        resp = requests.get(
            "https://raw.githubusercontent.com/crossing96/Pub-Xel/main/data/latest.json",
            timeout=3,
        )
        if not resp.ok:
            print("Update check failed (bad response).")
            return
        info = resp.json()
        latest = info.get("version", "")
        dl_url = info.get("download_url", "")
        current = __version__

        c_major, c_minor, c_patch = parse_version(current)
        l_major, l_minor, l_patch = parse_version(latest)

        print("new version: ", latest)
        print("current version: ", current)

        if (l_major > c_major) or (l_major == c_major and l_minor > c_minor) or (
            l_major == c_major and l_minor == c_minor and l_patch > c_patch
        ):
            message = f"A new version ({latest}) is available.\nYou're running {current}."
            if dialog_update(parent, message):
                if dl_url:
                    webbrowser.open(dl_url)
            save_check_timestamp()
        else:
            save_check_timestamp()
    except Exception:
        print("Update check failed.")
