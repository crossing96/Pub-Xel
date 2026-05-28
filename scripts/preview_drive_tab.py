#!/usr/bin/env python3
"""Open ui/drive_tab.ui in a live Qt window for visual review.

Run from the repo root:
    python scripts/preview_drive_tab.py

Nothing here is wired to any backend — every click prints a stub message so
you can confirm the layout, sizing, and enable/disable behaviour before the
final tab is pasted into ui/preferences.ui.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PyQt6 import uic
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QGridLayout,
    QGroupBox,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPlainTextEdit,
    QPushButton,
)

UI_PATH = ROOT / "ui" / "drive_tab.ui"
QUESTIONMARK_PATH = ROOT / "assets" / "questionmark.png"


class DriveTabPreview(QDialog):
    """Standalone preview of the Drive settings tab."""

    def __init__(self) -> None:
        super().__init__()
        uic.loadUi(str(UI_PATH), self)

        self._populate_questionmark()
        self._wire_stub_actions()

    # --- question-mark icon (matches the layout_q1/layout_q2 pattern in main.ui) ---
    def _populate_questionmark(self) -> None:
        layout_q = self.findChild(QGridLayout, "layout_q_drive")
        if layout_q is None:
            return
        if not QUESTIONMARK_PATH.exists():
            print(f"(preview) questionmark icon not found at {QUESTIONMARK_PATH}")
            return
        pixmap = QPixmap(str(QUESTIONMARK_PATH))
        self.label_q_drive = QLabel("")
        self.label_q_drive.setPixmap(pixmap)
        self.label_q_drive.setCursor(Qt.CursorShape.PointingHandCursor)
        self.label_q_drive.setToolTip(
            "Pub-Xel will open your browser to ask Google for read-only "
            "access to Drive. Click for details."
        )
        self.label_q_drive.mousePressEvent = self._on_help_clicked
        layout_q.addWidget(self.label_q_drive)

    def _on_help_clicked(self, _event) -> None:
        print("(preview) help icon clicked")

    # --- stub wiring so the preview is interactive ---
    def _wire_stub_actions(self) -> None:
        self.button_drive_connect = self.findChild(QPushButton, "button_drive_connect")
        self.button_drive_folder_add = self.findChild(QPushButton, "button_drive_folder_add")
        self.button_drive_folder_remove = self.findChild(QPushButton, "button_drive_folder_remove")
        self.button_drive_folder_open = self.findChild(QPushButton, "button_drive_folder_open")
        self.button_drive_download_dir_select = self.findChild(
            QPushButton, "button_drive_download_dir_select"
        )
        self.button_drive_download_dir_default = self.findChild(
            QPushButton, "button_drive_download_dir_default"
        )

        self.label_drive_status = self.findChild(QLabel, "label_drive_status")
        self.label_drive_account = self.findChild(QLabel, "label_drive_account")
        self.list_drive_folders = self.findChild(QListWidget, "list_drive_folders")
        self.plainTextEdit_drive_download_dir = self.findChild(
            QPlainTextEdit, "plainTextEdit_drive_download_dir"
        )
        self.groupBox_drive_folders = self.findChild(QGroupBox, "groupBox_drive_folders")
        self.groupBox_drive_download = self.findChild(QGroupBox, "groupBox_drive_download")
        self.checkbox_drive_recursive = self.findChild(QCheckBox, "checkbox_drive_recursive")

        # Fake "connected" toggle so the preview can demo both states
        self._connected = False
        self.button_drive_connect.clicked.connect(self._toggle_connected)
        self.button_drive_folder_add.clicked.connect(self._stub_add_folder)
        self.button_drive_folder_remove.clicked.connect(self._stub_remove_folder)
        self.button_drive_folder_open.clicked.connect(self._stub_open_folder)
        self.button_drive_download_dir_select.clicked.connect(self._stub_select_dir)
        self.button_drive_download_dir_default.clicked.connect(self._stub_use_default_dir)

    def _toggle_connected(self) -> None:
        self._connected = not self._connected
        if self._connected:
            self.label_drive_status.setText("Connected")
            self.label_drive_status.setStyleSheet("QLabel { color: #2a8a2a; }")
            self.label_drive_account.setText("preview-user@example.com")
            self.button_drive_connect.setText("Sign out")
            self.groupBox_drive_folders.setEnabled(True)
            self.groupBox_drive_download.setEnabled(True)
        else:
            self.label_drive_status.setText("Not connected")
            self.label_drive_status.setStyleSheet("QLabel { color: #888; }")
            self.label_drive_account.setText("—")
            self.button_drive_connect.setText("Connect Google Drive…")
            self.groupBox_drive_folders.setEnabled(False)
            self.groupBox_drive_download.setEnabled(False)
        print(f"(preview) connected = {self._connected}")

    def _stub_add_folder(self) -> None:
        text, ok = QInputDialog.getText(
            self,
            "Add Drive Folder",
            "Paste a Drive folder URL or ID:",
        )
        if not ok or not text.strip():
            return
        folder_id = text.strip()
        item = QListWidgetItem(f"(stub name)  —  {folder_id}")
        self.list_drive_folders.addItem(item)
        print(f"(preview) added folder: {folder_id}")

    def _stub_remove_folder(self) -> None:
        row = self.list_drive_folders.currentRow()
        if row < 0:
            print("(preview) nothing selected to remove")
            return
        item = self.list_drive_folders.takeItem(row)
        print(f"(preview) removed: {item.text() if item else '?'}")

    def _stub_open_folder(self) -> None:
        item = self.list_drive_folders.currentItem()
        print(f"(preview) would open in Drive: {item.text() if item else '(no selection)'}")

    def _stub_select_dir(self) -> None:
        from PyQt6.QtWidgets import QFileDialog

        path = QFileDialog.getExistingDirectory(self, "Select Download Folder")
        if path:
            self.plainTextEdit_drive_download_dir.setPlainText(path)
            print(f"(preview) download dir = {path}")

    def _stub_use_default_dir(self) -> None:
        self.plainTextEdit_drive_download_dir.setPlainText("")
        print("(preview) download dir reset to default")


def main() -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    dlg = DriveTabPreview()
    dlg.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
