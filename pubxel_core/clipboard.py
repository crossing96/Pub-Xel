# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

"""
Qt clipboard reader (main thread only).

Classification order in read_clipboard():
  1. image  — screenshot / image-only clipboard (unchanged messaging)
  2. empty  — nothing usable
  3. files  — Explorer/Finder copy (local file URLs); ONLY case with file_paths set
  4. text   — plain text / Excel / pasted paths / PubMed links: split on delimiters,
              then each token -> PMID, file stem, or original string

ClipboardRead usage by action (do not swap without updating this doc):

  main_openfile (Open Files button / hotkey)
    - kind == "files"  -> open clip.file_paths directly (Explorer/Finder copy).
    - any other kind   -> clip.ids only -> process_ids -> library folders.
    - Never use file_paths when kind is not "files".

  main_openpubmed, open_inspect_window (and window_inspect via clipboard_ids)
    - Always clip.ids only.
    - Never use file_paths, even when kind == "files".

  update_clipboard_current
    - clip.preview only.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import List, Literal, Optional

from PyQt6.QtGui import QGuiApplication

from pubxel_core.ids import (
    ids_from_clipboard_text,
    list_to_string,
    set_preserve_order,
    strip_last_extension,
)

ClipboardKind = Literal["text", "files", "image", "empty", "unsupported"]

_MSG_IMAGE = (
    "Clipboard contains an image. Copy article ID(s) as text, "
    "or copy PDF file(s) from your file manager."
)
_MSG_EMPTY = "Clipboard is empty."
_MSG_FILES_NO_IDS = "No ID(s) could be extracted from copied file name(s)."


@dataclass
class ClipboardRead:
    kind: ClipboardKind
    # Parsed ID strings (PMIDs, file stems, non-PubMed IDs, etc.).
    ids: List[str] = field(default_factory=list)
    # Absolute local paths ONLY when kind=="files" (Explorer/Finder copy).
    file_paths: List[str] = field(default_factory=list)
    preview: str = ""
    user_message: Optional[str] = None


def _preview_for_files(file_paths: List[str]) -> str:
    if not file_paths:
        return "No file(s) on clipboard."
    names = [os.path.basename(p) for p in file_paths]
    n = len(names)
    shown = names[:3]
    body = ", ".join(shown)
    if n > 3:
        body = f"{body}, ..."
    return f"{n} file(s): {body}"


def _ids_from_file_paths(file_paths: List[str]) -> List[str]:
    stems = [strip_last_extension(os.path.basename(p)) for p in file_paths]
    return list(set_preserve_order(stems))


def _collect_local_file_paths(mime) -> List[str]:
    paths: List[str] = []
    if mime.hasUrls():
        for url in mime.urls():
            if url.isLocalFile():
                path = url.toLocalFile()
                if path:
                    paths.append(path)
    return list(set_preserve_order(paths))


def _clipboard_plain_text_blob(mime) -> str:
    """Plain text plus non-local URL strings (e.g. copied PubMed link without text/plain)."""
    parts: List[str] = []
    if mime.hasText():
        text = (mime.text() or "").strip()
        if text:
            parts.append(text)
    if mime.hasUrls():
        for url in mime.urls():
            if not url.isLocalFile():
                s = url.toString().strip()
                if s:
                    parts.append(s)
    return "\n".join(parts)


def read_clipboard() -> ClipboardRead:
    """
    Read the system clipboard via Qt (call from the Qt main thread only).
    """
    clipboard = QGuiApplication.clipboard()
    mime = clipboard.mimeData() if clipboard is not None else None
    if mime is None:
        return ClipboardRead(
            kind="empty",
            preview="Clipboard is empty.",
            user_message=_MSG_EMPTY,
        )

    file_paths = _collect_local_file_paths(mime)
    text_blob = _clipboard_plain_text_blob(mime)

    # 1. Image-only clipboard (no Explorer files, no plain-text blob)
    if mime.hasImage() and not file_paths and not text_blob:
        return ClipboardRead(
            kind="image",
            preview="Clipboard: image",
            user_message=_MSG_IMAGE,
        )

    # 2. Empty
    if not file_paths and not text_blob and not mime.hasImage():
        return ClipboardRead(
            kind="empty",
            preview="Clipboard is empty.",
            user_message=_MSG_EMPTY,
        )

    # 3. Explorer / Finder file copy
    if file_paths:
        ids = _ids_from_file_paths(file_paths)
        return ClipboardRead(
            kind="files",
            ids=ids,
            file_paths=file_paths,
            preview=_preview_for_files(file_paths),
            user_message=None if ids else _MSG_FILES_NO_IDS,
        )

    # 4. Text-like: delimiter split, then per-token processing (see ids_from_clipboard_text)
    ids = ids_from_clipboard_text(text_blob)
    if ids:
        return ClipboardRead(
            kind="text",
            ids=ids,
            preview=list_to_string(ids),
            user_message=None,
        )

    if mime.hasImage():
        return ClipboardRead(
            kind="image",
            preview="Clipboard: image",
            user_message=_MSG_IMAGE,
        )

    return ClipboardRead(
        kind="empty",
        preview="Clipboard is empty.",
        user_message=_MSG_EMPTY,
    )


def write_clipboard(text: str) -> None:
    """Write plain text to the system clipboard (call from the Qt main thread only)."""
    clipboard = QGuiApplication.clipboard()
    if clipboard is not None:
        clipboard.setText(text or "")


def message_for_action(clip: ClipboardRead, *, needs_ids: bool = True) -> Optional[str]:
    """Return a user-facing error message if ``clip`` cannot supply ID(s) for an action."""
    if needs_ids and clip.ids:
        return None
    if needs_ids and not clip.ids:
        if clip.kind == "files":
            return clip.user_message or _MSG_FILES_NO_IDS
        if clip.kind == "text":
            return "No ID(s) on clipboard."
        return clip.user_message or _MSG_EMPTY
    return None
