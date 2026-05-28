# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

"""Recent Excel worksheet paths (settings + optional File menu refresh)."""

from __future__ import annotations

import os
from typing import Callable, List, Optional

from pubxel_core import runtime as rt
from pubxel_core.settings import load_settings, save_settings_key

SETTINGS_KEY = "recent_worksheets"
MAX_RECENT = 10

_menu_rebuild_callback: Optional[Callable[[], None]] = None


def set_recent_menu_rebuild_callback(callback: Optional[Callable[[], None]]) -> None:
    """Register a main-window hook to refresh File → recent entries (set once at startup)."""
    global _menu_rebuild_callback
    _menu_rebuild_callback = callback


def _notify_menu_rebuild() -> None:
    if _menu_rebuild_callback is not None:
        _menu_rebuild_callback()


def _normalize_path(path: str) -> Optional[str]:
    if not path or not str(path).strip():
        return None
    path = os.path.normpath(os.path.abspath(str(path).strip()))
    if not path.lower().endswith(".xlsx"):
        return None
    return path


def get_recent_worksheets() -> List[str]:
    """Return stored recent worksheet paths (newest first)."""
    settings = rt.settings if rt.settings else load_settings()
    raw = settings.get(SETTINGS_KEY, [])
    if not isinstance(raw, list):
        return []
    return [str(p) for p in raw if p]


def register_recent_worksheet(path: str) -> bool:
    """
    Add or bump a worksheet path in recent history (newest first, max 10).

    Callable from anywhere after save/import. Returns False if path ignored.
    """
    normalized = _normalize_path(path)
    if not normalized:
        return False

    settings = rt.settings if rt.settings else load_settings()
    recent = get_recent_worksheets()
    recent = [p for p in recent if os.path.normcase(p) != os.path.normcase(normalized)]
    recent.insert(0, normalized)
    recent = recent[:MAX_RECENT]

    rt.settings = save_settings_key(settings, SETTINGS_KEY, recent)
    _notify_menu_rebuild()
    return True


def remove_recent_worksheet(path: str) -> None:
    """Remove one path from recent history."""
    normalized = _normalize_path(path) or path
    settings = rt.settings if rt.settings else load_settings()
    recent = get_recent_worksheets()
    recent = [
        p
        for p in recent
        if os.path.normcase(p) != os.path.normcase(normalized)
        and os.path.normcase(p) != os.path.normcase(str(path).strip())
    ]
    rt.settings = save_settings_key(settings, SETTINGS_KEY, recent)
    _notify_menu_rebuild()


def format_recent_menu_label(path: str, max_len: int = 72) -> str:
    if len(path) <= max_len:
        return path
    return "…" + path[-(max_len - 1) :]


def try_register_active_workbook() -> bool:
    """If Excel has a saved .xlsx workbook open, add it to recent list."""
    try:
        import xlwings as xw

        wb = xw.books.active
        path = wb.fullname
    except Exception:
        return False
    return register_recent_worksheet(path)
