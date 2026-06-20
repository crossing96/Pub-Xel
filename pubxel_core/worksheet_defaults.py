# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This module stores default worksheet column layout metadata for future
# worksheet generation logic. It is intentionally data-only for now.

from __future__ import annotations

from typing import List, TypedDict


class WorksheetColumnDefault(TypedDict):
    column_name: str
    default_width: float
    default_wrap: bool


DEFAULT_WORKSHEET_COLUMN_LAYOUT: List[WorksheetColumnDefault] = [
    {"column_name": "Ref", "default_width": 10.25, "default_wrap": False},
    {"column_name": "DOI", "default_width": 8.25, "default_wrap": False},
    {"column_name": "AuthorYear", "default_width": 16.25, "default_wrap": True},
    {"column_name": "Authors", "default_width": 8.25, "default_wrap": False},
    {"column_name": "Year", "default_width": 8.25, "default_wrap": False},
    {"column_name": "Journal", "default_width": 12.25, "default_wrap": True},
    {"column_name": "Title", "default_width": 52.38, "default_wrap": True},
    {"column_name": "Abstract", "default_width": 8.25, "default_wrap": False},
    {"column_name": "Citation", "default_width": 73.31, "default_wrap": True},
    {"column_name": "Citation2025", "default_width": 73.31, "default_wrap": True},
    {"column_name": "IF2025", "default_width": 8.25, "default_wrap": False},
    {"column_name": "Q2025", "default_width": 8.25, "default_wrap": False},
    {"column_name": "Identifier", "default_width": 8.25, "default_wrap": False},
    {"column_name": "Funding", "default_width": 8.25, "default_wrap": False},
]

