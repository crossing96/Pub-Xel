# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# Build Pub-Xel Excel worksheets from pubsheet_all_columns.xlsx by removing
# columns disabled in preferences.

from __future__ import annotations

import json
import os
import shutil
from typing import Any

import xlwings as xw

from pubxel_core.metadata_store import MetadataDict
from pubxel_core.paths import appdatadir, pubsheet_all_columns_path
from pubxel_core.pubmed import build_worksheet_header_map, fill_worksheet_rows, normalize_pmid_list
from pubxel_core.settings import SettingsDict
from pubxel_core.worksheet_defaults import DEFAULT_WORKSHEET_COLUMN_LAYOUT


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on"}
    return False


def _load_settings(settings: SettingsDict | None, settings_path: str | None) -> SettingsDict:
    if settings is not None:
        return settings
    resolved = settings_path or os.path.join(appdatadir, "settings.json")
    with open(resolved, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError(f"Expected JSON object in {resolved}")
    return data


def _resolve_template_path(template_path: str | None) -> str:
    if template_path:
        candidate = os.path.abspath(template_path)
        if not os.path.exists(candidate):
            raise FileNotFoundError(f"Template not found: {candidate}")
        return candidate
    if not os.path.exists(pubsheet_all_columns_path):
        raise FileNotFoundError(f"Template not found: {pubsheet_all_columns_path}")
    return pubsheet_all_columns_path


def build_column_specs(settings: SettingsDict, *, all_columns: bool = False) -> list[str]:
    enabled_map = settings.get("worksheet_column_enabled", {})

    columns: list[str] = []
    for item in DEFAULT_WORKSHEET_COLUMN_LAYOUT:
        name = item["column_name"]
        if not all_columns:
            enabled_raw = enabled_map.get(name, 1)
            if not _as_bool(enabled_raw):
                continue
        columns.append(name)
    return columns


def _read_header_map(ws: xw.Sheet) -> dict[str, int]:
    headers: dict[str, int] = {}
    col = 1
    while True:
        value = ws.range((1, col)).value
        if value is None or str(value).strip() == "":
            break
        headers[str(value).strip()] = col
        col += 1
    return headers


def create_worksheet(
    output_path: str,
    *,
    settings: SettingsDict | None = None,
    settings_path: str | None = None,
    template_path: str | None = None,
    all_columns: bool = False,
) -> str:
    """Save a new worksheet by copying pubsheet_all_columns and deleting disabled columns."""
    resolved_settings = _load_settings(settings, settings_path)
    column_specs = build_column_specs(resolved_settings, all_columns=all_columns)
    if not column_specs:
        raise ValueError("No worksheet columns selected in preferences.")

    resolved_template_path = _resolve_template_path(template_path)
    resolved_output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(resolved_output_path) or ".", exist_ok=True)
    shutil.copyfile(resolved_template_path, resolved_output_path)

    keep_names = set(column_specs)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open(resolved_output_path)
        try:
            ws = wb.sheets[0]
            if ws.api.ListObjects.Count < 1:
                raise ValueError("Template worksheet must contain at least one Excel Table.")

            header_map = _read_header_map(ws)
            missing = [name for name in keep_names if name not in header_map]
            if missing:
                raise ValueError(
                    "Template is missing expected column(s): " + ", ".join(missing)
                )

            remove_indices = sorted(
                (
                    header_map[name]
                    for name in header_map
                    if name not in keep_names
                ),
                reverse=True,
            )
            for col_index in remove_indices:
                ws.api.Columns(col_index).Delete()

            wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    return resolved_output_path


def create_filled_worksheet(
    output_path: str,
    pmids: list[str],
    metadata: MetadataDict,
    *,
    settings: SettingsDict | None = None,
    settings_path: str | None = None,
    template_path: str | None = None,
) -> str:
    """Create a trimmed worksheet and fill rows from preloaded PubMed metadata."""
    normalized_pmids = normalize_pmid_list(pmids)
    if not normalized_pmids:
        raise ValueError("No PubMed ID(s) to include in the worksheet.")
    if not metadata:
        raise ValueError("No PubMed metadata available to fill the worksheet.")

    resolved_settings = _load_settings(settings, settings_path)
    column_specs = build_column_specs(resolved_settings)
    if not column_specs:
        raise ValueError("No worksheet columns selected in preferences.")

    resolved_template_path = _resolve_template_path(template_path)
    resolved_output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(resolved_output_path) or ".", exist_ok=True)
    shutil.copyfile(resolved_template_path, resolved_output_path)

    keep_names = set(column_specs)
    num_data_rows = len(normalized_pmids)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open(resolved_output_path)
        try:
            ws = wb.sheets[0]
            if ws.api.ListObjects.Count < 1:
                raise ValueError("Template worksheet must contain at least one Excel Table.")

            header_map = _read_header_map(ws)
            missing = [name for name in keep_names if name not in header_map]
            if missing:
                raise ValueError(
                    "Template is missing expected column(s): " + ", ".join(missing)
                )

            remove_indices = sorted(
                (header_map[name] for name in header_map if name not in keep_names),
                reverse=True,
            )
            for col_index in remove_indices:
                ws.api.Columns(col_index).Delete()

            num_cols = len(column_specs)
            table = ws.api.ListObjects.Item(1)
            table_row_count = max(2, 1 + num_data_rows)
            resized_range = ws.range((1, 1), (table_row_count, num_cols)).api
            table.Resize(resized_range)

            header_names = [ws.range((1, col)).value for col in range(1, num_cols + 1)]
            header2 = build_worksheet_header_map(header_names)
            rows = [(1 + idx, pmid) for idx, pmid in enumerate(normalized_pmids)]
            fill_worksheet_rows(ws, metadata, header2, rows)

            wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    return resolved_output_path
