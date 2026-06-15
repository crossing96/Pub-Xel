"""Create a new Pub-Xel worksheet from current settings.

Standalone utility script (not integrated into the app UI flow).
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import runpy
import shutil
import sys
from typing import Any

import xlwings as xw

# Allow running this script directly from outside project root.
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

DATA_DIR = os.path.join(PROJECT_ROOT, "data")


def _resolve_appdata_dir() -> str:
    os_name = platform.system()
    if os_name == "Windows":
        return os.path.join(os.getenv("APPDATA", ""), "pubxel")
    if os_name == "Darwin":
        return os.path.expanduser("~/Library/Application Support/pubxel")
    return os.path.join(os.path.expanduser("~"), ".pubxel")


APPDATA_DIR = _resolve_appdata_dir()


def _load_default_layout() -> list[dict[str, Any]]:
    module_path = os.path.join(PROJECT_ROOT, "pubxel_core", "worksheet_defaults.py")
    module_vars = runpy.run_path(module_path)
    layout = module_vars.get("DEFAULT_WORKSHEET_COLUMN_LAYOUT", [])
    if not isinstance(layout, list):
        raise ValueError("DEFAULT_WORKSHEET_COLUMN_LAYOUT must be a list")
    return layout


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on"}
    return False


def _load_json(path: str) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError(f"Expected JSON object in {path}")
    return data


def _resolve_template_path(template_path: str | None) -> str:
    if template_path:
        candidate = os.path.abspath(template_path)
        if not os.path.exists(candidate):
            raise FileNotFoundError(f"Template not found: {candidate}")
        return candidate

    candidates = [
        os.path.join(APPDATA_DIR, "pubsheet.xlsx"),
        os.path.join(DATA_DIR, "pubsheet.xlsx"),
    ]
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    raise FileNotFoundError("Could not find pubsheet.xlsx in appdata or data folder.")


def _build_column_specs(settings: dict[str, Any]) -> list[dict[str, Any]]:
    default_layout = _load_default_layout()
    enabled_map = settings.get("worksheet_column_enabled", {})
    width_overrides = settings.get("worksheet_column_width", {})
    wrap_overrides = settings.get("worksheet_column_wrap", {})

    columns: list[dict[str, Any]] = []
    for item in default_layout:
        name = item["column_name"]
        enabled_raw = enabled_map.get(name, 1)
        if not _as_bool(enabled_raw):
            continue

        width = width_overrides.get(name, item["default_width"])
        wrap = wrap_overrides.get(name, item["default_wrap"])
        columns.append(
            {
                "column_name": name,
                "width": float(width),
                "wrap": _as_bool(wrap),
            }
        )
    return columns


def create_worksheet(
    output_path: str,
    settings_path: str | None = None,
    template_path: str | None = None,
) -> str:
    resolved_settings_path = (
        os.path.abspath(settings_path)
        if settings_path
        else os.path.join(APPDATA_DIR, "settings.json")
    )
    if not os.path.exists(resolved_settings_path):
        raise FileNotFoundError(f"Settings file not found: {resolved_settings_path}")

    resolved_template_path = _resolve_template_path(template_path)
    settings = _load_json(resolved_settings_path)
    column_specs = _build_column_specs(settings)
    if not column_specs:
        raise ValueError("No enabled worksheet columns found in settings.")

    resolved_output_path = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(resolved_output_path), exist_ok=True)
    shutil.copyfile(resolved_template_path, resolved_output_path)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open(resolved_output_path)
        try:
            ws = wb.sheets[0]
            if ws.api.ListObjects.Count < 1:
                raise ValueError("Template worksheet must contain at least one Excel Table.")

            table = ws.api.ListObjects.Item(1)
            table_row_count = int(table.Range.Rows.Count)
            if table_row_count < 2:
                # At minimum, keep header row + one data row in the table.
                table_row_count = 2

            # Expand the existing table to match the enabled columns while preserving table style.
            resized_range = ws.range(
                (1, 1),
                (table_row_count, len(column_specs)),
            ).api
            table.Resize(resized_range)

            for idx, spec in enumerate(column_specs, start=1):
                col_name = spec["column_name"]
                col_width = spec["width"]
                col_wrap = spec["wrap"]

                ws.range((1, idx)).value = col_name
                ws.range((1, idx)).column_width = col_width

                # If wrap is enabled, apply to the whole column (row 1 through Excel max row).
                if col_wrap:
                    ws.range((1, idx), (1048576, idx)).api.WrapText = True
                else:
                    ws.range((1, idx), (1048576, idx)).api.WrapText = False

            # Keep table formatting, but start with an empty data body.
            if table.DataBodyRange is not None:
                table.DataBodyRange.ClearContents()

            wb.save()
        finally:
            wb.close()
    finally:
        app.quit()

    return resolved_output_path


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Create a new worksheet using current pubsheet template + appdata settings."
        )
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Output .xlsx path for the new worksheet.",
    )
    parser.add_argument(
        "--settings",
        default=None,
        help="Optional settings.json path. Default: appdata settings.json",
    )
    parser.add_argument(
        "--template",
        default=None,
        help="Optional pubsheet.xlsx path. Default: appdata pubsheet.xlsx or data/pubsheet.xlsx",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    out = create_worksheet(
        output_path=args.output,
        settings_path=args.settings,
        template_path=args.template,
    )
    print(f"Worksheet created: {out}")
