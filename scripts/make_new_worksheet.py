"""Create a new Pub-Xel worksheet from current settings.

Standalone utility script (not integrated into the app UI flow).
"""

from __future__ import annotations

import argparse
import os
import sys

# Allow running this script directly from outside project root.
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from pubxel_core.worksheet_builder import create_worksheet


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Create a new worksheet from pubsheet_all_columns.xlsx + app settings."
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
        help="Optional pubsheet_all_columns.xlsx path.",
    )
    parser.add_argument(
        "--all-columns",
        action="store_true",
        help="Keep every column from pubsheet_all_columns.xlsx (ignore enabled checkboxes).",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    out = create_worksheet(
        args.output,
        settings_path=args.settings,
        template_path=args.template,
        all_columns=args.all_columns,
    )
    print(f"Worksheet created: {out}")
