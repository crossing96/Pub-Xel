#!/usr/bin/env python3
"""Import all Pub-Xel modules without starting the GUI. Used in CI and local checks."""

from __future__ import annotations

import importlib
import sys
from pathlib import Path

# Repo root (parent of scripts/)
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


MODULES = [
    "data.version",
    "pubxel_core.paths",
    "pubxel_core.settings",
    "pubxel_core.ids",
    "pubxel_core.clipboard",
    "pubxel_core.metadata_store",
    "pubxel_core.excel_ops",
    "pubxel_core.pubmed",
    "pubxel_core.mainfunctions",
    "pubxel_core.runtime",
    "pubxel_core.update",
    "pubxel_core.welcome",
    "pubxel_core.ui.helpers",
    "pubxel_core.ui.workers",
    "pubxel_core.ui.tray",
    "pubxel_core.ui.widgets",
    "pubxel_core.ui.preferences",
    "pubxel_core.ui.dialogs_extra",
    "pubxel_core.ui.main_window",
    "pubxel_core.app",
    "pubxel_core",
]


def main() -> int:
    failed: list[tuple[str, str]] = []
    for name in MODULES:
        try:
            importlib.import_module(name)
            print(f"OK: {name}")
        except Exception as exc:
            failed.append((name, repr(exc)))
            print(f"FAIL: {name}: {exc}", file=sys.stderr)

    if failed:
        print("\nImport failures:", file=sys.stderr)
        for name, err in failed:
            print(f"  - {name}: {err}", file=sys.stderr)
        return 1

    # Quick API sanity checks (no network, no GUI)
    from pubxel_core.ids import string_to_list
    from pubxel_core.metadata_store import MetadataStore, metadataStore

    assert string_to_list("1|2") == ["1", "2"]
    assert MetadataStore is metadataStore
    print("OK: API sanity checks")
    return 0


if __name__ == "__main__":
    sys.exit(main())
