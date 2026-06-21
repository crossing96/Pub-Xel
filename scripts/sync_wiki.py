# Pub-Xel - sync wiki/*.md to a GitHub Wiki git checkout.
# Strips .md from internal page links (GitHub Wiki convention).

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

SKIP_FILES = {"README.md"}
# [Page](Page.md) or [Page](Page.md#anchor) -> drop .md before ) or #
INTERNAL_WIKI_LINK = re.compile(
    r"(\]\()([A-Za-z0-9_-]+)\.md(?=[)#])"
)


def transform_markdown(text: str) -> str:
    return INTERNAL_WIKI_LINK.sub(r"\1\2", text)


def sync_wiki(source_dir: Path, dest_dir: Path) -> list[str]:
    if not source_dir.is_dir():
        raise FileNotFoundError(f"Wiki source not found: {source_dir}")

    dest_dir.mkdir(parents=True, exist_ok=True)
    written: set[str] = set()

    for src in sorted(source_dir.glob("*.md")):
        if src.name in SKIP_FILES:
            continue
        content = src.read_text(encoding="utf-8")
        (dest_dir / src.name).write_text(transform_markdown(content), encoding="utf-8")
        written.add(src.name)

    for existing in dest_dir.glob("*.md"):
        if existing.name not in written:
            existing.unlink()

    return sorted(written)


def main() -> int:
    parser = argparse.ArgumentParser(description="Sync wiki/ to a GitHub Wiki checkout.")
    parser.add_argument(
        "source",
        nargs="?",
        default="wiki",
        help="Source directory (default: wiki)",
    )
    parser.add_argument(
        "dest",
        nargs="?",
        default="Pub-Xel.wiki",
        help="GitHub Wiki clone directory (default: Pub-Xel.wiki)",
    )
    args = parser.parse_args()

    source_dir = Path(args.source).resolve()
    dest_dir = Path(args.dest).resolve()

    try:
        pages = sync_wiki(source_dir, dest_dir)
    except FileNotFoundError as exc:
        print(exc, file=sys.stderr)
        return 1

    print(f"Synced {len(pages)} page(s) to {dest_dir}:")
    for name in pages:
        print(f"  {name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
