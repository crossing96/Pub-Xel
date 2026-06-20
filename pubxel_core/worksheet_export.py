# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>

from __future__ import annotations

import csv
import os
from typing import List

from pubxel_core.metadata_store import MetadataDict
from pubxel_core.pubmed import iter_worksheet_export_rows, normalize_pmid_list
from pubxel_core.settings import SettingsDict
from pubxel_core.worksheet_builder import build_column_specs


def write_worksheet_tsv(
    path: str,
    pmids: List[str],
    metadata: MetadataDict,
    settings: SettingsDict,
) -> str:
    """Write worksheet data as a UTF-8 TSV file (no xlwings)."""
    normalized_pmids = normalize_pmid_list(pmids)
    if not normalized_pmids:
        raise ValueError("No PubMed ID(s) to include in the export.")
    if not metadata:
        raise ValueError("No PubMed metadata available for export.")

    column_names = build_column_specs(settings)
    if not column_names:
        raise ValueError("No worksheet columns selected in preferences.")

    header_row, data_rows = iter_worksheet_export_rows(
        normalized_pmids, metadata, column_names
    )

    resolved_path = os.path.abspath(path)
    os.makedirs(os.path.dirname(resolved_path) or ".", exist_ok=True)

    with open(resolved_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter="\t", lineterminator="\n")
        writer.writerow(header_row)
        writer.writerows(data_rows)

    return resolved_path
