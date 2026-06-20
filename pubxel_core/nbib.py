# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>

from __future__ import annotations

import re
from typing import List, Tuple

from pubxel_core.pubmed import normalize_pmid

NBIB_EXCEL_MAX_ARTICLES = 200
PUBMED_Nbib_ERROR = "Can only import .nbib file from PubMed"


def parse_nbib_records(text: str) -> List[str]:
    """Split nbib file content into MEDLINE record segments."""
    text = text.replace("\r\n", "\n")
    if "\n\n" in text:
        parts = text.split("\n\n")
    else:
        parts = [text]
    return [part.strip() for part in parts if part.strip()]


def _pmid_from_record(record: str) -> str | None:
    for line in record.replace("\r\n", "\n").split("\n"):
        if line.startswith("PMID-"):
            raw = line.split("-", 1)[1].strip()
            if raw:
                return normalize_pmid(raw)
    return None


def validate_pubmed_nbib(records: List[str]) -> Tuple[List[str], int]:
    """
    Validate PubMed nbib records and return PMIDs in file order.

    Raises ValueError if any record lacks a PMID line.
    """
    pmids: List[str] = []
    for record in records:
        pmid = _pmid_from_record(record)
        if not pmid:
            raise ValueError(PUBMED_Nbib_ERROR)
        pmids.append(pmid)
    return pmids, len(pmids)


def load_nbib_file(path: str) -> Tuple[List[str], List[str], int]:
    """Read an nbib file and return (records, pmids, article_count)."""
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        text = f.read()
    records = parse_nbib_records(text)
    if not records:
        raise ValueError(PUBMED_Nbib_ERROR)
    pmids, count = validate_pubmed_nbib(records)
    return records, pmids, count
