# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

import os
import re
from typing import Iterable, List, Optional

# PMIDs in PubMed article URLs (modern /pubmed/ legacy paths).
_PUBMED_URL_PMID = re.compile(
    r"(?:pubmed\.ncbi\.nlm\.nih\.gov/|ncbi\.nlm\.nih\.gov/pubmed/)(\d{1,9})\b",
    re.IGNORECASE,
)


def strip_last_extension(name: str) -> str:
    """Return basename stem: no dot unchanged; else drop last ``.`` and suffix."""
    if "." not in name:
        return name
    return name.rsplit(".", 1)[0]


def set_preserve_order(input_list: Iterable[str]) -> List[str]:
    seen: set[str] = set()
    return [x for x in input_list if not (x in seen or seen.add(x))]


def pmids_from_pubmed_urls_in_text(text: str) -> List[str]:
    """Extract PMIDs from PubMed article URLs in ``text`` (order preserved, deduped)."""
    if not text or not text.strip():
        return []
    pmids = [
        m
        for m in _PUBMED_URL_PMID.findall(text)
        if m.isdigit() and len(m) <= 9
    ]
    return list(set_preserve_order(pmids))


def split_clipboard_text(text: str) -> List[str]:
    """
    Split clipboard plain text on Excel-style delimiters (tab, line breaks, angle brackets).

    Colon is intentionally not a delimiter so PubMed URLs stay a single token.
    """
    if not text or not text.strip():
        return []
    s = text
    s = s.replace("\t", "|")
    s = s.replace("\r\n", "|")
    s = s.replace("\n", "|")
    s = s.replace("<", "|")
    s = s.replace(">", "|")
    s = re.sub(r"\|+", "|", s)
    return [part.strip() for part in s.split("|") if part.strip()]


def _looks_like_file_path(token: str) -> bool:
    token = token.strip()
    if re.match(r"^[A-Za-z]:\\", token) or token.startswith("\\\\") or token.startswith("/"):
        return True
    return "\\" in token or "/" in token


def _normalize_numeric_pmid(token: str) -> Optional[str]:
    s = token.strip()
    if s.isdigit() and len(s) <= 9:
        return s.lstrip("0") or "0"
    return None


def ids_from_clipboard_token(token: str) -> List[str]:
    """
    Map one post-split clipboard token to zero or more ID strings.

    - PubMed article URL -> PMID(s)
    - File path or file name -> stem (extension removed)
    - All-digit token -> normalized PMID
    - Otherwise -> original token
    """
    token = token.strip()
    if not token:
        return []

    url_pmids = pmids_from_pubmed_urls_in_text(token)
    if url_pmids:
        return url_pmids

    if _looks_like_file_path(token):
        base = os.path.basename(token.rstrip("\\/"))
        if base:
            return [strip_last_extension(base)]
        return []

    pmid = _normalize_numeric_pmid(token)
    if pmid is not None:
        return [pmid]

    return [token]


def ids_from_clipboard_text(text: str) -> List[str]:
    """
    Parse ID(s) from clipboard plain text: delimiter split, then per-token processing.
    """
    ids: List[str] = []
    for token in split_clipboard_text(text):
        ids.extend(ids_from_clipboard_token(token))
    return list(set_preserve_order(ids))


def string_to_list(input: Optional[str]) -> Optional[List[str]]:
    if input is None:
        return None
    # tab and line break to delimiter
    input = input.replace("\t", "|")
    input = input.replace("\r\n", "|")
    input = input.replace("\n", "|")
    input = input.replace("<", "|")
    input = input.replace(">", "|")
    input = input.replace(":", "|")
    # 중복되는 | 제거
    to_remove = "|"
    pattern = "(?P<char>[" + re.escape(to_remove) + "])(?P=char)+"
    input = re.sub(pattern, r"\1", input)
    input = str.split(input, sep="|")
    # list 공백값 제거
    input = [item for item in input if item.strip() != ""]
    # apply strip()
    input = [i.strip() for i in input]
    # Remove duplicates from the list
    input = list(set_preserve_order(input))
    return input


def list_to_string(list: Optional[List[str]], chr: int = 60) -> str:
    if list is None:
        return ""
    if not list:
        return ""
    listlength = len(list)
    longlist = False
    if listlength > 3:
        list = list[:3]
        longlist = True
    string = f"{listlength} Selection{'s' if listlength > 1 else ''}: " + "<" + "> <".join(list) + ">"
    if len(string) > chr and string.count(">") > 1:
        string = string[: string.rfind(">", 0, len(string) - 1) + 1]
        longlist = True
    if len(string) > chr and string.count(">") > 1:
        string = string[: string.rfind(">", 0, len(string) - 1) + 1]
    if len(string) > chr or longlist:
        string = string[:chr] + " ..."
    return string
