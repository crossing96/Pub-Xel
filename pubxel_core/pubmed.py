# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

import datetime
import re
from typing import Any, Callable, Dict, List, Optional, Tuple, Union

import requests
import xlwings as xw
from bs4 import BeautifulSoup

from pubxel_core.ids import set_preserve_order
from pubxel_core.metadata_store import (
    ArticleRow,
    AuthorRow,
    MetadataDict,
    MetadataStore,
    enrich_metadata,
)
from pubxel_core.paths import journal_combined_path, metadata_path


def value_from_dict(
    dictionary: Dict[str, str],
    key: str,
    outputtype: str = "string",
    default: str = "",
) -> Union[str, List[str]]:
    if outputtype not in ["string", "first", "list"]:
        raise ValueError("Invalid outputtype. Expected 'string', 'first', or 'list'.")
    if key in dictionary:
        value = dictionary[key]
        if outputtype == "string":
            return value
        elif outputtype == "first":
            return value.split("|")[0]
        elif outputtype == "list":
            return value.split("|")
    return default


def normalize_pmid(pmid: Union[str, int]) -> str:
    """Normalize a single PMID/accession for SQLite and metadata dict keys."""
    s = str(pmid).strip()
    if not s.isnumeric() or len(s) > 9:
        return s
    return s.lstrip("0") or "0"


def normalize_pmid_list(PMID_list: Union[str, List[str]]) -> List[str]:
    """Normalize, dedupe, and filter a PMID list (same rules as obtain_pubmed_data)."""
    if not isinstance(PMID_list, list):
        PMID_list = [PMID_list]

    PMID_list = [str(PMID) for PMID in PMID_list]
    PMID_list = [PMID.lstrip("0") or "0" for PMID in PMID_list]
    PMID_list = list(set_preserve_order(PMID_list))
    PMID_list = [PMID for PMID in PMID_list if PMID.isnumeric()]
    PMID_list = [PMID for PMID in PMID_list if len(PMID) <= 9]
    return PMID_list


def _normalize_medline_segment(segment: str) -> str:
    return segment.replace("\r\n", "\n").replace("\n", "\r\n")


def _medline_to_dict(text: str) -> Dict[str, str]:
    text = text.replace(" \r\n      ", " ")
    text = text.replace("\r\n      ", " ")
    result: Dict[str, str] = {}
    for line in text.strip().split("\r\n"):
        if "- " in line:
            key, value = line.split("- ", 1)
            key = key.rstrip()
            value = value.rstrip()
            if key in result:
                result[key] += "|" + value
            else:
                result[key] = value
    return result


def _get_first_doi_value(values: Any) -> str:
    try:
        if not isinstance(values, list):
            return ""
        for v in values:
            if isinstance(v, str) and v.endswith(" [doi]"):
                return v[:-6]
        return ""
    except Exception:
        return ""


def _empty_to_none(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str) and value.strip() == "":
        return None
    return value


def _parse_authors_from_segment(segment: str) -> List[Dict[str, Any]]:
    authors: List[Dict[str, Any]] = []
    for raw_line in _normalize_medline_segment(segment).split("\r\n"):
        line = raw_line.rstrip("\n")
        if not line.strip():
            continue
        m = re.match(r"^(CN|FAU|AU)\s*-\s*(.*)$", line)
        if not m:
            continue
        tag, value = m.group(1), m.group(2).strip()
        if value == "":
            continue
        if tag == "CN":
            authors.append({"type": "CORPORATE", "full_name": value, "short_name": None})
        elif tag == "FAU":
            authors.append({"type": "PERSON", "full_name": value, "short_name": None})
        elif tag == "AU":
            attached = False
            for i in range(len(authors) - 1, -1, -1):
                if authors[i]["type"] == "PERSON" and authors[i]["short_name"] is None:
                    authors[i]["short_name"] = value
                    attached = True
                    break
            if not attached:
                authors.append({"type": "PERSON", "full_name": None, "short_name": value})
    return authors


def _rows_from_medline_segment(
    pmid: str,
    segment: str,
    retrievedate: str,
) -> Tuple[Optional[ArticleRow], List[AuthorRow]]:
    segment = _normalize_medline_segment(segment)
    pmid_dict = _medline_to_dict(segment)

    title = value_from_dict(pmid_dict, "TI", outputtype="string", default="")
    abstract = value_from_dict(pmid_dict, "AB", outputtype="string", default="")
    journal = value_from_dict(pmid_dict, "TA", outputtype="string", default="")
    fulljournal = value_from_dict(pmid_dict, "JT", outputtype="string", default="")
    source = value_from_dict(pmid_dict, "SO", outputtype="string", default="")
    date_str = value_from_dict(pmid_dict, "DP", outputtype="string", default="")
    language = value_from_dict(pmid_dict, "LA", outputtype="string", default="")
    publicationtype = value_from_dict(pmid_dict, "PT", outputtype="string", default="")
    issn = value_from_dict(pmid_dict, "IS", outputtype="string", default="")
    si = value_from_dict(pmid_dict, "SI", outputtype="string", default="")
    gr = value_from_dict(pmid_dict, "GR", outputtype="string", default="")
    cin = value_from_dict(pmid_dict, "CIN", outputtype="string", default="")
    volume = value_from_dict(pmid_dict, "VI", outputtype="string", default="")
    issue = value_from_dict(pmid_dict, "IP", outputtype="string", default="")
    page = value_from_dict(pmid_dict, "PG", outputtype="string", default="")
    provider = value_from_dict(pmid_dict, "OWN", outputtype="string", default="")

    year = None
    if date_str:
        m_year = re.search(r"\b(\d{4})\b", date_str)
        if m_year:
            year = m_year.group(1)

    aid_list = value_from_dict(pmid_dict, "AID", outputtype="list", default=[])
    doi = _get_first_doi_value(aid_list)

    accession = pmid
    title = _empty_to_none(title)
    abstract = _empty_to_none(abstract)
    journal = _empty_to_none(journal)
    fulljournal = _empty_to_none(fulljournal)
    source = _empty_to_none(source)
    date_str = _empty_to_none(date_str)
    year = _empty_to_none(year)
    language = _empty_to_none(language)
    publicationtype = _empty_to_none(publicationtype)
    issn = _empty_to_none(issn)
    si = _empty_to_none(si)
    gr = _empty_to_none(gr)
    cin = _empty_to_none(cin)
    volume = _empty_to_none(volume)
    issue = _empty_to_none(issue)
    page = _empty_to_none(page)
    provider = _empty_to_none(provider)
    doi = _empty_to_none(doi)
    link = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}" if pmid else None

    article_row: ArticleRow = (
        accession,
        pmid,
        title,
        abstract,
        journal,
        year,
        source,
        date_str,
        doi,
        link,
        volume,
        issue,
        page,
        language,
        publicationtype,
        fulljournal,
        issn,
        si,
        gr,
        cin,
        retrievedate,
        provider,
    )

    authors_rows: List[AuthorRow] = []
    authors = _parse_authors_from_segment(segment)
    for idx, author in enumerate(authors, start=1):
        authors_rows.append(
            (
                accession,
                idx,
                author.get("type"),
                _empty_to_none(author.get("full_name")),
                _empty_to_none(author.get("short_name")),
            )
        )

    return article_row, authors_rows


def _upsert_medline_segments(segments: List[Tuple[str, str]]) -> MetadataDict:
    retrievedate = datetime.date.today().isoformat()
    articles_rows: List[ArticleRow] = []
    authors_rows: List[AuthorRow] = []

    for pmid, segment in segments:
        article_row, author_rows = _rows_from_medline_segment(pmid, segment, retrievedate)
        if article_row:
            articles_rows.append(article_row)
            authors_rows.extend(author_rows)

    if not articles_rows and not authors_rows:
        return {}

    store = MetadataStore(metadata_path)
    try:
        store.upsert_metadata(articles_rows, authors_rows)
        accessions = [row[0] for row in articles_rows if row[0]]
        raw = store.get_metadata(accessions)
        return enrich_metadata(raw)
    finally:
        store.close()


def import_nbib_to_metadata(records: List[str]) -> MetadataDict:
    """Parse PubMed nbib MEDLINE records and upsert into SQLite (no network)."""
    segments: List[Tuple[str, str]] = []
    for record in records:
        segment = _normalize_medline_segment(record)
        pmid = None
        for line in segment.split("\r\n"):
            if line.startswith("PMID-"):
                pmid = normalize_pmid(line.split("-", 1)[1].strip())
                break
        if not pmid:
            raise ValueError("Can only import .nbib file from PubMed")
        segments.append((pmid, segment))
    return _upsert_medline_segments(segments)


def resolve_metadata_for_pmids(
    PMID_list: Union[str, List[str]],
    on_partial: Optional[Callable[[MetadataDict], None]] = None,
    on_fetch_start: Optional[Callable[[], None]] = None,
) -> MetadataDict:
    """
    Load metadata cache-first: SQLite for PMIDs already stored, PubMed only for missing.

    When ``on_partial`` is provided, it is invoked after loading cached rows (if any),
    and again after a PubMed fetch when any PMIDs were missing from SQLite.
    When ``on_fetch_start`` is provided, it is invoked only when a network fetch to
    PubMed is about to start (i.e., at least one PMID is missing from SQLite).
    """
    pmids = normalize_pmid_list(PMID_list)
    if not pmids:
        return {}

    store = MetadataStore(metadata_path)
    merged: MetadataDict = {}
    try:
        missing = store.missing(pmids)
        missing_set = set(missing)
        cached = [p for p in pmids if p not in missing_set]

        if cached:
            merged = enrich_metadata(store.get_metadata(cached))
            if on_partial:
                on_partial(merged)

        if missing:
            if on_fetch_start:
                on_fetch_start()
            fetched = obtain_pubmed_data(missing)
            merged.update(fetched)
            if on_partial:
                on_partial(merged)
    finally:
        store.close()

    return merged


def obtain_pubmed_data(PMID_list: Union[str, List[str]]) -> MetadataDict:
    """
    Fetch PubMed MEDLINE data via NCBI, persist to SQLite, and return enriched metadata.

    Fetches from the NCBI ctxp API, upserts into ``metadata_article.sqlite``,
    then returns ``enrich_metadata(store.get_metadata(...))`` for callers.
    """
    PMID_list = normalize_pmid_list(PMID_list)
    if not PMID_list:
        return {}

    if len(PMID_list) > 1:
        PMID_list_str = ",".join(PMID_list)
    else:
        PMID_list_str = PMID_list[0]

    url = "https://api.ncbi.nlm.nih.gov/lit/ctxp/v1/pubmed/?format=medline&id=" + PMID_list_str

    n_articles = len(PMID_list)
    preview = ", ".join(PMID_list[:5])
    if n_articles > 5:
        preview = f"{preview}, ..."
    print(f"Obtaining PubMed data for {n_articles} article(s): {preview}")

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
    except requests.exceptions.HTTPError as e:
        raise ValueError(f"HTTP Error.\nInvalid PMID(s). Please try again.\n{e}")
    except requests.exceptions.ConnectionError:
        raise ValueError("Connection Error.\nFailed to connect. Please check your internet connection.")
    except requests.exceptions.Timeout:
        raise ValueError("Timeout Error.\nFailed to connect (Timeout). Please try again later.")
    except requests.exceptions.RequestException as e:
        raise ValueError(f"Error.\nAn unexpected error occurred: {e}")

    html_doc = response.text
    soup = BeautifulSoup(html_doc, "html.parser")
    data = soup.get_text()

    if not data.startswith("PMID"):
        raise ValueError("Error.\nInvalid PMID(s). Please try again.")

    html_dict: Dict[str, str] = {}
    for segment in data.split("\r\n\r\n"):
        for line in segment.split("\r\n"):
            if line.startswith("PMID"):
                pmid_key = line.split("-", 1)[1].strip()
                html_dict[pmid_key] = segment
                break

    segments: List[Tuple[str, str]] = []
    for pmid in PMID_list:
        if pmid in html_dict:
            segments.append((pmid, html_dict[pmid]))

    return _upsert_medline_segments(segments)


# Worksheet column header (lowercase) -> internal field key
WORKSHEET_COLUMN_HEADER: Dict[str, str] = {
    "ref": "pmid",
    "doi": "doi",
    "funding": "gr",
    "grants": "gr",
    "identifier": "si",
    "secondaryid": "si",
    "firstauthor": "fa",
    "author": "au",
    "year": "yr",
    "authoryear": "fayr",
    "journal": "jo",
    "title": "ti",
    "abstract": "ab",
    "citation": "cite",
    "output2": "ou2",
    "authors": "au2",
    "if2022": "if2022",
    "citation2022": "cite2022",
    "q2022": "q2022",
    "if2023": "if2023",
    "citation2023": "cite2023",
    "q2023": "q2023",
    "if2024": "if2024",
    "citation2024": "cite2024",
    "q2024": "q2024",
    "if2025": "if2025",
    "citation2025": "cite2025",
    "q2025": "q2025",
}


def _sql_if_empty(value: Any) -> str:
    if value == "" or value is None:
        return "-"
    return value


def _journal_if_empty(value: Any) -> str:
    if value == "" or value is None:
        return "-"
    return value


def parse_pmid_cell_value(value: Any) -> Optional[str]:
    """Normalize a worksheet cell value to a PMID string, or None if invalid."""
    if isinstance(value, int) and value > 0:
        return str(value)
    if isinstance(value, float) and value > 0 and value.is_integer():
        return str(int(value))
    if isinstance(value, str):
        s = value.strip()
        if s.replace(".", "", 1).isdigit():
            f = float(s)
            if f > 0 and f.is_integer():
                return str(int(f))
    return None


def build_worksheet_header_map(header_names: List[Any]) -> Dict[str, int]:
    """Map internal field keys to 0-based column indices from row-1 header labels."""
    header2: Dict[str, int] = {}
    for col_idx, name in enumerate(header_names):
        if name is None:
            continue
        key = WORKSHEET_COLUMN_HEADER.get(str(name).lower())
        if key is not None:
            header2[key] = col_idx
    return header2


def load_impact_factor_dict() -> Dict[str, Tuple[str, ...]]:
    impactfactordict: Dict[str, Tuple[str, ...]] = {}
    with open(journal_combined_path, "r", encoding="utf8") as file:
        lines = file.readlines()
        if not lines:
            return impactfactordict
        tsv_header = [h.strip() for h in lines[0].rstrip("\n").split("\t")]
        hidx = {name: i for i, name in enumerate(tsv_header)}
        for line in lines[1:]:
            parts = line.rstrip("\n").split("\t")
            if "abb" not in hidx or (
                "IF_2022" not in hidx
                and "q_2022" not in hidx
                and "IF_2023" not in hidx
                and "q_2023" not in hidx
                and "IF_2024" not in hidx
                and "q_2024" not in hidx
                and "IF_2025" not in hidx
                and "q_2025" not in hidx
            ):
                continue
            max_idx = max(
                hidx.get("abb", -1),
                hidx.get("IF_2022", -1),
                hidx.get("q_2022", -1),
                hidx.get("IF_2023", -1),
                hidx.get("q_2023", -1),
                hidx.get("IF_2024", -1),
                hidx.get("q_2024", -1),
                hidx.get("IF_2025", -1),
                hidx.get("q_2025", -1),
            )
            if len(parts) <= max_idx:
                continue
            abb = parts[hidx["abb"]].strip()
            if not abb:
                continue
            IF2022 = parts[hidx["IF_2022"]].strip() if "IF_2022" in hidx else ""
            quartile2022 = parts[hidx["q_2022"]].strip() if "q_2022" in hidx else ""
            IF2023 = parts[hidx["IF_2023"]].strip() if "IF_2023" in hidx else ""
            quartile2023 = parts[hidx["q_2023"]].strip() if "q_2023" in hidx else ""
            IF2024 = parts[hidx["IF_2024"]].strip() if "IF_2024" in hidx else ""
            quartile2024 = parts[hidx["q_2024"]].strip() if "q_2024" in hidx else ""
            IF2025 = parts[hidx["IF_2025"]].strip() if "IF_2025" in hidx else ""
            quartile2025 = parts[hidx["q_2025"]].strip() if "q_2025" in hidx else ""
            if (
                IF2022
                or quartile2022
                or IF2023
                or quartile2023
                or IF2024
                or quartile2024
                or IF2025
                or quartile2025
            ):
                impactfactordict[abb] = (
                    IF2022,
                    quartile2022,
                    IF2023,
                    quartile2023,
                    IF2024,
                    quartile2024,
                    IF2025,
                    quartile2025,
                )
    return impactfactordict


def _needs_impact_factors(header2: Dict[str, int]) -> bool:
    for year in ("2022", "2023", "2024", "2025"):
        if (
            header2.get(f"if{year}", -1) >= 0
            or header2.get(f"cite{year}", -1) >= 0
            or header2.get(f"q{year}", -1) >= 0
        ):
            return True
    return False


def _format_worksheet_cell(key: str, value: Any) -> Any:
    if key.startswith("if") or key.startswith("q"):
        return _journal_if_empty(value)
    if key == "pmid":
        return value
    return _sql_if_empty(value)


def build_worksheet_row_values(
    pmid: str,
    metadata: MetadataDict,
    header2: Dict[str, int],
    impactfactordict: Dict[str, Tuple[str, ...]],
) -> Dict[str, Any]:
    """Compute internal field values for one PMID row (formatted for worksheet cells)."""
    article_dict = metadata[pmid].get("article", {}) or {}

    journal = article_dict.get("journal", "")
    title = article_dict.get("title", "")
    source = article_dict.get("source", "") or ""
    firstauthorlastnameetal = article_dict.get("firstauthorlastnameetal", "")
    cite = (
        firstauthorlastnameetal.rstrip(".")
        + ". "
        + title
        + " "
        + source.rstrip()
        + " PMID: "
        + pmid
        + "."
    )
    authors_list = metadata[pmid].get("authors", []) or []
    authors_text = ", ".join(
        a.get("short_name") or a.get("full_name") or ""
        for a in authors_list
        if (a.get("short_name") or a.get("full_name"))
    )

    IF2022 = IF2023 = IF2024 = IF2025 = ""
    Q2022 = Q2023 = Q2024 = Q2025 = ""
    cite2022 = cite2023 = cite2024 = cite2025 = ""
    need_any = _needs_impact_factors(header2)
    if need_any:
        jkey = (journal or "").upper().strip().rstrip(".")
        IF2022, Q2022, IF2023, Q2023, IF2024, Q2024, IF2025, Q2025 = impactfactordict.get(
            jkey, ("", "", "", "", "", "", "", "")
        )

        if header2.get("cite2022", -1) >= 0:
            if IF2022:
                pattern = re.escape(journal)
                replacement = r"\1 (IF: " + IF2022 + ")"
                source2022 = re.sub(f"({pattern})", replacement, source, count=1)
            else:
                source2022 = source
            cite2022 = (
                firstauthorlastnameetal.rstrip(".")
                + ". "
                + title
                + " "
                + source2022.rstrip()
                + " PMID: "
                + pmid
                + "."
            )

        if header2.get("cite2023", -1) >= 0:
            if IF2023:
                pattern = re.escape(journal)
                replacement = r"\1 (IF: " + IF2023 + ")"
                source2023 = re.sub(f"({pattern})", replacement, source, count=1)
            else:
                source2023 = source
            cite2023 = (
                firstauthorlastnameetal.rstrip(".")
                + ". "
                + title
                + " "
                + source2023.rstrip()
                + " PMID: "
                + pmid
                + "."
            )

        if header2.get("cite2024", -1) >= 0:
            if IF2024:
                pattern = re.escape(journal)
                replacement = r"\1 (IF: " + IF2024 + ")"
                source2024 = re.sub(f"({pattern})", replacement, source, count=1)
            else:
                source2024 = source
            cite2024 = (
                firstauthorlastnameetal.rstrip(".")
                + ". "
                + title
                + " "
                + source2024.rstrip()
                + " PMID: "
                + pmid
                + "."
            )

        if header2.get("cite2025", -1) >= 0:
            if IF2025:
                pattern = re.escape(journal)
                replacement = r"\1 (IF: " + IF2025 + ")"
                source2025 = re.sub(f"({pattern})", replacement, source, count=1)
            else:
                source2025 = source
            cite2025 = (
                firstauthorlastnameetal.rstrip(".")
                + ". "
                + title
                + " "
                + source2025.rstrip()
                + " PMID: "
                + pmid
                + "."
            )

    raw: Dict[str, Any] = {
        "pmid": pmid,
        "doi": article_dict.get("doi", ""),
        "gr": (article_dict.get("gr", "") or "").replace("|", "; "),
        "si": (article_dict.get("si", "") or "").replace("|", "; "),
        "au2": authors_text,
        "au": authors_text,
        "fa": firstauthorlastnameetal,
        "ti": title,
        "ab": article_dict.get("abstract", ""),
        "jo": article_dict.get("journal", ""),
        "yr": article_dict.get("year", ""),
        "fayr": article_dict.get("authoryear", ""),
        "cite": cite,
        "if2022": IF2022,
        "cite2022": cite2022,
        "q2022": Q2022,
        "if2023": IF2023,
        "cite2023": cite2023,
        "q2023": Q2023,
        "if2024": IF2024,
        "cite2024": cite2024,
        "q2024": Q2024,
        "if2025": IF2025,
        "cite2025": cite2025,
        "q2025": Q2025,
    }

    return {key: _format_worksheet_cell(key, raw.get(key, "")) for key in header2}


def iter_worksheet_export_rows(
    pmids: List[str],
    metadata: MetadataDict,
    column_names: List[str],
) -> Tuple[List[str], List[List[str]]]:
    """Build header and data rows for worksheet export (Excel TSV, etc.)."""
    header2 = build_worksheet_header_map(column_names)
    need_any = _needs_impact_factors(header2)
    impactfactordict = load_impact_factor_dict() if need_any else {}

    data_rows: List[List[str]] = []
    for pmid in pmids:
        if pmid not in metadata:
            row_values = {"pmid": pmid}
        else:
            row_values = build_worksheet_row_values(pmid, metadata, header2, impactfactordict)

        cells: List[str] = []
        for col_name in column_names:
            key = WORKSHEET_COLUMN_HEADER.get(str(col_name).lower())
            if key is None:
                cells.append("")
            else:
                val = row_values.get(key, "")
                cells.append("" if val is None else str(val))
        data_rows.append(cells)

    return column_names, data_rows


def fill_worksheet_rows(
    ws: xw.Sheet,
    metadata: MetadataDict,
    header2: Dict[str, int],
    rows: List[Tuple[int, str]],
) -> Tuple[List[str], List[str]]:
    """Fill worksheet data rows. Each item in ``rows`` is (0-based row index, PMID)."""
    identified_pmids: List[str] = []
    unidentified_pmids: List[str] = []
    need_any = _needs_impact_factors(header2)
    impactfactordict = load_impact_factor_dict() if need_any else {}

    for row_i, pmidstring in rows:
        if pmidstring not in metadata:
            unidentified_pmids.append(pmidstring)
            if header2.get("pmid", -1) >= 0:
                ws[row_i, header2["pmid"]].value = pmidstring
            continue

        identified_pmids.append(pmidstring)
        row_values = build_worksheet_row_values(
            pmidstring, metadata, header2, impactfactordict
        )
        for key, col_idx in header2.items():
            if key in row_values:
                ws[row_i, col_idx].value = row_values[key]

    return identified_pmids, unidentified_pmids


def input_pubmed_data() -> Optional[str]:
    header2: Dict[str, int] = {}
    try:
        app = xw.apps.active
        wb = xw.books.active
        ws = xw.sheets.active
    except Exception:
        raise ValueError("Please open the Excel Worksheet first.")

    try:
        rng = wb.app.selection
    except Exception:
        raise ValueError("No selection made. Please make a selection in the Excel sheet.")

    # rng = trim_range(rng,reselect=True): this action is already done in check_file_exist

    if rng is None:
        raise ValueError("No selection made.. Please make a selection in the Excel sheet.")

    if rng.count > 200:
        raise ValueError("Please select 200 or fewer cells!")

    ###Extend the selection to the entire row
    ###identify relevant first row selection, and expand rng
    # Step 1: Identify column interval of the selection
    column_interval = (rng.column, rng.column + rng.columns.count - 1)
    # Step 2: Select a range of the first row of the sheet, ranging from the column range of the selection
    header_range = ws.range((1, column_interval[0]), (1, column_interval[1]))
    # Step 3: Extend the range to include the entire row
    header_range = header_range.expand("right")

    # Check if any range is selected for header_range
    if header_range.count == 0:
        header_range.select()
        raise ValueError(
            "Error: Please ensure the following before trying again.:\n"
            "1. The table header must be located in the first row of the entire Excel sheet (Row 1).\n"
            "2. The column header containing PMIDs is labeled as 'Ref'."
        )

    # Check if "ref" is present in the header_range
    ref_count = sum(1 for cell in header_range if cell.value is not None and cell.value.lower() == "ref")
    print("ref_count: ", ref_count)
    if ref_count == 0:
        header_range.select()
        raise ValueError(
            "Error: Please ensure the following before trying again..:\n"
            "1. The table header must be located in the first row of the entire Excel sheet (Row 1).\n"
            "2. The column header containing PMIDs is labeled as 'Ref'."
        )
    elif ref_count > 1:
        header_range.select()
        raise ValueError(
            "Error: Multiple 'ref' columns found. Please ensure there is only one 'ref' column in the table header."
        )

    # Step 4: Return the column interval of the header_range
    new_column_interval = (header_range.column, header_range.column + header_range.columns.count - 1)

    # Step 5: Extend the initial selection to include the columns identified in new_column_interval, then select the header_range
    extended_selection = ws.range((rng.row, new_column_interval[0]), (rng.row + rng.rows.count - 1, new_column_interval[1]))
    extended_selection.select()
    rng = wb.app.selection

    # identify existent headers
    for i in list(rng[0, :]):
        if ws[0, i.column - 1].value is None:
            continue
        if ws[0, i.column - 1].value.lower() in WORKSHEET_COLUMN_HEADER:
            header2[WORKSHEET_COLUMN_HEADER.get(ws[0, i.column - 1].value.lower())] = i.column - 1

    rng_col_range = range(rng[:, 0].row - 1, rng[:, 0].row - 1 + len(rng[:, 0]))
    requested_ref_count = sum(1 for number in rng_col_range if number != 0)

    PMID_list: List[str] = []
    nonPMID_list: List[Any] = []

    print(header2)
    PMIDs = [ws[i, header2.get("pmid")].value for i in rng_col_range]

    for PMID in PMIDs:
        pmidstring = parse_pmid_cell_value(PMID)
        if pmidstring:
            PMID_list.append(pmidstring)
        else:
            nonPMID_list.append(PMID)

    metadata = obtain_pubmed_data(PMID_list)

    if not isinstance(metadata, dict):
        return None

    rows_to_fill: List[Tuple[int, str]] = []
    for i in rng_col_range:
        if i == 0:
            continue
        pmid_col = header2.get("pmid")
        if pmid_col is None:
            continue
        pmidstring = parse_pmid_cell_value(ws[i, pmid_col].value)
        if pmidstring:
            rows_to_fill.append((i, pmidstring))

    identifiedPMID_list, unidentifiedPMID_list = fill_worksheet_rows(
        ws, metadata, header2, rows_to_fill
    )

    if True:
        print("Number of requested references: " + str(requested_ref_count))
        print("non-PMID references: " + str(nonPMID_list))
        print("PMIDs: " + str(PMID_list))
        print("identified PMIDs: " + str(identifiedPMID_list))
        print("Unidentified PMIDs: " + str(unidentifiedPMID_list))

    from pubxel_core.recent_worksheets import try_register_active_workbook

    try_register_active_workbook()
    return "Import successful"
