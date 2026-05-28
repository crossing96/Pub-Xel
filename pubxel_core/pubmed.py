# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

import datetime
import os
import re
import time
from typing import Any, Callable, Dict, List, Optional, Union

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
from pubxel_core.paths import data_dir, metadata_path


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

    Returns
    -------
    MetadataDict
        ``{ pmid: { "article": {...}, "authors": [...] } }`` with citation fields
        added by ``enrich_metadata`` (cite, authoryear, etc.).
    """

    PMID_list = normalize_pmid_list(PMID_list)
    if not PMID_list:
        return {}

    def html_to_dict(text):
        """
        Convert MEDLINE-like text block into a dict of key -> '|' joined values.
        Non-author / non-CN fields can be safely read from here.

        This is not used for author order, only for quick value lookup.
        """
        text = text.replace(" \r\n      ", " ")
        text = text.replace("\r\n      ", " ")
        result = {}
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

    def get_first_doi_value(values):
        """
        Returns the first DOI string (without ' [doi]') from a list.
        If no valid DOI is found or input is invalid, returns an empty string.
        """
        try:
            if not isinstance(values, list):
                return ""
            for v in values:
                if isinstance(v, str) and v.endswith(" [doi]"):
                    return v[:-6]  # strip " [doi]"
            return ""
        except Exception:
            return ""

    def empty_to_none(value):
        """
        Convert empty string or whitespace-only string to None.
        Anything else is returned as-is.
        """
        if value is None:
            return None
        if isinstance(value, str) and value.strip() == "":
            return None
        return value

    def parse_authors_from_segment(segment):
        """
        Parse CN / FAU / AU in correct MEDLINE order from a raw segment string.

        Returns
        -------
        authors : list of dicts
            Each dict:
              {
                "type": "PERSON" or "CORPORATE",
                "full_name": <str or None>,
                "short_name": <str or None>
              }

        Rules / quirks handled:
        - CN can appear anywhere and multiple times.
        - FAU and AU may be missing or mismatched.
        - AU without preceding FAU produces an author with short_name only.
        - Titles / abstracts etc are ignored here; only CN/FAU/AU processed.
        - Papers with only CN and no human authors are supported.
        - AU may represent corporate names in disguise; we do not try to guess,
          we just treat AU as PERSON unless paired with CN.
        """
        authors = []

        # Break by line, but keep as-is to preserve order
        for raw_line in segment.split("\r\n"):
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
                # Corporate author / collaboration group
                authors.append({
                    "type": "CORPORATE",
                    "full_name": value,
                    "short_name": None,
                })

            elif tag == "FAU":
                # Full author name (person)
                authors.append({
                    "type": "PERSON",
                    "full_name": value,
                    "short_name": None,
                })

            elif tag == "AU":
                # Short author name, should attach to last PERSON without short_name
                attached = False
                for i in range(len(authors) - 1, -1, -1):
                    if authors[i]["type"] == "PERSON" and authors[i]["short_name"] is None:
                        authors[i]["short_name"] = value
                        attached = True
                        break
                if not attached:
                    # AU without FAU before it → treat as standalone PERSON with short_name only
                    authors.append({
                        "type": "PERSON",
                        "full_name": None,
                        "short_name": value,
                    })

        return authors

    # Build ID string for NCBI ctxp API
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

    # DEBUG: artificial delay before network request (remove when done debugging)
    time.sleep(3)

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

    # Map each PMID to its raw MEDLINE segment
    html_dict = {}
    segments = data.split("\r\n\r\n")
    for segment in segments:
        lines = segment.split("\r\n")
        for line in lines:
            if line.startswith("PMID"):
                pmid_key = line.split("-", 1)[1].strip()
                html_dict[pmid_key] = segment
                break

    articles_rows = []
    authors_rows = []

    retrievedate = datetime.date.today().isoformat()  # YYYY-MM-DD

    # Build SQL-input-friendly rows
    for PMID in PMID_list:
        if PMID not in html_dict:
            continue

        segment = html_dict[PMID]
        PMID_dict = html_to_dict(segment)

        # Core bibliographic fields
        title = value_from_dict(PMID_dict, "TI", outputtype="string", default="")
        abstract = value_from_dict(PMID_dict, "AB", outputtype="string", default="")
        journal = value_from_dict(PMID_dict, "TA", outputtype="string", default="")
        fulljournal = value_from_dict(PMID_dict, "JT", outputtype="string", default="")
        source = value_from_dict(PMID_dict, "SO", outputtype="string", default="")
        date_str = value_from_dict(PMID_dict, "DP", outputtype="string", default="")
        language = value_from_dict(PMID_dict, "LA", outputtype="string", default="")
        publicationtype = value_from_dict(PMID_dict, "PT", outputtype="string", default="")
        issn = value_from_dict(PMID_dict, "IS", outputtype="string", default="")
        si = value_from_dict(PMID_dict, "SI", outputtype="string", default="")
        gr = value_from_dict(PMID_dict, "GR", outputtype="string", default="")
        cin = value_from_dict(PMID_dict, "CIN", outputtype="string", default="")
        volume = value_from_dict(PMID_dict, "VI", outputtype="string", default="")
        issue = value_from_dict(PMID_dict, "IP", outputtype="string", default="")
        page = value_from_dict(PMID_dict, "PG", outputtype="string", default="")
        provider = value_from_dict(PMID_dict, "OWN", outputtype="string", default="")

        # Year: pull first 4-digit number from DP; if none, store None
        year = None
        if date_str:
            m_year = re.search(r"\b(\d{4})\b", date_str)
            if m_year:
                year = m_year.group(1)

        # DOI: use AID as list and pick first [doi]
        aid_list = value_from_dict(PMID_dict, "AID", outputtype="list", default=[])
        doi = get_first_doi_value(aid_list)

        # Normalize empties to None for SQL NULLs
        accession = PMID  # accession = PMID string, primary key in SQL
        pmid = PMID  # same as accession, non-PK column
        title = empty_to_none(title)
        abstract = empty_to_none(abstract)
        journal = empty_to_none(journal)
        fulljournal = empty_to_none(fulljournal)
        source = empty_to_none(source)
        date_str = empty_to_none(date_str)
        year = empty_to_none(year)
        language = empty_to_none(language)
        publicationtype = empty_to_none(publicationtype)
        issn = empty_to_none(issn)
        si = empty_to_none(si)
        gr = empty_to_none(gr)
        cin = empty_to_none(cin)
        volume = empty_to_none(volume)
        issue = empty_to_none(issue)
        page = empty_to_none(page)
        provider = empty_to_none(provider)
        doi = empty_to_none(doi)

        # Link is always derived from PMID string
        link = f"https://pubmed.ncbi.nlm.nih.gov/{PMID}" if PMID else None

        # Main article row (for main table)
        article_row = (
            accession,  # accession (TEXT, PRIMARY KEY)
            pmid,  # pmid (TEXT)
            title,  # title (TEXT)
            abstract,  # abstract (TEXT)
            journal,  # journal (TEXT)
            year,  # year (TEXT)
            source,  # source (TEXT, 'SO')
            date_str,  # date (TEXT, 'DP')
            doi,  # doi (TEXT)
            link,  # link (TEXT)
            volume,  # volume (TEXT)
            issue,  # issue (TEXT)
            page,  # page (TEXT)
            language,  # language (TEXT, 'LA')
            publicationtype,  # publicationtype (TEXT, 'PT')
            fulljournal,  # fulljournal (TEXT, 'JT')
            issn,  # ISSN (TEXT, 'IS')
            si,  # SI (TEXT, 'SI')
            gr,  # GR (TEXT, 'GR')
            cin,  # CIN (TEXT, 'CIN')
            retrievedate,  # retrievedate (TEXT, YYYY-MM-DD)
            provider,  # provider (TEXT, 'OWN')
        )
        articles_rows.append(article_row)

        # Author / CN rows (for author table), preserving MEDLINE order
        authors = parse_authors_from_segment(segment)

        # If there are no authors at all (no CN, no FAU/AU), nothing to add
        for idx, author in enumerate(authors, start=1):
            author_type = author.get("type")
            full_name = empty_to_none(author.get("full_name"))
            short_name = empty_to_none(author.get("short_name"))
            authors_rows.append(
                (
                    accession,  # accession (TEXT) as foreign key to main table
                    idx,  # author_order (INTEGER, 1-based)
                    author_type,  # author_type (TEXT: 'PERSON' or 'CORPORATE')
                    full_name,  # full_name (TEXT or NULL)
                    short_name,  # short_name (TEXT or NULL)
                )
            )

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


def input_pubmed_data() -> Optional[str]:
    # settings
    # header {column name : requested varaible}. column name all lower case
    header = {
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
    }

    header2 = {}
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

    def NAifempty(value):
        if value == "" or value is None:
            return "NA"
        else:
            return value

    # identify existent headers
    for i in list(rng[0, :]):
        if ws[0, i.column - 1].value is None:
            continue
        if ws[0, i.column - 1].value.lower() in header:
            header2[header.get(ws[0, i.column - 1].value.lower())] = i.column - 1

    rng_col_range = range(rng[:, 0].row - 1, rng[:, 0].row - 1 + len(rng[:, 0]))
    requested_ref_count = sum(1 for number in rng_col_range if number != 0)  # length of rng_col_range but minus 1 if contain 0.

    impactfactordict = {}

    need_2022 = header2.get("if2022", -1) >= 0 or header2.get("cite2022", -1) >= 0 or header2.get("q2022", -1) >= 0
    need_2023 = header2.get("if2023", -1) >= 0 or header2.get("cite2023", -1) >= 0 or header2.get("q2023", -1) >= 0
    need_2024 = header2.get("if2024", -1) >= 0 or header2.get("cite2024", -1) >= 0 or header2.get("q2024", -1) >= 0
    need_any = need_2022 or need_2023 or need_2024

    if need_any:
        IF_COMBINED_PATH = os.path.join(data_dir, "journal_combined_2.txt")
        print("load impactfactor from combined dataset")
        with open(IF_COMBINED_PATH, "r", encoding="utf8") as file:
            lines = file.readlines()
            if not lines:
                pass  # no data
            else:
                tsv_header = [h.strip() for h in lines[0].rstrip("\n").split("\t")]
                hidx = {name: i for i, name in enumerate(tsv_header)}
                for line in lines[1:]:  # Skip the header line
                    parts = line.rstrip("\n").split("\t")
                    # ensure required columns exist
                    if "abb" not in hidx or (
                        "IF_2022" not in hidx
                        and "q_2022" not in hidx
                        and "IF_2023" not in hidx
                        and "q_2023" not in hidx
                        and "IF_2024" not in hidx
                        and "q_2024" not in hidx
                    ):
                        continue
                    # ensure row long enough to access indexes we need
                    max_idx = max(
                        hidx.get("abb", -1),
                        hidx.get("IF_2022", -1),
                        hidx.get("q_2022", -1),
                        hidx.get("IF_2023", -1),
                        hidx.get("q_2023", -1),
                        hidx.get("IF_2024", -1),
                        hidx.get("q_2024", -1),
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
                    if IF2022 or quartile2022 or IF2023 or quartile2023 or IF2024 or quartile2024:
                        impactfactordict[abb] = (IF2022, quartile2022, IF2023, quartile2023, IF2024, quartile2024)

    PMID_list = []
    identifiedPMID_list = []
    unidentifiedPMID_list = []
    nonPMID_list = []

    print(header2)
    # Get all PMIDs possible by i's and make them to a list
    PMIDs = [ws[i, header2.get("pmid")].value for i in rng_col_range]

    for PMID in PMIDs:
        if isinstance(PMID, (int, float)) and PMID > 0 and PMID.is_integer():
            PMIDstring = str(int(PMID))
            PMID_list.append(str(int(PMIDstring)))
        elif isinstance(PMID, str) and PMID.replace(".", "", 1).isdigit() and float(PMID).is_integer() and float(PMID) > 0:
            PMIDstring = str(int(float(PMID)))
            PMID_list.append(str(int(PMIDstring)))
        else:
            nonPMID_list.append(PMID)

    metadata = obtain_pubmed_data(PMID_list)

    if not isinstance(metadata, dict):
        return

    # input values
    for i in rng_col_range:

        if i == 0:
            continue

        PMID = ws[i, header2.get("pmid")].value

        # Normalize PMID to a positive integer string
        PMIDstring = None
        if isinstance(PMID, int) and PMID > 0:
            PMIDstring = str(PMID)

        elif isinstance(PMID, float) and PMID > 0 and PMID.is_integer():
            PMIDstring = str(int(PMID))

        elif isinstance(PMID, str):
            s = PMID.strip()
            if s.replace(".", "", 1).isdigit():
                f = float(s)
                if f > 0 and f.is_integer():
                    PMIDstring = str(int(f))

        if not PMIDstring:
            continue

        # Now PMIDstring is a clean pmid string like "31452104"
        if PMIDstring not in metadata:
            unidentifiedPMID_list.append(PMIDstring)
            continue

        identifiedPMID_list.append(PMIDstring)

        article_dict = metadata[PMIDstring].get("article", {}) or {}

        journal = article_dict.get("journal", "")
        title = article_dict.get("title", "")
        source = article_dict.get("source", "") or ""
        firstauthorlastnameetal = article_dict.get("firstauthorlastnameetal", "")
        cite = firstauthorlastnameetal.rstrip(".") + ". " + title + " " + source.rstrip() + " PMID: " + PMIDstring + "."
        authors_list = metadata[PMIDstring].get("authors", []) or []
        authors_text = ", ".join(
            a.get("short_name") or a.get("full_name") or ""
            for a in authors_list
            if (a.get("short_name") or a.get("full_name"))
        )

        if need_any:
            jkey = (journal or "").upper().strip().rstrip(".")
            IF2022, Q2022, IF2023, Q2023, IF2024, Q2024 = impactfactordict.get(jkey, ("", "", "", "", "", ""))

            if header2.get("cite2022", -1) >= 0:
                if IF2022:
                    pattern = re.escape(journal)
                    replacement = r"\1 (IF: " + IF2022 + ")"
                    source2022 = re.sub(f"({pattern})", replacement, source, count=1)
                else:
                    source2022 = source
                cite2022 = firstauthorlastnameetal.rstrip(".") + ". " + title + " " + source2022.rstrip() + " PMID: " + PMIDstring + "."

            if header2.get("cite2023", -1) >= 0:
                if IF2023:
                    pattern = re.escape(journal)
                    replacement = r"\1 (IF: " + IF2023 + ")"
                    source2023 = re.sub(f"({pattern})", replacement, source, count=1)
                else:
                    source2023 = source
                cite2023 = firstauthorlastnameetal.rstrip(".") + ". " + title + " " + source2023.rstrip() + " PMID: " + PMIDstring + "."

            if header2.get("cite2024", -1) >= 0:
                if IF2024:
                    pattern = re.escape(journal)
                    replacement = r"\1 (IF: " + IF2024 + ")"
                    source2024 = re.sub(f"({pattern})", replacement, source, count=1)
                else:
                    source2024 = source
                cite2024 = firstauthorlastnameetal.rstrip(".") + ". " + title + " " + source2024.rstrip() + " PMID: " + PMIDstring + "."

        if header2.get("doi", -1) >= 0:
            ws[i, header2.get("doi")].value = article_dict.get("doi", "")
        if header2.get("gr", -1) >= 0:
            ws[i, header2.get("gr")].value = (article_dict.get("gr", "") or "").replace("|", "; ")
        if header2.get("si", -1) >= 0:
            ws[i, header2.get("si")].value = (article_dict.get("si", "") or "").replace("|", "; ")
        if header2.get("au2", -1) >= 0:
            ws[i, header2.get("au2")].value = NAifempty(authors_text)
        if header2.get("au", -1) >= 0:
            ws[i, header2.get("au")].value = NAifempty(authors_text)
        if header2.get("fa", -1) >= 0:
            ws[i, header2.get("fa")].value = NAifempty(firstauthorlastnameetal)
        if header2.get("ti", -1) >= 0:
            ws[i, header2.get("ti")].value = NAifempty(title)
        if header2.get("ab", -1) >= 0:
            ws[i, header2.get("ab")].value = article_dict.get("abstract", "")
        if header2.get("jo", -1) >= 0:
            ws[i, header2.get("jo")].value = article_dict.get("journal", "")
        if header2.get("yr", -1) >= 0:
            ws[i, header2.get("yr")].value = article_dict.get("year", "")
        if header2.get("fayr", -1) >= 0:
            ws[i, header2.get("fayr")].value = article_dict.get("authoryear", "")
        if header2.get("cite", -1) >= 0:
            ws[i, header2.get("cite")].value = NAifempty(cite)
        if header2.get("if2022", -1) >= 0:
            ws[i, header2.get("if2022")].value = NAifempty(IF2022)
        if header2.get("cite2022", -1) >= 0:
            ws[i, header2.get("cite2022")].value = NAifempty(cite2022)
        if header2.get("q2022", -1) >= 0:
            ws[i, header2.get("q2022")].value = NAifempty(Q2022)
        if header2.get("if2023", -1) >= 0:
            ws[i, header2.get("if2023")].value = NAifempty(IF2023)
        if header2.get("cite2023", -1) >= 0:
            ws[i, header2.get("cite2023")].value = NAifempty(cite2023)
        if header2.get("q2023", -1) >= 0:
            ws[i, header2.get("q2023")].value = NAifempty(Q2023)
        if header2.get("if2024", -1) >= 0:
            ws[i, header2.get("if2024")].value = NAifempty(IF2024)
        if header2.get("cite2024", -1) >= 0:
            ws[i, header2.get("cite2024")].value = NAifempty(cite2024)
        if header2.get("q2024", -1) >= 0:
            ws[i, header2.get("q2024")].value = NAifempty(Q2024)

    # # Enable screen updating
    # app.api.ScreenUpdating = True

    if True:
        print("Number of requested references: " + str(requested_ref_count))
        print("non-PMID references: " + str(nonPMID_list))
        print("PMIDs: " + str(PMID_list))
        print("identified PMIDs: " + str(identifiedPMID_list))
        print("Unidentified PMIDs: " + str(unidentifiedPMID_list))

    from pubxel_core.recent_worksheets import try_register_active_workbook

    try_register_active_workbook()
    return "Import successful"
