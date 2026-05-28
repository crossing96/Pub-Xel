# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

import os
import re

import xlwings as xw

from pubxel_core.clipboard import write_clipboard
from pubxel_core.ids import set_preserve_order

# Search location priority (lower number = preferred when the same basename appears
# in more than one place): main library root, main library subfolders, secondary roots.
_SEARCH_PRIORITY_MAIN = 0
_SEARCH_PRIORITY_MAIN_SUB = 1
_SEARCH_PRIORITY_SECONDARY = 2


def _search_include_subfolders(include_subfolders=None) -> bool:
    if include_subfolders is not None:
        return bool(include_subfolders)
    from pubxel_core import runtime as rt

    return bool(rt.settings.get("mainlib_include_subfolders", 0))


def _norm_dir_path(path: str) -> str:
    return os.path.normcase(os.path.normpath(os.path.abspath(path)))


def _search_priority_for_directory(
    dir_path: str, mainlibdir: str, seclibdirs_norm: list[str]
) -> int | None:
    """Classify which search tier a directory belongs to, or None if out of scope."""
    d = _norm_dir_path(dir_path)
    main_norm = _norm_dir_path(mainlibdir)
    if d == main_norm:
        return _SEARCH_PRIORITY_MAIN
    if d.startswith(main_norm + os.sep):
        return _SEARCH_PRIORITY_MAIN_SUB
    for sec in seclibdirs_norm:
        if d == sec:
            return _SEARCH_PRIORITY_SECONDARY
    return None


def build_library_catalog(
    mainlibdir: str,
    seclibdir=None,
    include_subfolders=None,
) -> dict[str, tuple[int, str]]:
    """Build basename -> (priority, full_path) for all searchable library files.

    Search scope:
    - Main library folder (top level only)
    - Main library subfolders (recursive), when ``mainlib_include_subfolders`` is on
    - Secondary library folders (top level only; subfolders are not searched)

    When the same basename exists in multiple locations, the entry with the best
    (lowest) priority wins: main > main subfolder > secondary.
    """
    if seclibdir is None:
        seclibdir = []
    include_subfolders = _search_include_subfolders(include_subfolders)
    catalog: dict[str, tuple[int, str]] = {}
    seclibdirs = [d for d in seclibdir if os.path.isdir(d)]
    seclib_norm = [_norm_dir_path(d) for d in seclibdirs]
    main_norm = _norm_dir_path(mainlibdir) if mainlibdir and os.path.isdir(mainlibdir) else ""

    def consider(full_path: str, container_dir: str) -> None:
        if not os.path.isfile(full_path):
            return
        pri = _search_priority_for_directory(container_dir, mainlibdir, seclib_norm)
        if pri is None:
            return
        base = os.path.basename(full_path)
        existing = catalog.get(base)
        if existing is None or pri < existing[0]:
            catalog[base] = (pri, full_path)

    if main_norm and os.path.isdir(mainlibdir):
        for name in os.listdir(mainlibdir):
            consider(os.path.join(mainlibdir, name), mainlibdir)
        if include_subfolders:
            for root, _dirs, files in os.walk(mainlibdir):
                if _norm_dir_path(root) == main_norm:
                    continue
                for name in files:
                    consider(os.path.join(root, name), root)

    for sec in seclibdirs:
        for name in os.listdir(sec):
            consider(os.path.join(sec, name), sec)

    return catalog


def files_name_to_path(
    fileName,
    mainlibdir,
    seclibdir=None,
    include_subfolders=None,
    file_map=None,
):
    """Resolve basenames to full paths using the same search rules as ``process_ids``."""
    if not fileName:
        return []
    if file_map is None:
        if seclibdir is None:
            seclibdir = []
        catalog = build_library_catalog(mainlibdir, seclibdir, include_subfolders)
        file_map = {name: entry[1] for name, entry in catalog.items()}
    path_list = []
    for file in fileName:
        full_path = file_map.get(file)
        if full_path:
            path_list.append(full_path)
    return path_list


def _build_catalog_index(catalog_names: list[str]) -> tuple[dict[str, list[str]], dict[str, list[str]]]:
    """Build fast lookup indexes from catalog basenames."""
    numeric_candidates: dict[str, list[str]] = {}
    main_by_stem: dict[str, list[str]] = {}
    for filename in catalog_names:
        m_lead = re.match(r"^(\d+)", filename)
        if m_lead:
            numeric_candidates.setdefault(m_lead.group(1), []).append(filename)
        m_main = re.match(r"^([^\.]+)\.[^.]+$", filename)
        if m_main:
            main_by_stem.setdefault(m_main.group(1), []).append(filename)
    return numeric_candidates, main_by_stem


def process_ids(
    ids,
    maindir,
    seclibdir=None,
    include_subfolders=None,
    catalog=None,
    catalog_index=None,
):
    if seclibdir is None:
        seclibdir = []
    seclibdir = [d for d in seclibdir if os.path.isdir(d)]

    valid_ids = []
    pubmed_ids = []
    non_pubmed_valid_ids = []
    invalid_ids = []
    valid_ids_with_m_files = []
    valid_ids_without_m_files = []
    pubmed_ids_with_s_files = []
    pubmed_ids_without_s_files = []
    pubmed_ids_with_m_files = []
    pubmed_ids_without_m_files = []
    nonpubmed_ids_with_m_files = []
    nonpubmed_ids_without_m_files = []
    all_m_files = []
    all_s_files = []

    empty = (
        valid_ids,
        pubmed_ids,
        non_pubmed_valid_ids,
        invalid_ids,
        valid_ids_with_m_files,
        valid_ids_without_m_files,
        pubmed_ids_with_m_files,
        pubmed_ids_without_m_files,
        pubmed_ids_with_s_files,
        pubmed_ids_without_s_files,
        all_m_files,
        all_s_files,
        nonpubmed_ids_with_m_files,
        nonpubmed_ids_without_m_files,
        {},
    )

    if ids is None:
        return empty

    if isinstance(ids, str):
        ids = [ids]
    ids = list(set_preserve_order(ids))

    if catalog is None:
        catalog = build_library_catalog(maindir, seclibdir, include_subfolders)
    catalog_names = list(catalog.keys())
    file_map = {name: entry[1] for name, entry in catalog.items()}
    if catalog_index is None:
        catalog_index = _build_catalog_index(catalog_names)
    numeric_candidates, main_by_stem = catalog_index

    for id in ids:
        if re.match(r"^[0-9]+$", id):
            valid_ids.append(id)
            pubmed_ids.append(id)
            m_files = []
            s_files = []
            id_main_re = re.compile(rf"^{id}\.[^.]+$")
            for filename in numeric_candidates.get(id, []):
                if id_main_re.match(filename):
                    m_files.append(filename)
                else:
                    s_files.append(filename)
            m_files = list(dict.fromkeys(m_files))
            s_files = list(dict.fromkeys(s_files))
            if m_files:
                valid_ids_with_m_files.append(id)
                pubmed_ids_with_m_files.append(id)
                all_m_files.extend(m_files)
            else:
                valid_ids_without_m_files.append(id)
                pubmed_ids_without_m_files.append(id)
            if s_files:
                pubmed_ids_with_s_files.append(id)
                all_s_files.extend(s_files)
            else:
                pubmed_ids_without_s_files.append(id)
        elif re.match(r"^[^0-9].*$", id):
            valid_ids.append(id)
            non_pubmed_valid_ids.append(id)
            m_files = list(main_by_stem.get(id, []))
            if m_files:
                valid_ids_with_m_files.append(id)
                nonpubmed_ids_with_m_files.append(id)
                all_m_files.extend(m_files)
            else:
                valid_ids_without_m_files.append(id)
                nonpubmed_ids_without_m_files.append(id)
        else:
            invalid_ids.append(id)

    return (
        valid_ids,
        pubmed_ids,
        non_pubmed_valid_ids,
        invalid_ids,
        valid_ids_with_m_files,
        valid_ids_without_m_files,
        pubmed_ids_with_m_files,
        pubmed_ids_without_m_files,
        pubmed_ids_with_s_files,
        pubmed_ids_without_s_files,
        all_m_files,
        all_s_files,
        nonpubmed_ids_with_m_files,
        nonpubmed_ids_without_m_files,
        file_map,
    )


def copy_list(lst):  # copy list or string.
    if not lst:  # Check if the list is empty
        return None
    if isinstance(lst, str):
        result = lst
    else:
        result = "\n".join(lst)
    write_clipboard(result)


def trim_range(rng, reselect=True):
    wb = xw.books.active
    ws = xw.sheets.active

    used_range = ws.used_range
    rng = ws.range(
        (rng.row, rng.column),
        (rng.rows[-1].row, rng.columns[-1].column),
    )
    if rng.row < used_range.row:
        rng = ws.range(
            (used_range.row, rng.column),
            (rng.rows[-1].row, rng.columns[-1].column),
        )
    if rng.rows[-1].row > used_range.rows[-1].row:
        rng = ws.range(
            (rng.row, rng.column),
            (used_range.rows[-1].row, rng.columns[-1].column),
        )
    if rng.column < used_range.column:
        rng = ws.range(
            (rng.row, used_range.column),
            (rng.rows[-1].row, rng.columns[-1].column),
        )
    if rng.columns[-1].column > used_range.columns[-1].column:
        rng = ws.range(
            (rng.row, rng.column),
            (rng.rows[-1].row, used_range.columns[-1].column),
        )

    rowmin = rng.rows[-1].row
    rowmax = rng.row
    columnmin = rng.columns[-1].column
    columnmax = rng.column

    notNone = False
    for row in rng.rows:
        for cell in row:
            if cell.value is None:
                continue
            notNone = True
            if cell.row < rowmin:
                rowmin = cell.row
            if cell.row > rowmax:
                rowmax = cell.row
            if cell.column < columnmin:
                columnmin = cell.column
            if cell.column > columnmax:
                columnmax = cell.column

    if not notNone:
        print("No values found in the")
        return rng
    else:
        rng = ws.range((rowmin, columnmin), (rowmax, columnmax))
        if reselect:
            rng.select()
        return rng


def check_file_exist(mainlibdir, seclibdir=None):
    if seclibdir is None:
        seclibdir = []

    def is_none(value):
        return value is None

    def num_to_str(num):
        if isinstance(num, str):
            return num
        elif num % 1 == 0:
            return str(int(num))
        else:
            return str(num)

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

    if rng is None:
        raise ValueError("No selection made.. Please make a selection in the Excel sheet.")

    if rng.count > 1000:
        raise ValueError("Please select 1000 or fewer cells!")

    rng = trim_range(rng, reselect=True)

    if rng is None:
        raise ValueError("No selection made.. Please make a selection in the Excel sheet.")

    column_interval = (rng.column, rng.column + rng.columns.count - 1)
    header_range = ws.range((1, column_interval[0]), (1, column_interval[1]))

    if header_range.count == 0:
        header_range.select()
        raise ValueError(
            "Error: Please ensure the following before trying again.:\n"
            "1. The table header must be located in the first row of the entire Excel sheet (Row 1).\n"
            "2. The column header containing PMIDs is labeled as 'Ref'."
        )

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

    ref_column = None
    for cell in header_range:
        if cell.value is not None and cell.value.lower() == "ref":
            ref_column = cell.column
            break

    extended_selection = ws.range((rng.row, ref_column), (rng.row + rng.rows.count - 1, ref_column))
    extended_selection.select()
    rng = wb.app.selection

    isfileColor = (188, 219, 255)

    totalfilecount = 0
    isfilecount = 0
    nofilecount = 0

    catalog = build_library_catalog(mainlibdir, seclibdir)
    catalog_index = _build_catalog_index(list(catalog.keys()))

    for i in list(rng):
        b = i.value
        if is_none(b):
            continue
        b = num_to_str(b)
        if "|" in b:
            continue
        isfilecurrentcell = True

        if len(process_ids(b, mainlibdir, seclibdir, catalog=catalog, catalog_index=catalog_index)[4]) > 0:
            isfilecount = isfilecount + 1
            totalfilecount = totalfilecount + 1
        else:
            nofilecount = nofilecount + 1
            totalfilecount = totalfilecount + 1
            isfilecurrentcell = False

        if isfilecurrentcell:
            i.color = isfileColor
        elif i.color == isfileColor:
            i.color = None

    print(
        "requested cells: " + str(rng.count) + "\n"
        "requested files: " + str(totalfilecount) + "\n"
        "isfile count: " + str(isfilecount) + "\n"
        "nofile count: " + str(nofilecount) + "\n"
    )

    return (
        str(rng.count) + " cells checked." + "\n" + "Cells with files are now colored blue."
    )
