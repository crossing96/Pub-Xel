# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# Backward-compatible re-exports. Prefer importing from pubxel_core.ids,
# pubxel_core.excel_ops, pubxel_core.pubmed, or pubxel_core.paths directly.

from pubxel_core.excel_ops import (
    check_file_exist,
    copy_list,
    files_name_to_path,
    process_ids,
    trim_range,
)
from pubxel_core.ids import list_to_string, set_preserve_order, string_to_list
from pubxel_core.metadata_store import MetadataStore, enrich_metadata, metadataStore
from pubxel_core.paths import appdatadir, data_dir, metadata_path, settings_path
from pubxel_core.pubmed import input_pubmed_data, obtain_pubmed_data, value_from_dict
from pubxel_core.settings import load_settings, save_settings, save_settings_key

__all__ = [
    "appdatadir",
    "check_file_exist",
    "copy_list",
    "data_dir",
    "enrich_metadata",
    "files_name_to_path",
    "input_pubmed_data",
    "list_to_string",
    "load_settings",
    "MetadataStore",
    "metadataStore",
    "metadata_path",
    "obtain_pubmed_data",
    "process_ids",
    "save_settings",
    "save_settings_key",
    "settings_path",
    "set_preserve_order",
    "string_to_list",
    "trim_range",
    "value_from_dict",
]
