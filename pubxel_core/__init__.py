# Pub-Xel core package: paths, settings, metadata, IDs, Excel ops, PubMed.

from pubxel_core.app import bootstrap, run

from pubxel_core.clipboard import ClipboardRead, read_clipboard
from pubxel_core.excel_ops import (
    check_file_exist,
    copy_list,
    files_name_to_path,
    process_ids,
)
from pubxel_core.ids import list_to_string, string_to_list
from pubxel_core.metadata_store import MetadataStore, enrich_metadata, metadataStore
from pubxel_core.paths import (
    appdatadir,
    assets_dir,
    data_dir,
    metadata_path,
    os_name,
    project_dir,
    settings_path,
    ui_dir,
)
from pubxel_core.pubmed import input_pubmed_data, obtain_pubmed_data
from pubxel_core.recent_worksheets import register_recent_worksheet, try_register_active_workbook
from pubxel_core.settings import SettingsDict, load_settings, save_settings, save_settings_key

__all__ = [
    "bootstrap",
    "run",
    "ClipboardRead",
    "read_clipboard",
    "appdatadir",
    "assets_dir",
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
    "SettingsDict",
    "metadata_path",
    "obtain_pubmed_data",
    "os_name",
    "process_ids",
    "project_dir",
    "register_recent_worksheet",
    "try_register_active_workbook",
    "save_settings",
    "save_settings_key",
    "settings_path",
    "string_to_list",
    "ui_dir",
]
