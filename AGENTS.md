# Pub-Xel — agent guide

Biomedical reference manager integrated with **Microsoft Excel** and **PubMed**. PyQt6 desktop app; Windows and macOS.

## Entry point

- **`Pub-Xel.py`** — PyInstaller entry script; calls `pubxel_core.app.run()`.
- **`pubxel_core/app.py`** — `bootstrap()` (QApplication, lock, splash, settings) and `run()` (main window + event loop).
- Do not rename `Pub-Xel.py` without updating `build.ps1`, `setup.iss`, and CI workflows.

## Layout

| Path | Role |
|------|------|
| `Pub-Xel.py` | Thin entry: `sys.exit(run())` |
| `pubxel_core/app.py` | `bootstrap()`, `run()` — full startup and event loop |
| `pubxel_core/runtime.py` | Mutable app state (`settings`, paths, flags) set at startup |
| `pubxel_core/ui/main_window.py` | Main Qt window |
| `pubxel_core/ui/preferences.py` | Preferences dialog |
| `pubxel_core/ui/widgets.py` | About, inspect, worksheet columns, popup widgets |
| `pubxel_core/ui/workers.py` | Excel thread, global hotkey listener |
| `pubxel_core/clipboard.py` | Qt clipboard (`read_clipboard`, main thread). **Open Files:** `file_paths` if `kind==files`, else `ids`→library. **Inspect/PubMed:** `ids` only, never `file_paths`. See module docstring. |
| `pubxel_core/ui/tray.py` | System tray icon |
| `pubxel_core/ui/helpers.py` | `dialog_onebutton`, open folder, shutdown helpers |
| `pubxel_core/update.py` | Weekly update check |
| `pubxel_core/paths.py` | **Single source of truth** for `project_dir`, `assets_dir`, `ui_dir`, `data_dir`, `journal_combined_path`, `pubsheet_all_columns_path`, `appdatadir`, `settings_path`, `metadata_path`, `os_name` |
| `pubxel_core/settings.py` | Load/save `settings.json` under app data dir |
| `pubxel_core/metadata_store.py` | SQLite metadata (`MetadataStore`, `metadataStore` alias, `enrich_metadata`) |
| `pubxel_core/ids.py` | `split_clipboard_text` + per-token `ids_from_clipboard_token` (PMID URL, path stem, numeric); legacy `string_to_list` still uses `:` delimiter |
| `pubxel_core/excel_ops.py` | xlwings + library files: `process_ids`, `files_name_to_path`, `check_file_exist`, `trim_range`, `copy_list` |
| `pubxel_core/pubmed.py` | `obtain_pubmed_data` (NCBI fetch → SQLite upsert → enriched `MetadataDict`); `input_pubmed_data` (Excel fill + impact factors) |
| `pubxel_core/mainfunctions.py` | **Shim** — re-exports the above for older imports |
| `pubxel_core/welcome.py` | First-run welcome dialog |
| `data/` | Bundled defaults: `settingsdefault.json`, `version.py`, `journal_combined_2025.txt`, Excel templates |
| `ui/` | Qt Designer `.ui` files |
| `assets/` | Icons, splash, welcome images |

## User data (not in repo)

| OS | App data directory |
|----|-------------------|
| Windows | `%APPDATA%\pubxel` |
| macOS | `~/Library/Application Support/pubxel` |
| Linux (unsupported in UI) | `~/.pubxel` |

Contains `settings.json`, `pubsheet.xlsx` (generated on first run from column preferences), `metadata/metadata_article.sqlite`, lock file, update-check timestamp.

## Run locally

```powershell
pip install -r requirements.txt
python Pub-Xel.py
```

Requires Excel open for xlwings features. Supported OS: Windows, macOS (`paths.os_name`).

## Build

```powershell
.\build.ps1
```

PyInstaller onedir; bundles `data`, `ui`, `assets`, `pubxel_core`. See `.github/workflows/`.

## Import conventions

- Prefer explicit imports, e.g. `from pubxel_core.ids import string_to_list`.
- Avoid `from pubxel_core.mainfunctions import *` in new code (shim exists for compatibility).
- Version: `from data.version import __version__`

## Module-level state (`pubxel_core/runtime.py`)

Set during `bootstrap()`; UI modules use `from pubxel_core import runtime as rt`:

- `rt.settings` — dict from `load_settings()`, updated via `save_settings` / `save_settings_key`
- `rt.action_in_progress` — blocks overlapping inspect/open actions; writes go through `rt.try_begin_action()` / `rt.end_action()` / `rt.action_guard()` (thread-safe; reads are still free)
- `rt.mainlibdir`, `rt.seclibdir`, `rt.outdir` — from settings after startup
- Hotkey listener — started from `run()` when hotkeys are enabled

## Alternate entry (development)

```powershell
python -c "from pubxel_core.app import run; run()"
```

## Tooling (Phase 4)

| Tool | Purpose |
|------|---------|
| `scripts/smoke_imports.py` | Import-check all modules without starting the GUI |
| `pyproject.toml` | Ruff config (`pip install ruff`, then `ruff check pubxel_core`) |
| `.github/workflows/smoke.yml` | CI smoke on push/PR |

```powershell
python scripts/smoke_imports.py
```

## PubMed metadata flow

- **`obtain_pubmed_data(pmids)`** — NCBI MEDLINE fetch → `MetadataStore.upsert_metadata` → `get_metadata` → `enrich_metadata` → returns `MetadataDict`
- **`input_pubmed_data()`** — reads Excel selection, calls `obtain_pubmed_data`, writes cells (plus journal impact factors from `data/journal_combined_2025.txt`)

## Types and naming

- `SettingsDict` — type alias for `settings.json` contents (`pubxel_core.settings`)
- `MetadataStore` — preferred SQLite metadata class; `metadataStore` remains as an alias
- `ArticleRow` / `AuthorRow` / `MetadataDict` — PubMed/SQLite row types in `metadata_store.py`
- Public APIs in `ids.py`, `settings.py`, `pubmed.py` have type hints for agent/IDE use
