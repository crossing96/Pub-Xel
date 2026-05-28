# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

import json
import os
import tempfile
from typing import Any, Dict, TypeAlias

from pubxel_core.paths import settings_path

SettingsDict: TypeAlias = Dict[str, Any]


def load_settings() -> SettingsDict:
    if not os.path.exists(settings_path):
        return {}
    with open(settings_path, "r") as file:
        return json.load(file)


def save_settings(settings: SettingsDict) -> None:
    # Write atomically to avoid corrupting settings.json on crash/power loss.
    # (Performance impact is negligible because saves happen only on user actions.)
    dir_name = os.path.dirname(settings_path)
    os.makedirs(dir_name, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix="settings_", suffix=".json", dir=dir_name)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as file:
            json.dump(settings, file, indent=4)
        os.replace(tmp_path, settings_path)
    finally:
        # If something failed before os.replace, make sure we don't leave temp files behind.
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass


def save_settings_key(settings: SettingsDict, key: str, value: Any) -> SettingsDict:
    settings[key] = value
    save_settings(settings)
    return settings
