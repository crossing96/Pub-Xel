# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

import os
import platform

os_name = platform.system()

if os_name == "Windows":
    appdatadir = os.path.join(os.getenv("APPDATA"), "pubxel")
elif os_name == "Darwin":
    appdatadir = os.path.expanduser("~/Library/Application Support/pubxel")
else:
    appdatadir = os.path.join(os.path.expanduser("~"), ".pubxel")

os.makedirs(appdatadir, exist_ok=True)

settings_path = os.path.join(appdatadir, "settings.json")

metadata_dir = os.path.join(appdatadir, "metadata")
os.makedirs(metadata_dir, exist_ok=True)
metadata_path = os.path.join(metadata_dir, "metadata_article.sqlite")

_package_dir = os.path.dirname(os.path.abspath(__file__))
project_dir = os.path.dirname(_package_dir)
assets_dir = os.path.join(project_dir, "assets")
ui_dir = os.path.join(project_dir, "ui")
data_dir = os.path.join(project_dir, "data")
journal_combined_path = os.path.join(data_dir, "journal_combined_2025.txt")
pubsheet_all_columns_path = os.path.join(data_dir, "pubsheet_all_columns.xlsx")
pubsheetinitial_path = os.path.join(data_dir, "pubsheetinitial.xlsx")
