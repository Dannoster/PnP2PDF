import os
import json


DEFAULT_SETTINGS = {"paper_width"   : 210,
            "paper_height"  : 297,
            "card_width"    : 62.5,
            "card_height"   : 88.7,
            "resize"        : False,
            "double_sided"  : False,
            "rows"          : 3,
            "columns"       : 3,
            "gap"           : 2,
            "ud_szone"      : 0, 
            "lr_szone"      : 0,
            "cards_folder"  : "",
            "save_to"       : ""}


def load_save(path: str) -> dict:
    """Read settings from JSON, filling in any key missing from an older file."""
    settings = dict(DEFAULT_SETTINGS)
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as file:
                stored = json.load(file)
        except (OSError, ValueError) as error:
            print(f"WARNING: could not read settings from {path}: {error}")
            stored = {}
        if isinstance(stored, dict):
            for key in DEFAULT_SETTINGS:
                if key in stored:
                    settings[key] = stored[key]
    return settings


def create_save(path: str, save: dict):
    with open(path, "w", encoding="utf-8") as file:
        json.dump(save, file, indent=2, ensure_ascii=False)
