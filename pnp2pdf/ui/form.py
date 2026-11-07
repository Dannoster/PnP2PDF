"""Translate between settings dictionaries and the window's widget values.

This module deliberately has no GUI imports: the window and the tests can use
it without importing ``customtkinter`` or ``tkinter``.
"""

# Widget keys in the exact order of the window contract.
FIELD_KEYS = [
    "-PWIDTH-", "-PHEIGHT-", "-CWIDTH-", "-CHEIGHT-",
    "-SIZECHECKBOX-", "-DOUBLECHECKBOX-",
    "-ROWS-", "-COLUMNS-", "-GAP-", "-HSZONES-", "-VSZONES-",
    "-CFOLDER-", "-PDFFOLDER-",
]

# settings key -> widget key, split by the type they carry.
_NUMERIC = {
    "paper_width": "-PWIDTH-",
    "paper_height": "-PHEIGHT-",
    "card_width": "-CWIDTH-",
    "card_height": "-CHEIGHT-",
    "rows": "-ROWS-",
    "columns": "-COLUMNS-",
    "gap": "-GAP-",
    "ud_szone": "-HSZONES-",
    "lr_szone": "-VSZONES-",
}
_BOOLEAN = {
    "resize": "-SIZECHECKBOX-",
    "double_sided": "-DOUBLECHECKBOX-",
}
_TEXT = {
    "cards_folder": "-CFOLDER-",
    "save_to": "-PDFFOLDER-",
}

_TRUTHY = {"1", "true", "yes", "on"}


def _as_bool(value) -> bool:
    if isinstance(value, str):
        return value.strip().lower() in _TRUTHY
    return bool(value)


def initial_values(settings: dict) -> dict:
    """Build the widget-key dict for a settings dict, numbers as strings."""
    values = {}
    for name, widget in _NUMERIC.items():
        values[widget] = str(settings.get(name, ""))
    for name, widget in _BOOLEAN.items():
        values[widget] = bool(settings.get(name, False))
    for name, widget in _TEXT.items():
        values[widget] = str(settings.get(name, ""))
    return {key: values[key] for key in FIELD_KEYS}


def check_values(values: dict) -> dict:
    """Normalise the checkbox entries so ``JobOptions.from_form`` gets bools."""
    checked = dict(values)
    for key in checked:
        if "CHECKBOX" in key:
            checked[key] = _as_bool(checked[key])
    return checked
