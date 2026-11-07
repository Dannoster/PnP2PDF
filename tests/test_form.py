from pnp2pdf.options import JobOptions
from pnp2pdf.settings import DEFAULT_SETTINGS
from pnp2pdf.ui.form import FIELD_KEYS, check_values, initial_values


def test_initial_values_has_every_contract_key():
    values = initial_values(DEFAULT_SETTINGS)

    assert list(values) == FIELD_KEYS
    assert set(values) == set(FIELD_KEYS)


def test_initial_values_stringifies_numbers():
    values = initial_values(DEFAULT_SETTINGS)

    assert values["-PWIDTH-"] == "210"
    assert values["-PHEIGHT-"] == "297"
    assert values["-CWIDTH-"] == "62.5"
    assert values["-CHEIGHT-"] == "88.7"
    assert values["-ROWS-"] == "3"
    assert values["-COLUMNS-"] == "3"
    assert values["-GAP-"] == "2"
    assert values["-HSZONES-"] == "0"
    assert values["-VSZONES-"] == "0"
    assert all(isinstance(values[key], str)
               for key in ("-PWIDTH-", "-PHEIGHT-", "-CWIDTH-", "-CHEIGHT-",
                           "-ROWS-", "-COLUMNS-", "-GAP-", "-HSZONES-",
                           "-VSZONES-"))


def test_from_form_works_once_only_the_folders_are_filled():
    settings = dict(DEFAULT_SETTINGS, cards_folder="/tmp/cards", save_to="/tmp/out")

    options = JobOptions.from_form(check_values(initial_values(settings)))

    assert options.cards_folder == "/tmp/cards"
    assert options.output_folder == "/tmp/out"
    assert options.paper_width == 210
    assert options.paper_height == 297
    assert options.card_width == 62.5
    assert options.card_height == 88.7
    assert options.rows == 3
    assert options.columns == 3
    assert options.smaller_size is False
    assert options.double_sided is False


def test_from_form_on_bare_defaults_ignores_empty_folders():
    options = JobOptions.from_form(check_values(initial_values(DEFAULT_SETTINGS)))

    assert options.paper_height == 297
    assert options.smaller_size is False
    assert options.double_sided is False


def test_check_values_turns_strings_and_ints_into_bool():
    checked = check_values({"-SIZECHECKBOX-": "True", "-DOUBLECHECKBOX-": 0})

    assert checked["-SIZECHECKBOX-"] is True
    assert checked["-DOUBLECHECKBOX-"] is False


def test_check_values_accepts_numbers_and_false_strings():
    assert check_values({"-SIZECHECKBOX-": "False"})["-SIZECHECKBOX-"] is False
    assert check_values({"-DOUBLECHECKBOX-": 1})["-DOUBLECHECKBOX-"] is True
    assert check_values({"-DOUBLECHECKBOX-": True})["-DOUBLECHECKBOX-"] is True
