from pathlib import Path

import pytest


def _form(**overrides):
    values = {
        '-CFOLDER-': '',
        '-PDFFOLDER-': '',
        '-PWIDTH-': '210',
        '-PHEIGHT-': '297',
        '-CWIDTH-': '62,5',
        '-CHEIGHT-': '88,7',
        '-ROWS-': '3',
        '-COLUMNS-': '3',
        '-GAP-': '2',
        '-HSZONES-': '0',
        '-VSZONES-': '0',
        '-SIZECHECKBOX-': False,
        '-DOUBLECHECKBOX-': False,
    }
    values.update(overrides)
    return values


def test_from_form_accepts_comma_and_dot(pnp):
    comma = pnp.JobOptions.from_form(_form(**{'-CWIDTH-': '62,5'}))
    dot = pnp.JobOptions.from_form(_form(**{'-CWIDTH-': '62.5'}))

    assert comma.card_width == 62.5
    assert dot.card_width == 62.5


def test_from_form_rejects_text_in_size_field(pnp):
    with pytest.raises(pnp.InputError) as excinfo:
        pnp.JobOptions.from_form(_form(**{'-CWIDTH-': 'abc'}))

    assert "Card width" in str(excinfo.value)
    assert "abc" in str(excinfo.value)


def test_from_form_rejects_fractional_rows(pnp):
    with pytest.raises(pnp.InputError) as excinfo:
        pnp.JobOptions.from_form(_form(**{'-ROWS-': '2.5'}))

    assert "Rows" in str(excinfo.value)
    assert "2.5" in str(excinfo.value)


def test_validate_rejects_bad_values(pnp):
    with pytest.raises(pnp.InputError):
        pnp.JobOptions(cards_folder='  ', output_folder='out').validate()
    with pytest.raises(pnp.InputError):
        pnp.JobOptions(cards_folder='cards', output_folder='out', rows=0).validate()
    with pytest.raises(pnp.InputError):
        pnp.JobOptions(cards_folder='cards', output_folder='out', gap=-1).validate()


def test_create_pdf_writes_pdf(pnp, card_folder, tmp_path):
    folder = card_folder(names=['1.png', '2.png'])
    options = pnp.JobOptions(cards_folder=str(folder),
                             output_folder=str(tmp_path / 'out'))

    path = pnp.create_pdf(options)

    assert Path(path).exists()
    assert list(Path(path).glob('*.pdf'))
