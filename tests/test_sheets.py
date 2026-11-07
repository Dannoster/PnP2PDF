import pytest


def _cards(pnp, count):
    return [pnp.Card(f"{i}.png", 62.5, 88.7) for i in range(count)]


def test_sheets_layout(pnp):
    sheets = pnp.create_sheets(_cards(pnp, 7), 210, 297, 2, 3, 2, 0, 0)

    assert len(sheets) == 2
    assert [len(sheet.cards) for sheet in sheets] == [6, 1]


def test_create_sheets_does_not_consume_input(pnp):
    cards = _cards(pnp, 7)

    sheets = pnp.create_sheets(cards, 210, 297, 2, 3, 2, 0, 0)
    again = pnp.create_sheets(cards, 210, 297, 2, 3, 2, 0, 0)

    assert len(cards) == 7
    assert len(sheets) == 2
    assert len(again) == 2


def test_partial_sheet_is_allowed(pnp):
    sheets = pnp.create_sheets(_cards(pnp, 1), 210, 297, 2, 3, 2, 0, 0)

    assert len(sheets) == 1
    assert len(sheets[0].cards) == 1


@pytest.mark.parametrize("rows, columns", [(0, 3), (2, 0), (-1, 3), (2, -1)])
def test_invalid_grid_rejected(pnp, rows, columns):
    with pytest.raises(ValueError):
        pnp.create_sheets(_cards(pnp, 4), 210, 297, rows, columns, 2, 0, 0)


def test_empty_card_list_returns_no_sheets(pnp):
    assert pnp.create_sheets([], 210, 297, 2, 3, 2, 0, 0) == []
