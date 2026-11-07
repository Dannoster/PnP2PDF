"""Build the committed example sets to prove the supported layouts work."""

from pathlib import Path

import pytest

pytest.importorskip("PIL")

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"


def test_single_sided(pnp):
    cards = pnp.create_cards(str(EXAMPLES / "single-sided"), 63.5, 88.9,
                             False, False)

    assert len(cards) == 9
    assert all(card.back is None for card in cards)


def test_one_back(pnp):
    cards = pnp.create_cards(str(EXAMPLES / "one-back"), 63.5, 88.9,
                             False, True)

    assert len(cards) == 6
    backs = {card.back.dir for card in cards}
    assert len(backs) == 1
    assert Path(next(iter(backs))).name == "red-back_63,5_88,9.png"
    assert all(card.back.width == pytest.approx(63.5) for card in cards)
    assert all(card.back.height == pytest.approx(88.9) for card in cards)


def test_two_backs(pnp):
    cards = pnp.create_cards(str(EXAMPLES / "two-backs"), 63.5, 88.9,
                             False, True)

    assert len(cards) == 10
    red = [c for c in cards if Path(c.back.dir).name == "red-back_63,5_88,9.png"]
    blue = [c for c in cards if Path(c.back.dir).name == "blue-back_64,5_91,3.png"]
    assert len(red) == 6
    assert len(blue) == 4
    assert {Path(c.back.dir).name for c in cards} == {
        "red-back_63,5_88,9.png", "blue-back_64,5_91,3.png"}
    assert blue[0].back.width == pytest.approx(64.5)
    assert blue[0].back.height == pytest.approx(91.3)


def test_double_sided(pnp):
    cards = pnp.create_cards(str(EXAMPLES / "double-sided"), 63.5, 88.9,
                             False, True)

    assert len(cards) == 4
    paired = [card for card in cards if card.back is not None]
    singles = [card for card in cards if card.back is None]
    assert len(paired) == 2
    assert len(singles) == 2

    by_name = {Path(card.dir).name: card for card in cards}
    assert Path(by_name["1-front.png"].back.dir).name == "2-back.png"
    assert Path(by_name["3-front.png"].back.dir).name == "4-back.png"
    assert all(Path(card.back.dir).parent.name == "Double Sided"
               for card in paired)
    assert all(Path(card.dir).parent.name == "Single Sided" for card in singles)
