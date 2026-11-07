from pathlib import Path

import pytest


def test_cards_sorted_naturally(pnp, card_folder):
    folder = card_folder(names=[f"{i}.png" for i in range(1, 13)])

    cards = pnp.create_cards(str(folder), 62.5, 88.7, False, False)

    assert [Path(card.dir).name for card in cards] == [
        f"{i}.png" for i in range(1, 13)
    ]


def test_back_size_with_dot(pnp, card_folder):
    folder = card_folder(names=["1.png", "1_62.5_63.0.png", "1/card.png"])

    cards = pnp.create_cards(str(folder), 62.5, 88.7, False, True)

    sized = [c for c in cards if Path(c.back.dir).name == "1_62.5_63.0.png"]
    assert len(sized) == 1
    assert sized[0].back.width == pytest.approx(62.5)
    assert sized[0].back.height == pytest.approx(63.0)


def test_back_size_with_comma(pnp, card_folder):
    folder = card_folder(names=["1_64,5_91,3.png", "1/card.png"])

    cards = pnp.create_cards(str(folder), 62.5, 88.7, False, True)

    assert len(cards) == 1
    assert cards[0].back.width == pytest.approx(64.5)
    assert cards[0].back.height == pytest.approx(91.3)


def test_back_without_size_falls_back_to_card_size(pnp, card_folder):
    folder = card_folder(names=["1_back.png", "1/card.png"])

    cards = pnp.create_cards(str(folder), 62.5, 88.7, False, True)

    assert len(cards) == 1
    assert cards[0].back.width == pytest.approx(62.5)
    assert cards[0].back.height == pytest.approx(88.7)


def test_back_size_scaled_by_koef(pnp, card_folder):
    names = ["1_62.5_63.0.png", "1/card.png"]

    full = pnp.create_cards(str(card_folder(names=names)), 62.5, 88.7, False, True)
    small = pnp.create_cards(str(card_folder(names=names)), 62.5, 88.7, True, True)

    full_back = next(c for c in full if Path(c.back.dir).name == "1_62.5_63.0.png")
    small_back = next(c for c in small if Path(c.back.dir).name == "1_62.5_63.0.png")

    assert full_back.back.width == pytest.approx(62.5)
    assert full_back.back.height == pytest.approx(63.0)
    assert small_back.back.width == pytest.approx(full_back.back.width * 0.98)
    assert small_back.back.height == pytest.approx(full_back.back.height * 0.98)


def test_non_image_files_ignored(pnp, card_folder):
    folder = card_folder(names=["1.png", "notes.txt", "card.png.backup"])

    cards = pnp.create_cards(str(folder), 62.5, 88.7, False, False)

    assert len(cards) == 1
    assert Path(cards[0].dir).name == "1.png"


def test_empty_folder_raises_value_error(pnp, card_folder):
    folder = card_folder(names=[])

    with pytest.raises(ValueError) as excinfo:
        pnp.create_cards(str(folder), 62.5, 88.7, False, False)

    message = str(excinfo.value)
    assert str(folder) in message
    assert "png" in message.lower()


def test_missing_folder_raises_value_error(pnp, tmp_path):
    missing = tmp_path / "does-not-exist"

    with pytest.raises(ValueError) as excinfo:
        pnp.create_cards(str(missing), 62.5, 88.7, False, False)

    message = str(excinfo.value)
    assert str(missing) in message
    assert "png" in message.lower()
