import math
import os
import time

import pytest
from PIL import Image

from pnp2pdf import preview, render
from pnp2pdf.layout import create_sheets
from pnp2pdf.model import Card, CardBack

PAPER_W, PAPER_H = 210.0, 297.0
CARD_W, CARD_H = 62.5, 88.7
TARGET_PX = 560


def _standard_cards(folder, count=12, with_backs=False,
                    card_size=(CARD_W, CARD_H)):
    folder.mkdir(parents=True, exist_ok=True)
    cards = []
    for i in range(1, count + 1):
        colour = (i * 15 % 255, 40, 120)
        front = folder / f"Card_{i}.png"
        Image.new("RGB", (300, 400), colour).save(front)
        back = None
        if with_backs:
            back_path = folder / f"Card_{i}_back.png"
            Image.new("RGB", (300, 400), colour).save(back_path)
            back = CardBack(str(back_path), card_size[0], card_size[1])
        cards.append(Card(str(front), card_size[0], card_size[1], back))
    return cards


def _standard_sheets(folder, count=12, with_backs=False,
                     card_size=(CARD_W, CARD_H)):
    cards = _standard_cards(folder, count, with_backs, card_size)
    return create_sheets(cards, PAPER_W, PAPER_H, 3, 3, 2, 0, 0)


@pytest.mark.parametrize("width_mm, height_mm", [
    (210.0, 297.0),
    (297.0, 420.0),
    (100000.0, 100000.0),
])
def test_fit_px_per_mm_stays_within_limits(width_mm, height_mm):
    px_per_mm = preview.fit_px_per_mm(width_mm, height_mm, TARGET_PX)

    assert px_per_mm > 0
    assert math.isfinite(px_per_mm)
    assert width_mm * px_per_mm <= preview.MAX_SIDE_PX
    assert height_mm * px_per_mm <= preview.MAX_SIDE_PX
    assert (width_mm * px_per_mm) * (height_mm * px_per_mm) <= preview.MAX_PIXELS


@pytest.mark.parametrize("width_mm, height_mm", [
    (0.0, 0.0),
    (0.0, 297.0),
    (210.0, 0.0),
    (-10.0, 297.0),
    (210.0, -10.0),
    (-5.0, -5.0),
])
def test_fit_px_per_mm_survives_nonpositive_sizes(width_mm, height_mm):
    px_per_mm = preview.fit_px_per_mm(width_mm, height_mm, TARGET_PX)

    assert px_per_mm > 0
    assert math.isfinite(px_per_mm)
    assert width_mm * px_per_mm <= preview.MAX_SIDE_PX
    assert height_mm * px_per_mm <= preview.MAX_SIDE_PX


def test_preview_pages_standard_set(tmp_path):
    sheets = _standard_sheets(tmp_path / "cards")
    pages = preview.preview_pages(sheets)

    assert len(pages) == 2
    assert all(page.mode == "RGB" for page in pages)
    for page in pages:
        assert max(page.size) == pytest.approx(TARGET_PX, abs=1)

    assert len(preview.preview_pages(sheets, limit=1)) == 1


def test_double_sided_backs_are_mirrored(tmp_path):
    sheets = _standard_sheets(tmp_path / "cards", with_backs=True)

    assert preview.page_count(sheets, False) == len(sheets)
    assert preview.page_count(sheets, True) == 2 * len(sheets)

    pages = preview.preview_pages(sheets)
    assert len(pages) == 2 * len(sheets)

    front, back = pages[0], pages[1]
    assert front.tobytes() != back.tobytes()

    px_per_mm = preview.fit_px_per_mm(sheets[0].width, sheets[0].height)
    geom = render.page_geometry(sheets[0], px_per_mm)
    x_left = geom["x0"] + geom["cwidth"] // 2
    x_right = (geom["x0"] + (geom["columns"] - 1) * (geom["cwidth"] + geom["gap"])
               + geom["cwidth"] // 2)
    for y in range(geom["y0"] + 5, geom["y0"] + geom["cheight"] - 5, 10):
        assert front.getpixel((x_right, y)) == back.getpixel((x_left, y))


def test_preview_page_background_is_white(tmp_path):
    sheets = _standard_sheets(tmp_path / "cards")
    pages = preview.preview_pages(sheets)

    for page in pages:
        assert page.getpixel((0, 0)) == (255, 255, 255)


def test_low_detail_never_opens_the_cards(tmp_path, monkeypatch):
    opened = []
    real_open = Image.open

    def counting_open(*args, **kwargs):
        opened.append(args[0])
        return real_open(*args, **kwargs)

    sheets = _standard_sheets(tmp_path / "cards", card_size=(0.1, 0.1))
    preview.clear_cache()
    monkeypatch.setattr(Image, "open", counting_open)

    pages = preview.preview_pages(sheets)

    assert pages
    assert opened == []


def test_card_files_opened_once_across_calls(tmp_path, monkeypatch):
    opened = []
    real_open = Image.open

    def counting_open(path, *args, **kwargs):
        opened.append(os.path.basename(path))
        return real_open(path, *args, **kwargs)

    sheets = _standard_sheets(tmp_path / "cards")
    preview.clear_cache()
    monkeypatch.setattr(Image, "open", counting_open)

    preview.preview_pages(sheets)
    preview.preview_pages(sheets)

    assert len(opened) == len(set(opened)) == 12


def test_preview_budget(tmp_path):
    sheets = _standard_sheets(tmp_path / "cards")
    preview.clear_cache()

    start = time.perf_counter()
    pages = preview.preview_pages(sheets)
    elapsed = time.perf_counter() - start

    assert len(pages) == 2
    assert elapsed < 1.0
