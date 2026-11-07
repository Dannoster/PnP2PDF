import io
import re
from pathlib import Path

from PIL import Image, ImageStat


def _make_cards(pnp, folder, count=4, with_backs=False):
    folder.mkdir(parents=True, exist_ok=True)
    cards = []
    for i in range(count):
        front = folder / f"front_{i}.png"
        Image.new("RGB", (300, 400), (255, 0, 0)).save(front)
        back = None
        if with_backs:
            back_path = folder / f"back_{i}.png"
            Image.new("RGB", (300, 400), (0, 0, 255)).save(back_path)
            back = pnp.CardBack(str(back_path), 62.5, 88.7)
        cards.append(pnp.Card(str(front), 62.5, 88.7, back))
    return cards


def _pdf_in(out_dir):
    return next(Path(out_dir).rglob("*.pdf"))


def _page_count(pdf_path):
    data = Path(pdf_path).read_bytes()
    return data.count(b"/Type /Page") - data.count(b"/Type /Pages")


def _jpeg_magentas(pdf_path):
    data = Path(pdf_path).read_bytes()
    magentas = []
    for chunk in re.findall(rb"stream\r?\n(.*?)\r?\nendstream", data, re.S):
        if chunk[:2] != b"\xff\xd8":
            continue
        image = Image.open(io.BytesIO(chunk))
        image.load()
        if image.mode != "CMYK":
            image = image.convert("CMYK")
        magentas.append(ImageStat.Stat(image).mean[1])
    return magentas


def test_save_writes_single_pdf(pnp, tmp_path):
    cards = _make_cards(pnp, tmp_path / "cards")
    sheets = pnp.create_sheets(cards, 210, 297, 2, 2, 3, 5, 5)
    out_dir = tmp_path / "out"

    pnp.PDF_file(sheets).save(str(out_dir), double_sided=False)

    subfolders = [d for d in out_dir.iterdir() if d.is_dir()]
    assert len(subfolders) == 1
    files = list(subfolders[0].iterdir())
    assert len(files) == 1
    assert files[0].suffix == ".pdf"


def test_repeated_save_uses_distinct_folders(pnp, tmp_path):
    cards = _make_cards(pnp, tmp_path / "cards")
    out_dir = tmp_path / "out"

    for _ in range(2):
        sheets = pnp.create_sheets(cards, 210, 297, 2, 2, 3, 5, 5)
        pnp.PDF_file(sheets).save(str(out_dir), double_sided=False)

    subfolders = sorted(d for d in out_dir.iterdir() if d.is_dir())
    assert len(subfolders) == 2
    assert subfolders[0] != subfolders[1]
    for subfolder in subfolders:
        assert list(subfolder.glob("*.pdf"))


def test_page_count_matches_sheets(pnp, tmp_path):
    cards = _make_cards(pnp, tmp_path / "cards")
    out_dir = tmp_path / "single"
    sheets = pnp.create_sheets(cards, 210, 297, 2, 2, 3, 5, 5)
    pnp.PDF_file(sheets).save(str(out_dir), double_sided=False)
    assert _page_count(_pdf_in(out_dir)) == len(sheets)

    back_cards = _make_cards(pnp, tmp_path / "cards2", with_backs=True)
    double_dir = tmp_path / "double"
    double_sheets = pnp.create_sheets(back_cards, 210, 297, 2, 2, 3, 5, 5)
    pnp.PDF_file(double_sheets).save(str(double_dir), double_sided=True)
    assert _page_count(_pdf_in(double_dir)) == 2 * len(double_sheets)


def test_pages_are_cmyk(pnp, tmp_path):
    cards = _make_cards(pnp, tmp_path / "cards")
    sheets = pnp.create_sheets(cards, 210, 297, 2, 2, 3, 5, 5)
    out_dir = tmp_path / "out"

    pnp.PDF_file(sheets).save(str(out_dir), double_sided=False)

    data = _pdf_in(out_dir).read_bytes()
    assert b"/DeviceCMYK" in data
    assert b"/DeviceRGB" not in data


def test_double_sided_alternates_front_and_back(pnp, tmp_path):
    cards = _make_cards(pnp, tmp_path / "cards", count=4, with_backs=True)
    sheets = pnp.create_sheets(cards, 210, 297, 1, 2, 3, 5, 5)
    assert len(sheets) == 2
    out_dir = tmp_path / "out"

    pnp.PDF_file(sheets).save(str(out_dir), double_sided=True)

    magentas = _jpeg_magentas(_pdf_in(out_dir))
    assert len(magentas) == 4
    midpoint = (max(magentas) + min(magentas)) / 2
    assert (magentas[0] - midpoint) * (magentas[2] - midpoint) > 0
    assert (magentas[1] - midpoint) * (magentas[3] - midpoint) > 0
    assert (magentas[0] - midpoint) * (magentas[1] - midpoint) < 0
