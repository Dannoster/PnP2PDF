"""Draw the ready-to-build example card sets under ``examples/``.

Every card is a solid colour with a thin dark border, a big centred number and
the set name, so the printed order is obvious. No extra dependencies beyond
Pillow are used.

Run ``python3 examples/make_examples.py`` to (re)generate everything, or
``python3 examples/make_examples.py --only one-back`` for a single set.
"""

import argparse
import shutil
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

DPI = 300
CARD_MM = (63.5, 88.9)   # 750 x 1050 px
BACK_MM = (64.5, 91.3)   # 762 x 1078 px
BORDER_COLOUR = (35, 35, 35)

SETS = {
    "single-sided": {
        "description": "Nine one-sided cards, no backs.",
        "colour": (58, 150, 134),
        "label": "single-sided",
        "count": 9,
    },
    "one-back": {
        "description": "Six cards sharing the same red back.",
        "colour": (95, 106, 170),
        "label": "red-back",
        "count": 6,
        "back_colour": (176, 52, 52),
        "back_file": "red-back_63,5_88,9.png",
    },
    "two-backs": {
        "description": "Six red-backed cards plus four blue-backed cards.",
        "groups": [
            {
                "folder": "red-back",
                "file": "red-back_63,5_88,9.png",
                "colour": (176, 52, 52),
                "card_colour": (206, 92, 92),
                "count": 6,
            },
            {
                "folder": "blue-back",
                "file": "blue-back_64,5_91,3.png",
                "colour": (45, 80, 180),
                "card_colour": (92, 122, 210),
                "count": 4,
            },
        ],
    },
    "double-sided": {
        "description": "Two double-sided cards plus two single-sided cards.",
        "front_colour": (150, 90, 165),
        "back_colour": (90, 140, 100),
        "single_colour": (196, 140, 60),
        "double_count": 2,
        "single_count": 2,
    },
}


def _font(size):
    return ImageFont.load_default(size=size)


def _draw_centred(draw, text, font, width, centre_y):
    bbox = draw.textbbox((0, 0), text, font=font)
    x = (width - (bbox[2] - bbox[0])) // 2 - bbox[0]
    y = int(centre_y - (bbox[3] - bbox[1]) / 2) - bbox[1]
    draw.text((x, y), text, font=font, fill=BORDER_COLOUR)


def _face_image(size_mm, colour, number, label):
    width = round(size_mm[0] / 25.4 * DPI)
    height = round(size_mm[1] / 25.4 * DPI)

    image = Image.new("RGB", (width, height), colour)
    draw = ImageDraw.Draw(image)
    border = max(4, width // 90)
    draw.rectangle([0, 0, width - 1, height - 1],
                   outline=BORDER_COLOUR, width=border)

    if number:
        _draw_centred(draw, str(number), _font(int(height * 0.4)),
                      width, height * 0.45)
    _draw_centred(draw, label, _font(max(18, int(height * 0.05))),
                  width, height * 0.92)
    return image


def card_image(number, colour, label):
    """A card face in ``colour`` with a centred ``number`` and ``label``."""
    return _face_image(CARD_MM, colour, number, label)


def back_image(colour, label):
    """A card back in ``colour``, a little larger than the cards."""
    return _face_image(BACK_MM, colour, None, label)


def _shade(colour, index):
    """Slightly lighten ``colour`` so the print order is easy to see."""
    factor = 0.82 + 0.06 * ((index - 1) % 4)
    return tuple(min(255, int(channel * factor)) for channel in colour)


def _reset(folder):
    """Clear one example set. The shared ``examples/README.md`` is never touched."""
    folder = Path(folder)
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True)


def write_single_sided(root):
    spec = SETS["single-sided"]
    folder = Path(root) / "single-sided"
    _reset(folder)
    for index in range(1, spec["count"] + 1):
        image = card_image(index, _shade(spec["colour"], index), spec["label"])
        image.save(folder / f"card-{index:02d}.png")


def write_one_back(root):
    spec = SETS["one-back"]
    folder = Path(root) / "one-back"
    _reset(folder)

    back_image(spec["back_colour"], spec["label"]).save(folder / spec["back_file"])
    cards_dir = folder / spec["label"]
    cards_dir.mkdir()
    for index in range(1, spec["count"] + 1):
        image = card_image(index, _shade(spec["colour"], index), spec["label"])
        image.save(cards_dir / f"card-{index:02d}.png")


def write_two_backs(root):
    spec = SETS["two-backs"]
    folder = Path(root) / "two-backs"
    _reset(folder)

    for group in spec["groups"]:
        back_image(group["colour"], group["folder"]).save(folder / group["file"])
        cards_dir = folder / group["folder"]
        cards_dir.mkdir()
        for index in range(1, group["count"] + 1):
            image = card_image(index, _shade(group["card_colour"], index),
                               group["folder"])
            image.save(cards_dir / f"card-{index:02d}.png")


def write_double_sided(root):
    spec = SETS["double-sided"]
    folder = Path(root) / "double-sided"
    _reset(folder)

    double_dir = folder / "Double Sided"
    double_dir.mkdir(parents=True)
    for pair in range(1, spec["double_count"] + 1):
        front = (pair - 1) * 2 + 1
        back = front + 1
        card_image(front, spec["front_colour"], "Double Sided front").save(
            double_dir / f"{front}-front.png")
        card_image(back, spec["back_colour"], "Double Sided back").save(
            double_dir / f"{back}-back.png")

    single_dir = folder / "Single Sided"
    single_dir.mkdir()
    for index in range(1, spec["single_count"] + 1):
        image = card_image(index, _shade(spec["single_colour"], index),
                           "Single Sided")
        image.save(single_dir / f"card-{index:02d}.png")


WRITERS = {
    "single-sided": write_single_sided,
    "one-back": write_one_back,
    "two-backs": write_two_backs,
    "double-sided": write_double_sided,
}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--only", choices=sorted(WRITERS), metavar="NAME",
                        help="regenerate just one set")
    args = parser.parse_args(argv)

    root = Path(__file__).resolve().parent
    names = [args.only] if args.only else list(WRITERS)
    for name in names:
        WRITERS[name](root)
        print(f"wrote {root / name}")


if __name__ == "__main__":
    main()
