# Examples

Runnable card sets that show every folder layout PnP2PDF understands. They are
plain PNGs (a few kilobytes each), so a build works without running anything
first. The script `make_examples.py` regenerates them.

## Layout rules

- A **cards folder** holds the card backs as images. A back may encode its own
  size in millimetres after an underscore: `red-back_63,5_88,9.png` (comma or
  dot both work).
- Next to each back sits a **subfolder named exactly like the back without the
  size suffix** (`red-back/`) holding the cards printed with that back.
- `Double Sided/` holds cards that are their own back. Files are paired in
  sorted order as front, back, front, back, …
- `Single Sided/` holds cards without a back.

```
examples/
├── single-sided/          one-sided cards, no backs
│   └── card-01.png … card-09.png
├── one-back/              one shared back for every card
│   ├── red-back_63,5_88,9.png
│   └── red-back/card-01.png … card-06.png
├── two-backs/             two backs, groups in different colours
│   ├── red-back_63,5_88,9.png
│   ├── blue-back_64,5_91,3.png
│   ├── red-back/card-01.png … card-06.png
│   └── blue-back/card-01.png … card-04.png
└── double-sided/          the two special folders in one run
    ├── Double Sided/1-front.png, 2-back.png, 3-front.png, 4-back.png
    └── Single Sided/card-01.png, card-02.png
```

## Build commands

```bash
python3 -m pnp2pdf.cli --cards examples/single-sided  --out /tmp/ex --quiet
python3 -m pnp2pdf.cli --cards examples/one-back      --out /tmp/ex --double-sided --quiet
python3 -m pnp2pdf.cli --cards examples/two-backs     --out /tmp/ex --double-sided --quiet
python3 -m pnp2pdf.cli --cards examples/double-sided  --out /tmp/ex --double-sided --quiet
```

With `--double-sided` every sheet is followed by a mirrored back page.

- `single-sided`: 9 cards fit one 3×3 A4 sheet → 1 page.
- `one-back`: 6 cards, one mirrored back page → 2 pages.
- `two-backs`: 10 cards use two sheets and the backs alternate by group (red,
  then blue) → 4 pages. The blue back is larger than the card, which exercises
  the `_64,5_91,3` size suffix.
- `double-sided`: 2 paired cards plus 2 single-sided cards on one sheet → 2
  pages; the single-sided cards have no mirrored back.

Add `--preview out.png` instead of `--out` for a quick look at the first page
without writing the whole PDF:

```bash
python3 -m pnp2pdf.cli --cards examples/two-backs --preview /tmp/ex.png
```

## Regenerating

```bash
python3 examples/make_examples.py                 # all sets (README kept)
python3 examples/make_examples.py --only one-back # one set only
```
