# PnP2PDF

[![tests](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml/badge.svg)](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml)

**English** · [Русский](README.ru.md) · [简体中文](README.zh-CN.md) · [Español](README.es.md) · [Deutsch](README.de.md)

PnP2PDF assembles print-and-play cards into a ready-to-print PDF. Pick a folder
of card images, choose the paper and the grid, and get one PDF with cut lines,
safe zones and — when you need them — the card backs mirrored on the reverse
side.

- one window, no command line needed;
- live preview of the sheets before you print;
- cards without backs, one shared back, several different backs, and cards that
  are their own back;
- defaults that fit A4 and 62.5 × 88.7 mm (poker size) cards.

## Contents

- [Prebuilt app](#prebuilt-app)
- [Running from source](#running-from-source)
- [Quick start](#quick-start)
- [Card folder layouts](#card-folder-layouts)
- [Settings](#settings)
- [Examples](#examples)
- [Command line](#command-line)
- [Building a standalone app](#building-a-standalone-app)
- [Troubleshooting](#troubleshooting)
- [Development](#development)
- [Known issues](#known-issues)
- [License](#license)

## Prebuilt app

Download the archive for your system from
[Releases](https://github.com/Dannoster/PnP2PDF/releases), unpack it, run it. No
Python installation is needed.

macOS builds are not signed with an Apple Developer ID, so the first launch of a
downloaded copy needs the quarantine flag removed:

```bash
xattr -dr com.apple.quarantine /Applications/PnP2PDF.app
open /Applications/PnP2PDF.app
```

## Running from source

You need Python 3.11 or newer.

```bash
git clone https://github.com/Dannoster/PnP2PDF.git
cd PnP2PDF
pip install -r requirements.txt
python PnP2PDF.py
```

Installing the package adds two commands — `pnp2pdf` opens the window,
`pnp2pdf-cli` renders without one:

```bash
pip install .
pnp2pdf
```

## Quick start

1. **Cards folder** — pick the folder with your card images. The preview on the
   right fills in as soon as a folder is chosen.
2. **Paper and card size** — defaults are A4 (210 × 297 mm) and a 62.5 × 88.7 mm
   card. Tick *Smaller size (98%)* to print cards slightly smaller, which leaves
   more paper between them.
3. **Layout** — rows × columns, the gap between cards, and optional safe zones
   (margins your printer cannot reach).
4. **Double sided cards** — tick this when the folder is set up with card backs.
   See [Card folder layouts](#card-folder-layouts): without the tick, back
   images are printed as ordinary cards.
5. **Create PDF** — the progress bar runs to the end, the window stays open, the
   status line turns green with the file name, and **Show PDF** reveals the file
   in Finder, Explorer or your file manager.

The PDF is written into a timestamped folder next to your cards. Tick *Save the
PDF to a different folder* to put it somewhere else.

## Card folder layouts

The *cards folder* is the folder you choose in the window. Everything below
lives inside it. Supported images: `.png`, `.jpg`, `.jpeg`, `.tif`, `.tiff`.

Images are stretched to the card area, so use images with the same aspect ratio
as your cards.

### Cards without backs

Put the images in the folder, tick nothing:

```
my-cards/
├── card-01.png
├── card-02.png
└── card-03.png
```

### One back shared by every card

The back is an image whose name ends with its size in millimetres, and its cards
sit in a subfolder named like the back **without** the size part:

```
my-cards/
├── red-back_63,5_88,9.png     ← the back, with its size in the name
└── red-back/                  ← cards printed with that back
    ├── card-01.png
    └── card-02.png
```

Tick **Double sided cards** and build. Every sheet of fronts is followed by the
same sheet of backs, mirrored so that the cards line up when the sheet is
printed on both sides.

### Several backs

Repeat the pattern: one back file, one folder per back. Cards from different
backs can have different sizes and are laid out together.

```
my-cards/
├── red-back_63,5_88,9.png
├── blue-back_64,5_91,3.png
├── red-back/
│   └── card-01.png …
└── blue-back/
    └── card-01.png …
```

The size suffix is optional — without it the back is stretched to the card size.
With it, a back that is larger than the card is cropped to the card window, so
bleed and registration marks stay off the printed card. Use a comma or a dot for
the fractional part: `63,5` and `63.5` both work.

### Cards that are their own back

When a back belongs to exactly one card, put both pictures in a folder named
`Double Sided`. Files are paired in sorted order: front, back, front, back, …

```
my-cards/
└── Double Sided/
    ├── 1-front.png
    ├── 2-back.png
    └── 3-front.png …
```

An odd number of files is rejected with an explanation, because the pairing
would silently slip.

### Cards without a back in a mixed set

Add a `Single Sided` folder for them:

```
my-cards/
├── Double Sided/…
└── Single Sided/
    ├── card-01.png
    └── card-02.png
```

All four layouts can be combined in one folder and built in a single run.

> **Do not forget the tick.** The folder structure says *which back belongs to
> which card*; the **Double sided cards** checkbox asks for those backs to be
> printed. With the box unticked, back images are treated as ordinary fronts and
> end up on the page next to your cards.

## Settings

| Setting | Default | Notes |
| --- | --- | --- |
| Paper size | 210 × 297 mm | A4. Any size in millimetres. |
| Card size | 62.5 × 88.7 mm | Poker size. Images are stretched to this box. |
| Smaller size (98%) | off | Prints each card at 98% of the size above. |
| Rows × columns | 3 × 3 | Cards per sheet. |
| Gap | 2 mm | Space between cards, where the cut lines run. |
| Safe zone top/bottom | 0 mm | Margins the printer cannot reach. |
| Safe zone left/right | 0 mm | The same, horizontally. |
| Cards folder | — | The folder described above. |
| Save the PDF to | next to the cards | Optional; choose another folder here. |
| Double sided cards | off | Print the backs of the cards. |

Settings are remembered between runs as `presets.json` in the application data
folder (paths in [Troubleshooting](#troubleshooting)). A checkout keeps the file
next to the sources.

## Examples

[`examples/`](examples/) holds ready-made card sets for every layout — a few
kilobytes each — plus `make_examples.py` to regenerate them. Details are in
[`examples/README.md`](examples/README.md).

| Folder | What it shows | Pages |
| --- | --- | --- |
| `single-sided/` | nine cards, no backs | 1 |
| `one-back/` | six cards sharing one back | 2 |
| `two-backs/` | two backs, one larger than the cards | 4 |
| `double-sided/` | paired cards plus single-sided cards in one run | 2 |

```bash
python3 -m pnp2pdf.cli --cards examples/one-back --out /tmp/ex --double-sided --quiet
```

## Command line

`pnp2pdf-cli` writes the same PDFs without opening a window — useful for scripts
and batch runs.

```bash
pnp2pdf-cli --cards ~/cards --out ~/pdf
pnp2pdf-cli --cards ~/cards --out ~/pdf --double-sided
pnp2pdf-cli --cards ~/cards --preview /tmp/first-page.png
```

| Option | Meaning |
| --- | --- |
| `--cards DIR` | Folder with the card images (required). |
| `--out DIR` | Where the PDF goes. Default: next to the cards. |
| `--paper WxH` | Paper size in mm, default `210x297`. |
| `--card WxH` | Card size in mm, default `62.5x88.7`. |
| `--grid RxC` | Rows × columns, default `3x3`. |
| `--gap MM` | Gap between cards in mm, default `2`. |
| `--safe-ud MM` | Top/bottom safe zone in mm, default `0`. |
| `--safe-lr MM` | Left/right safe zone in mm, default `0`. |
| `--smaller-size` | Print cards at 98% of the given size. |
| `--double-sided` | Print the card backs. |
| `--preview PATH` | Write the first page as a PNG and exit. |
| `--quiet` | No progress output. |
| `--version` | Print the version and exit. |

Exit codes: `0` success, `1` unexpected error, `2` bad input, `130` cancelled.

## Building a standalone app

One command builds the bundled app and repairs its own environment:

```bash
./build_app.sh          # macOS and Linux
build_app.bat           # Windows
```

The script creates `.venv` when it is missing, installs or updates the
requirements, replaces a broken or foreign virtual environment, runs the tests
and packages the app. Results:

- macOS: `dist/PnP2PDF.app`
- Windows: `dist/PnP2PDF.exe`
- Linux: `dist/PnP2PDF`

`--no-tests` skips the test run; `PNP2PDF_BUILD_DRY=1` checks the environment and
stops before packaging. On macOS the bundle is re-signed after its version is
written, and the build fails loudly if the signature does not verify.

## Troubleshooting

**The packaged app quits or does nothing.** Errors that a windowed build would
swallow are written to `crash.log`, and the error dialog names the file. Send it
when reporting a problem.

- macOS: `~/Library/Application Support/PnP2PDF/crash.log`
- Windows: `%APPDATA%\PnP2PDF\crash.log`
- Linux: `$XDG_DATA_HOME/pnp2pdf/crash.log` (or `~/.local/share/pnp2pdf/crash.log`)

**"PnP2PDF is damaged and can't be opened" (macOS).** A downloaded, unsigned app
is quarantined — see [Prebuilt app](#prebuilt-app).

**"No card images found".** Images must end in `.png`, `.jpg`, `.jpeg`, `.tif` or
`.tiff`; cards with a back belong in the subfolder named after that back.

**The backs landed on the page as cards.** The **Double sided cards** box was not
ticked — see [Card folder layouts](#card-folder-layouts).

**The app opens as a folder (macOS).** The bundle or the launch database is
stale: rebuild with `./build_app.sh`, then run
`/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f dist/PnP2PDF.app`.

## Development

```bash
pip install -r requirements-dev.txt
pytest
```

The suite stubs CustomTkinter, so it runs without a display. Package layout:

```
pnp2pdf/
├── app.py          event loop and start-up
├── cards.py        scanning a folder, pairing cards with backs
├── layout.py       filling sheets
├── render.py       drawing a page: cards, cut lines, corners, safe zones
├── pdf.py          writing the CMYK PDF
├── preview.py      fast low-resolution rendering for the preview panel
├── job.py          one build: cards → sheets → PDF
├── cli.py          the command line entry point
└── ui/             window, form helpers, worker thread
```

## Known issues

- The Windows build may report an error after the window is closed; the PDF is
  already written at that point.
- Microsoft Defender sometimes flags PyInstaller builds as unknown software.
- The macOS build is unsigned, hence the quarantine step above. Signing and
  notarising need an Apple Developer ID.

## License

[MIT](LICENSE).
