# PnP2PDF

[![tests](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml/badge.svg)](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml)

What is it?

PnP2PDF (PnP to PDF) is a software that can help you to prepate your PnP cards for printing. The program helps you to assemble a PDF file from images of individual cards, double-sided cards are also supported.

## Requirements

- Python 3.11 or newer (the code uses `X | Y` type annotations evaluated at import).
- Install the dependencies with:

  ```bash
  pip install -r requirements.txt
  ```

## Installation

Run the program from a checkout with `python PnP2PDF.py` or `python -m pnp2pdf`.

`pip install .` installs the package and two entry points: `pnp2pdf` opens the
graphical window and `pnp2pdf-cli` renders without a window. The app is also
usable as a library-ish package: the core modules (`model`, `cards`, `layout`,
`render`, `pdf`, `job`, `options`) import without the GUI.

Prebuilt downloads:

- [Prebuilt releases](https://github.com/Dannoster/PnP2PDF/releases) for macOS and Windows

Or install the dependencies by hand:

1. [Install Python](https://www.python.org/downloads/)
2. [Install PIP](https://pip.pypa.io/en/stable/installation/)
3. Run `pip install -r requirements.txt` in terminal/command line
4. Open `PnP2PDF.py` file

## Building a standalone app

Build the bundled executable with one command:

```bash
./build_app.sh
```

On Windows use `build_app.bat`. Add `--no-tests` to skip the test run. The
script creates `.venv` and installs or updates the requirements itself.

## Command line

The window can be skipped with the `pnp2pdf-cli` entry point, or with
`python -m pnp2pdf --cards ... --out ...` from a checkout:

```bash
pnp2pdf-cli --cards ~/cards --out ~/pdf
pnp2pdf-cli --cards ~/cards --out ~/pdf --double-sided
```

Exit codes: `0` success, `1` unexpected error, `2` bad input, `130` cancelled.

## Program interface

<img width="584" alt="image" src="https://github.com/Dannoster/PnP2PDF/assets/91663466/76cb986b-e322-4fae-ba88-3925bf03d61b">

## Result

<img width="741" alt="image" src="https://github.com/Dannoster/PnP2PDF/assets/91663466/5b78db15-d9a8-49b6-8c18-ecdcf37a425f">

## Double-sided cards

1. Create folder.
2. Add all card backs you will use to this folder (If you have back template with size bigger than card you may name it like that `Card Back Name_64,5_91,3.png`, i.e. add `_{back width in mm}_{back height in mm}` in name of file, use a comma or a dot to separate the integer and fractional parts of the width/height).
3. Create folders with the same names like card backs from previous step (except you dont need add `_{back width in mm}_{back height in mm}` at the end of folder name) and place all cards that have corresponding back.
4. You can also create folder with name `Double Sided` for double sided cards that have no ordinary back, where you should place all sides together. Each pair of files will be printed as one double sided card. Be sure if files are sorted in alphabetical order.
5. Finally you can add folder named `Single Sided` with cards that don't have back and place your cards here.
6. That is all! Now choose the folder from step 1 as `Cards folder` in `PnP2PDF` window and enjoy your final result!

The final structure is shown below:

<img width="334" alt="image" src="https://github.com/Dannoster/PnP2PDF/assets/91663466/11636a16-23a6-4027-b687-1028d582da59">

P.S. You can add as much card backs and their folders as you want, not only two pair as in example.

## Development

Install the test dependencies and run the suite:

```bash
pip install -r requirements-dev.txt
pytest
```

The suite stubs PySimpleGUI, so no display and no GUI licence are needed.
Settings are stored as `presets.json` in `~/Library/Application Support/PnP2PDF`
on macOS and `%APPDATA%\PnP2PDF` on Windows when running the packaged app; on
Linux they go to `$XDG_DATA_HOME/pnp2pdf` (or `~/.local/share/pnp2pdf`). Running
from source keeps `presets.json` next to the project.

## Known issues

* Windows version shows error after closing a program (doesn't affect work).
* Windows version may cause Microsoft Defender alerts.
