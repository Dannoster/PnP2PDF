#!/usr/bin/env bash
set -eu

echo "==> Moving to the repository root"
cd "$(dirname "$0")"

NO_TESTS=0
for arg in "$@"; do
    case "$arg" in
        --no-tests) NO_TESTS=1 ;;
        *)
            echo "Unknown option: $arg" >&2
            exit 2
            ;;
    esac
done

PY="${PYTHON:-python3}"

echo "==> Checking that $PY is Python 3.11 or newer"
if [ "$("$PY" -c 'import sys; print(sys.version_info >= (3, 11))')" != "True" ]; then
    echo "Error: Python 3.11 or newer is required, but '$PY' is older." >&2
    echo "Set PYTHON=/path/to/python3.11 (or newer) and try again." >&2
    exit 1
fi

echo "==> Creating the virtual environment in .venv"
[ -x .venv/bin/python ] && .venv/bin/python -c "" 2>/dev/null || { rm -rf .venv; "$PY" -m venv .venv; }

echo "==> Upgrading pip and installing/updating the requirements"
.venv/bin/python -m pip install --quiet --upgrade pip
.venv/bin/python -m pip install -r requirements.txt -r requirements-dev.txt

if [ "$NO_TESTS" -eq 0 ]; then
    echo "==> Running the test suite"
    .venv/bin/python -m pytest -q
fi

PYINSTALLER_ARGS=(
    --noconfirm
    --clean
    --windowed
    --name PnP2PDF
    --collect-data customtkinter
    --collect-submodules customtkinter
    --add-data "pnp2pdf/color_profiles/sRGB-IEC61966-2.1.icc:color_profiles"
    --add-data "pnp2pdf/color_profiles/USWebCoatedSWOP.icc:color_profiles"
    --add-data "pnp2pdf/icon.png:."
    --icon=pnp2pdf/icon.png
)
if [ "$(uname -s)" = "Darwin" ]; then
    PYINSTALLER_ARGS+=(--osx-bundle-identifier=com.dannoster.pnp2pdf)
fi

if [ -n "${PNP2PDF_BUILD_DRY:-}" ]; then
    echo "==> Dry run: PyInstaller command that would run"
    echo ".venv/bin/pyinstaller ${PYINSTALLER_ARGS[*]} PnP2PDF.py"
    exit 0
fi

echo "==> Clearing build/ and dist/"
rm -rf build dist

echo "==> Running PyInstaller"
.venv/bin/pyinstaller "${PYINSTALLER_ARGS[@]}" PnP2PDF.py

if [ "$(uname -s)" = "Darwin" ]; then
    echo "==> Setting the bundle version from pyproject.toml"
    VERSION="$("$PY" -c "import tomllib; print(tomllib.load(open('pyproject.toml', 'rb'))['project']['version'])")"
    plutil -replace CFBundleShortVersionString -string "$VERSION" \
        dist/PnP2PDF.app/Contents/Info.plist
    plutil -replace CFBundleVersion -string "$VERSION" \
        dist/PnP2PDF.app/Contents/Info.plist
    echo "==> Re-signing the bundle ad hoc"
    codesign --force --sign - "dist/PnP2PDF.app"
    if ! codesign --verify --deep --strict "dist/PnP2PDF.app"; then
        echo "codesign could not re-sign the bundle" >&2
        exit 1
    fi
    echo "==> Built dist/PnP2PDF.app ($VERSION)"
    echo "    The app is ad-hoc signed. A browser-downloaded copy needs"
    echo "    xattr -dr com.apple.quarantine /Applications/PnP2PDF.app"
    echo "    before the first launch."
    echo "    Launch it with: open dist/PnP2PDF.app"
else
    echo "==> Built dist/PnP2PDF/PnP2PDF"
fi
