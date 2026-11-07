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

venv_python() { printf '%s' ".venv/bin/python"; }

venv_is_usable() {
    [ -x "$(venv_python)" ] && "$(venv_python)" -c 'import sys' >/dev/null 2>&1
}

# macOS can hand us a venv whose interpreter runs under Rosetta while pip
# fetched arm64 wheels (or the other way round); such a venv can never import
# compiled packages, so it is rebuilt instead of repaired.
venv_arch_matches_host() {
    [ "$(uname -s)" = "Darwin" ] || return 0
    local host venv
    host="$(uname -m)"
    venv="$("$(venv_python)" -c 'import platform; print(platform.machine())' 2>/dev/null || echo unknown)"
    [ "$host" = "$venv" ]
}

create_venv() {
    echo "==> Creating the virtual environment in .venv"
    rm -rf .venv
    "$PY" -m venv .venv
    "$(venv_python)" -m pip install --quiet --upgrade pip
}

install_requirements() {
    echo "==> Installing/updating the requirements"
    "$(venv_python)" -m pip install --quiet -r requirements.txt -r requirements-dev.txt
}

check_imports() {
    "$(venv_python)" - <<'PY'
import importlib, sys

modules = ("PIL.Image", "customtkinter", "pytest", "PyInstaller")

try:
    import tkinter  # noqa: F401
    has_tkinter = True
except (ImportError, OSError):
    has_tkinter = False
    print("not checking customtkinter: this system has no tkinter", file=sys.stderr)

broken = []
for module in modules:
    if module == "customtkinter" and not has_tkinter:
        continue
    try:
        importlib.import_module(module)
    except Exception:                # noqa: BLE001 - any failure means "broken"
        broken.append(module)

for module in broken:
    print(module)
if broken:
    sys.exit(1)
PY
}

dist_for_module() {
    case "$1" in
        PIL.Image) printf '%s' "Pillow" ;;
        customtkinter) printf '%s' "customtkinter" ;;
        pytest) printf '%s' "pytest" ;;
        PyInstaller) printf '%s' "pyinstaller" ;;
        *) printf '%s' "$1" ;;
    esac
}

repair_imports() {
    broken="$(check_imports 2>/dev/null)" || true
    [ -n "$broken" ] || return 0
    for module in $broken; do
        dist="$(dist_for_module "$module")"
        echo "==> $dist cannot be imported ($module) — reinstalling it"
        "$(venv_python)" -m pip install --quiet --force-reinstall --no-cache-dir "$dist"
    done
    check_imports
}

if ! venv_is_usable || ! venv_arch_matches_host; then
    create_venv
fi
install_requirements

if ! repair_imports; then
    echo "==> The environment still cannot import its packages — rebuilding .venv from scratch"
    create_venv
    install_requirements
    if ! repair_imports; then
        echo "Error: pip installed the requirements but they still cannot be imported." >&2
        echo "Delete .venv and run this script again; if it keeps failing, report it with the output above." >&2
        exit 1
    fi
fi

echo "==> Checking installed packages for conflicts"
"$(venv_python)" -m pip check || true

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
