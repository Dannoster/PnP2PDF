@echo off
setlocal

echo ==^> Moving to the repository root
cd /d "%~dp0"

if not defined PYTHON set "PYTHON=python"
set "PY=%PYTHON%"

echo ==^> Checking that %PY% is Python 3.11 or newer
"%PY%" -c "import sys; print(sys.version_info >= (3, 11))" | findstr /i "True" >nul
if errorlevel 1 (
    echo Error: Python 3.11 or newer is required, but "%PY%" is older. 1>&2
    echo Set PYTHON=C:\path\to\python.exe ^(3.11 or newer^) and try again. 1>&2
    exit /b 1
)

set "NO_TESTS="
for %%A in (%*) do (
    if /i "%%A"=="--no-tests" set "NO_TESTS=1"
)

echo ==^> Creating the virtual environment in .venv
if not exist ".venv" (
    "%PY%" -m venv .venv
    if errorlevel 1 exit /b 1
)

echo ==^> Upgrading pip and installing/updating the requirements
".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
if errorlevel 1 exit /b 1
".venv\Scripts\python.exe" -m pip install -r requirements.txt -r requirements-dev.txt
if errorlevel 1 exit /b 1

if not defined NO_TESTS (
    echo ==^> Running the test suite
    ".venv\Scripts\python.exe" -m pytest -q
    if errorlevel 1 exit /b 1
)

if defined PNP2PDF_BUILD_DRY (
    echo ==^> Dry run: PyInstaller command that would run
    echo .venv\Scripts\pyinstaller --noconfirm --clean --windowed --name PnP2PDF --add-data "pnp2pdf/color_profiles/sRGB-IEC61966-2.1.icc;color_profiles" --add-data "pnp2pdf/color_profiles/USWebCoatedSWOP.icc;color_profiles" --add-data "pnp2pdf/icon.png;." --icon=pnp2pdf/icon.png PnP2PDF.py
    exit /b 0
)

echo ==^> Clearing build\ and dist\
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo ==^> Running PyInstaller
".venv\Scripts\pyinstaller.exe" --noconfirm --clean --windowed --name PnP2PDF ^
    --add-data "pnp2pdf/color_profiles/sRGB-IEC61966-2.1.icc;color_profiles" ^
    --add-data "pnp2pdf/color_profiles/USWebCoatedSWOP.icc;color_profiles" ^
    --add-data "pnp2pdf/icon.png;." ^
    --icon=pnp2pdf/icon.png ^
    PnP2PDF.py
if errorlevel 1 exit /b 1

echo ==^> Built dist\PnP2PDF\PnP2PDF.exe

endlocal
