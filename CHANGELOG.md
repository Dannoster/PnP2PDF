# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [3.1.0] - 2026-11-07

### Added
- the window is now built with CustomTkinter and shows a live preview panel of the sheets being assembled
- the app bundle now carries a real version number instead of `0.0.0`
- crashes are written to a log file in the application data directory

### Changed
- the window no longer depends on PySimpleGUI; `customtkinter` replaces it
- the Files section now names both folders and the output folder is optional (empty means next to the cards)
- the window stays open after a build, the progress bar completes, and the finished PDF can be revealed in the file manager

### Documentation
- the README explains every card folder layout with a tree, the settings and the
  command line, and warns that the "Double sided cards" box is what prints backs
- the README is translated into Russian, Simplified Chinese, Spanish and German

### Fixed
- the macOS bundle is re-signed after its version is written, so it no longer reports "is damaged" on launch
- the build script detects and repairs a virtual environment whose packages cannot be imported
- the build script rebuilds a virtual environment that holds packages or scripts from another platform

## [3.0.0] - 2026-11-07

### Added
- pytest suite with GUI stubs
- GitHub Actions on Linux/macOS/Windows for Python 3.11-3.13
- GitHub Actions release workflow: pushing a `v*` tag publishes macOS and Windows builds to GitHub Releases
- `build_app.sh` / `build_app.bat` build the bundle from a working copy in one command
- MIT licence
- `pnp2pdf-cli` command line
- progress indicator and a Cancel button
- pull request and commit templates

### Changed
- the app is an installable package with a GUI-free core (`model`, `cards`, `layout`, `render`, `pdf`, `job`, `options`)
- settings live in the platform data directory (`~/Library/Application Support/PnP2PDF` on macOS, `%APPDATA%` on Windows) instead of next to the code
- settings are `presets.json` instead of a pickle
- PDF pages are assembled in memory
- dependencies require Pillow 13, PySimpleGUI 6.4, pytest 9.2 and PyInstaller 6.22; Python 3.10 is no longer supported (Pillow 13 needs 3.11)

### Fixed
- empty or wrong cards folder now reports which extensions are accepted instead of crashing
- odd file counts in double-sided mode are rejected with a message
- repeated runs no longer pick up the previous PDF
- `0` rows or columns no longer divide by zero
- ICC profiles are found regardless of the working directory
- cards are ordered 1, 2, ..., 10 instead of 1, 10, 11, 2
- paths work on Windows
- the four cut marks are now identical
- `Card_64.5_91.3.png` parses as 64.5 x 91.3

### Removed
- the disk round-trip through temporary `_page_*.jpeg` files

[3.1.0]: https://github.com/Dannoster/PnP2PDF/releases/tag/v3.1.0
[3.0.0]: https://github.com/Dannoster/PnP2PDF/releases/tag/v3.0.0
