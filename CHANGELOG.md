# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [3.0.0] - 2026-11-07

### Added
- pytest suite with GUI stubs
- GitHub Actions on Linux/macOS/Windows for Python 3.10-3.12
- MIT licence
- `pnp2pdf-cli` command line
- progress indicator and a Cancel button
- pull request and commit templates

### Changed
- the app is an installable package with a GUI-free core (`model`, `cards`, `layout`, `render`, `pdf`, `job`, `options`)
- settings live in the platform data directory (`~/Library/Application Support/PnP2PDF` on macOS, `%APPDATA%` on Windows) instead of next to the code
- settings are `presets.json` instead of a pickle
- PDF pages are assembled in memory

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

[3.0.0]: https://github.com/Dannoster/PnP2PDF/releases/tag/v3.0.0
