import os
import sys
from pathlib import Path

from pnp2pdf import resources


def test_source_resources_resolve():
    assert os.path.exists(resources.resource_path("icon.png"))
    assert os.path.exists(resources.icc_profile("sRGB-IEC61966-2.1.icc"))


def test_frozen_macos_layout(tmp_path, monkeypatch):
    frameworks = tmp_path / "Frameworks"
    macos = tmp_path / "MacOS"
    resources_dir = tmp_path / "Resources"
    for directory in (frameworks, macos, resources_dir):
        directory.mkdir()
    profile_dir = resources_dir / "color_profiles"
    profile_dir.mkdir()
    dummy = profile_dir / "dummy.icc"
    dummy.write_bytes(b"icc")

    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(frameworks), raising=False)
    monkeypatch.setattr(sys, "executable", str(macos / "PnP2PDF"))

    found = resources.resource_path(os.path.join("color_profiles", "dummy.icc"))

    assert Path(found) == dummy


def test_data_dir_platforms(monkeypatch):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    home = os.path.join(os.sep, "home", "user")

    assert resources.data_dir("darwin", home, {}) == os.path.join(
        home, "Library", "Application Support", "PnP2PDF")
    appdata = os.path.join(os.sep, "appdata")
    assert resources.data_dir("win32", home, {"APPDATA": appdata}) == os.path.join(
        appdata, "PnP2PDF")
    assert resources.data_dir("win32", home, {}) == os.path.join(
        home, "AppData", "Roaming", "PnP2PDF")
    assert resources.data_dir("linux", home, {}) == os.path.join(
        home, ".local", "share", "pnp2pdf")


def test_settings_path_creates_directory(tmp_path, monkeypatch):
    target = tmp_path / "missing" / "config"
    monkeypatch.setattr(resources, "data_dir", lambda *a, **k: str(target))

    result = resources.settings_path()

    assert result == str(target / "presets.json")
    assert target.is_dir()
