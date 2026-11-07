import sys
import types
from pathlib import Path

import pytest
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


class _StubWidget:
    """A dummy widget: subclassable and callable, ignoring every argument."""

    def __init__(self, *args, **kwargs):
        pass

    def __getattr__(self, name):
        return lambda *args, **kwargs: None


class _GuiStub(types.ModuleType):
    """Stands in for customtkinter so tests never need the real library."""

    def __getattr__(self, name):
        return _StubWidget


if "customtkinter" not in sys.modules:
    sys.modules["customtkinter"] = _GuiStub("customtkinter")


@pytest.fixture(scope="session")
def pnp():
    from pnp2pdf import app as pnp_module

    return pnp_module


@pytest.fixture
def log_dir(tmp_path, monkeypatch):
    """Point the crash log at ``tmp_path`` for the duration of a test."""
    from pnp2pdf import resources

    monkeypatch.setattr(resources, "data_dir", lambda *a, **k: str(tmp_path))
    return tmp_path


@pytest.fixture
def card_folder(tmp_path):
    """Factory that writes PNG card images into a folder and returns it."""

    def _card_folder(names, size=(300, 400)):
        folder = tmp_path / "cards"
        folder.mkdir(parents=True, exist_ok=True)
        for index, name in enumerate(names):
            colour = (
                (index * 40) % 256,
                (index * 80) % 256,
                (index * 120) % 256,
            )
            target = folder / name
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.suffix.lower() in (".png", ".jpg", ".jpeg", ".tif", ".tiff"):
                Image.new("RGB", size, colour).save(target)
            else:
                target.write_bytes(b"not an image")
        return folder

    return _card_folder
