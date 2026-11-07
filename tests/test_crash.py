"""Tests for the crash log: writing entries and hooking the interpreter."""

import importlib
import re
import sys

import pytest

import fake_ctk
from pnp2pdf import crash, resources
from pnp2pdf.settings import DEFAULT_SETTINGS

_TIMESTAMP = re.compile(
    r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[pid \d+\] ")


def _read(path):
    return path.read_text(encoding="utf-8")


def test_log_message_writes_one_line(log_dir):
    path = crash.log_message("hello")

    assert path == str(log_dir / crash.LOG_NAME)
    assert (log_dir / crash.LOG_NAME).is_file()
    lines = _read(log_dir / crash.LOG_NAME).splitlines()
    assert len(lines) == 1
    assert lines[0].endswith("hello")
    assert _TIMESTAMP.match(lines[0])


def test_log_exception_writes_the_traceback(log_dir):
    try:
        raise ValueError("boom")
    except ValueError:
        crash.log_exception("a test", *sys.exc_info())

    text = _read(log_dir / crash.LOG_NAME)
    assert "a test" in text
    assert "ValueError" in text
    assert "boom" in text
    assert "Traceback (most recent call last)" in text


def test_entries_are_appended(log_dir):
    crash.log_message("first")
    crash.log_message("second")

    lines = _read(log_dir / crash.LOG_NAME).splitlines()
    assert len(lines) == 2
    assert lines[0].endswith("first")
    assert lines[1].endswith("second")


def test_log_event_writes_one_line(log_dir):
    crash.log_event("click", button="create")

    lines = _read(log_dir / crash.LOG_NAME).splitlines()
    assert len(lines) == 1
    assert lines[0].endswith("click button=create")
    assert _TIMESTAMP.match(lines[0])


def test_log_event_flattens_newlines(log_dir):
    crash.log_event("folder", path="a\nb\tc")

    lines = _read(log_dir / crash.LOG_NAME).splitlines()
    assert len(lines) == 1
    assert lines[0].endswith("folder path=a b c")


def test_log_rotates_at_the_size_limit(log_dir):
    path = log_dir / crash.LOG_NAME
    path.write_bytes(b"x" * crash.MAX_BYTES)

    crash.log_event("click", button="create")

    assert len(path.read_bytes()) < crash.MAX_BYTES
    backup = log_dir / (crash.LOG_NAME + crash.BACKUP_SUFFIX)
    assert backup.is_file()
    assert backup.read_bytes() == b"x" * crash.MAX_BYTES
    assert crash.log_path() == str(path)


def test_install_handlers_returns_the_path_and_is_safe_twice(log_dir):
    first = crash.install_handlers()
    second = crash.install_handlers()

    assert first == str(log_dir / crash.LOG_NAME)
    assert second == first
    assert sys.excepthook is not sys.__excepthook__


def test_unwritable_directory_does_not_raise(tmp_path, monkeypatch):
    blocker = tmp_path / "blocker"
    blocker.write_text("not a directory", encoding="utf-8")
    monkeypatch.setattr(resources, "data_dir",
                        lambda *a, **k: str(blocker / "sub"))

    assert crash.log_message("hello") == str(blocker / "sub" / crash.LOG_NAME)
    try:
        raise RuntimeError("nested")
    except RuntimeError:
        crash.log_exception("nowhere", *sys.exc_info())


@pytest.fixture
def gui():
    fake_ctk.install()
    module = importlib.import_module("pnp2pdf.ui.window")
    return importlib.reload(module)


def test_tk_callback_exception_is_logged(gui, log_dir, monkeypatch):
    shown = []
    monkeypatch.setattr(gui.PnP2PDFWindow, "_show_error",
                        lambda self, message: shown.append(message))
    window = gui.PnP2PDFWindow(dict(DEFAULT_SETTINGS))
    try:
        try:
            raise RuntimeError("callback boom")
        except RuntimeError:
            window.report_callback_exception(*sys.exc_info())
    finally:
        window.destroy()

    text = _read(log_dir / crash.LOG_NAME)
    assert "tk callback" in text
    assert "RuntimeError" in text
    assert "callback boom" in text
    assert shown
