"""Smoke tests for the real ``PnP2PDFWindow`` driven through a fake toolkit.

The fake is installed before the window module is (re)loaded so the genuine
window code runs without CustomTkinter, Tk, or a display.
"""

import importlib
import time

import pytest

import fake_ctk
from pnp2pdf.settings import DEFAULT_SETTINGS
from pnp2pdf.ui.form import FIELD_KEYS


@pytest.fixture
def gui():
    handles = fake_ctk.install()
    module = importlib.import_module("pnp2pdf.ui.window")
    module = importlib.reload(module)
    return module, handles


def _window(module, folder, **overrides):
    settings = dict(DEFAULT_SETTINGS)
    if folder is not None:
        settings["cards_folder"] = str(folder)
    settings.update(overrides)
    return module.PnP2PDFWindow(settings)


def _cards(card_folder, count=12):
    return card_folder(names=[f"{index}.png" for index in range(1, count + 1)])


def _pump(window, predicate, timeout=5.0, what="the expected state"):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        window.update()
        if predicate():
            return
        time.sleep(0.01)
    raise AssertionError(f"timed out waiting for {what}")


def test_preview_appears_after_the_folder_is_chosen(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    try:
        _pump(window, lambda: len(window._page_views) == 2,
              what="two preview pages")

        assert len(window._page_views) == 2
        assert window._summary.text == "12 cards · 2 sheets · 2 pages"
        assert window._page_label.text == "1 / 2"

        image = window._picture.image
        assert isinstance(image, module.ctk.CTkImage)
        width, height = image.size
        assert width <= 348 and height <= 470, (width, height)
    finally:
        window.destroy()


def test_page_navigation(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    try:
        _pump(window, lambda: len(window._page_views) == 2)

        window._next_button.invoke()
        assert window._page_label.text == "2 / 2"

        window._prev_button.invoke()
        assert window._page_label.text == "1 / 2"
    finally:
        window.destroy()

    for extra in sorted(folder.glob("*.png"))[1:]:
        extra.unlink()
    single = _window(module, folder)
    try:
        _pump(single, lambda: len(single._page_views) == 1,
              what="a single-page preview")

        assert single._page_label.text == "1 / 1"
        assert single._prev_button.state == "disabled"
        assert single._next_button.state == "disabled"
    finally:
        single.destroy()


def test_clicking_a_thumbnail_switches_the_page(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    try:
        _pump(window, lambda: len(window._thumb_buttons) == 2,
              what="two thumbnail buttons")

        window._thumb_buttons[1].invoke()
        assert window._page_index == 1
        assert window._page_label.text == "2 / 2"
    finally:
        window.destroy()


def test_typing_in_a_field_re_renders_the_preview(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder, rows=4)
    try:
        _pump(window, lambda: len(window._page_views) == 1,
              what="the first preview")
        assert window._summary.text == "12 cards · 1 sheet · 1 page"

        entry = window._entries["-ROWS-"]
        entry.delete(0, "end")
        entry.insert(0, "2")
        entry.bindings["<KeyRelease>"](None)

        _pump(window,
              lambda: window._summary.text == "12 cards · 2 sheets · 2 pages",
              what="a re-rendered preview")

        assert window._summary.text == "12 cards · 2 sheets · 2 pages"
        assert len(window._page_views) == 2
    finally:
        window.destroy()


def test_build_from_the_window(gui, card_folder, tmp_path):
    module, _ = gui
    folder = _cards(card_folder)
    output = tmp_path / "pdfs"
    output.mkdir()
    window = _window(module, folder, save_to=str(output))
    try:
        _pump(window, lambda: len(window._page_views) == 2)
        window._run_button.invoke()

        _pump(window, lambda: window.destroyed, timeout=20.0,
              what="the build to finish")

        assert isinstance(window.result, dict)
        for key in FIELD_KEYS:
            assert key in window.result
        assert list(output.rglob("*.pdf")), "no PDF was written"
    finally:
        window.destroy()


def test_bad_input_shows_a_message_box_and_keeps_the_window_alive(
        gui, card_folder, tmp_path):
    module, handles = gui
    folder = _cards(card_folder)
    window = _window(module, folder, save_to=str(tmp_path))
    try:
        entry = window._entries["-ROWS-"]
        entry.delete(0, "end")
        entry.insert(0, "0")

        window._run_button.invoke()

        calls = handles["messagebox"].showerror_calls
        assert calls, "messagebox.showerror was not called"
        title, message = calls[-1]
        assert message
        assert window.destroyed is False
        assert window._run_button.state == "normal"
    finally:
        window.destroy()


def test_closing_the_window_stops_the_controller(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder, count=1)
    window = _window(module, folder)
    try:
        assert window.on_close is not None

        window.on_close()

        assert window.result == "Window closed"
        assert window._preview._stopped is True
        assert window.destroyed is True
    finally:
        window.destroy()


def test_an_empty_folder_shows_an_invitation(gui):
    module, _ = gui
    window = _window(module, "")
    try:
        window.update()

        assert window._preview_status.text == (
            "Choose a cards folder to see the preview.")
        assert window._preview_status.cget("text_color") == module._GREY
        assert window._picture.image is None
        assert window._summary.text == ""
        assert window._page_views == []

        assert window._preview.is_running() is False
        assert window._preview._pending is False
        assert window._preview._results.empty()
    finally:
        window.destroy()


def test_a_folder_without_images_shows_a_red_error(gui, tmp_path):
    module, _ = gui
    empty = tmp_path / "cards"
    empty.mkdir()
    window = _window(module, empty)
    try:
        _pump(window,
              lambda: window._preview_status.cget("text_color") == module._ERROR,
              what="the error message")

        assert window._preview_status.text
        assert window._picture.image is None
        assert window._page_views == []
    finally:
        window.destroy()


def test_create_pdf_stays_enabled_after_a_failed_preview(gui, tmp_path):
    module, _ = gui
    empty = tmp_path / "cards"
    empty.mkdir()
    window = _window(module, empty)
    try:
        _pump(window,
              lambda: window._preview_status.cget("text_color") == module._ERROR,
              what="the failed preview")

        assert window._run_button.state == "normal"
    finally:
        window.destroy()
