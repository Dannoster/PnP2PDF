"""Smoke tests for the real ``PnP2PDFWindow`` driven through a fake toolkit.

The fake is installed before the window module is (re)loaded so the genuine
window code runs without CustomTkinter, Tk, or a display.
"""

import importlib
import time

import pytest
from PIL import Image

import fake_ctk
from pnp2pdf.settings import DEFAULT_SETTINGS
from pnp2pdf.ui.form import FIELD_KEYS


@pytest.fixture(autouse=True)
def _isolated_log(log_dir):
    """Every window test logs into tmp_path, never the real data directory."""
    return log_dir


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


def test_the_picture_is_never_cleared_with_none(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    try:
        _pump(window, lambda: len(window._page_views) == 2,
              what="two preview pages")
        picture = window._picture
        seen = []
        original = picture.configure

        def record(**kwargs):
            original(**kwargs)
            if "image" in kwargs:
                seen.append(kwargs["image"])

        picture.configure = record
        stale = list(window._page_views) + list(window._page_thumbs)

        window._clear_preview()

        assert seen, "the preview label was never reconfigured"
        assert all(isinstance(image, module.ctk.CTkImage) for image in seen)
        assert picture.image is seen[-1]
        assert picture.image not in stale
    finally:
        window.destroy()


def test_a_failing_preview_update_is_logged_not_raised(
        gui, card_folder, tmp_path):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    picture = window._picture
    original = picture.configure

    def boom(**kwargs):
        raise RuntimeError('image "pyimage2" doesn\'t exist')

    try:
        _pump(window, lambda: len(window._page_views) == 2,
              what="two preview pages")
        picture.configure = boom
        window._show_page(0)
    finally:
        picture.configure = original
        window.destroy()

    text = (tmp_path / module.crash.LOG_NAME).read_text(encoding="utf-8")
    assert "ui-failed" in text


def test_the_preview_stops_after_three_failed_redraws(
        gui, card_folder, tmp_path):
    module, handles = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    picture = window._picture
    original = picture.configure

    def boom(**kwargs):
        raise RuntimeError('image "pyimage2" doesn\'t exist')

    def ui_failed_lines():
        text = (tmp_path / module.crash.LOG_NAME).read_text(encoding="utf-8")
        return [line for line in text.splitlines() if "ui-failed" in line]

    try:
        _pump(window, lambda: len(window._page_views) == 2,
              what="two preview pages")
        picture.configure = boom
        window._schedule_preview()

        _pump(window, lambda: window._ui_failures >= 3,
              what="three failed redraws")
        _pump(window, lambda: not window._preview.is_running()
              and not window._preview._pending,
              what="the preview queue to settle")
        window.update()

        assert window._ui_failures == 3
        lines = ui_failed_lines()
        assert len(lines) == 3, lines
        for attempt, line in enumerate(lines, start=1):
            assert f"attempt={attempt}" in line
        assert "Preview unavailable" in window._preview_status.text

        other = tmp_path / "more-cards"
        other.mkdir()
        Image.new("RGB", (300, 400), (20, 40, 60)).save(other / "1.png")
        handles["filedialog"].directory = str(other)
        window._choose_folder("-CFOLDER-")

        assert window._ui_failures == 0, "a new folder re-arms the preview"
        _pump(window, lambda: window._ui_failures >= 3,
              what="three more failed redraws")
        _pump(window, lambda: not window._preview.is_running()
              and not window._preview._pending,
              what="the preview queue to settle again")
        window.update()

        lines = ui_failed_lines()
        assert len(lines) == 6, lines
        for attempt, line in enumerate(lines[3:], start=1):
            assert f"attempt={attempt}" in line
    finally:
        picture.configure = original
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

        _pump(window, lambda: window._last_pdf, timeout=20.0,
              what="the build to finish")

        assert window.destroyed is False
        assert window._progress.value == 1
        assert window._status.text
        assert window._last_pdf.suffix == ".pdf"
        assert window._last_pdf.exists()
        assert window._last_pdf.name in window._status.text
        assert window._show_button.state == "normal"
        assert window._run_button.state == "normal"
        assert window._cancel_button.state == "disabled"
        assert isinstance(window.result, dict)
        for key in FIELD_KEYS:
            assert key in window.result
        assert list(output.rglob("*.pdf")), "no PDF was written"
    finally:
        window.on_close()


def test_create_pdf_logs_the_click(gui, card_folder, tmp_path):
    module, _ = gui
    folder = _cards(card_folder)
    output = tmp_path / "pdfs"
    output.mkdir()
    window = _window(module, folder, save_to=str(output))
    try:
        _pump(window, lambda: len(window._page_views) == 2)
        window._run_button.invoke()

        text = (tmp_path / module.crash.LOG_NAME).read_text(encoding="utf-8")
        assert "click button=create" in text
    finally:
        window.on_close()


def test_show_pdf_without_a_finished_build_does_nothing(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    try:
        assert window._last_pdf is None
        assert window._show_button.state == "disabled"

        window._reveal_pdf()

        assert window.destroyed is False
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


def test_closing_logs_a_single_window_close_line(gui, card_folder, tmp_path):
    module, _ = gui
    folder = _cards(card_folder, count=1)
    window = _window(module, folder)

    window._on_close()
    window._close()

    text = (tmp_path / module.crash.LOG_NAME).read_text(encoding="utf-8")
    assert text.count("window-close") == 1


def test_an_empty_folder_shows_an_invitation(gui):
    module, _ = gui
    window = _window(module, "")
    try:
        window.update()

        assert window._preview_status.text == (
            "Choose a cards folder to see the preview.")
        assert window._preview_status.cget("text_color") == module._GREY
        assert isinstance(window._picture.image, module.ctk.CTkImage)
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
        assert isinstance(window._picture.image, module.ctk.CTkImage)
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


def test_the_output_folder_is_hidden_by_default(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    try:
        assert window._output_row.packed is False
        assert window._output_hint.packed is True
        assert window._output_switch.get() is False
    finally:
        window.destroy()


def test_toggling_the_output_switch_shows_and_hides_the_row(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    try:
        window._output_switch.select()
        window._toggle_output_folder()

        assert window._output_row.packed is True
        assert window._output_hint.packed is False

        window._output_switch.deselect()
        window._toggle_output_folder()

        assert window._output_row.packed is False
        assert window._output_hint.packed is True
    finally:
        window.destroy()


def test_an_existing_output_folder_shows_the_row_ticked(gui, card_folder, tmp_path):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder, save_to=str(tmp_path))
    try:
        assert window._output_row.packed is True
        assert window._output_hint.packed is False
        assert window._output_switch.get() is True
    finally:
        window.destroy()


def test_files_section_names_both_folders(gui, card_folder):
    module, _ = gui
    folder = _cards(card_folder)
    window = _window(module, folder)
    try:
        assert fake_ctk.find_by_text(window, "Cards folder") is not None
        assert fake_ctk.find_by_text(window, "Save the PDF to") is not None
        assert window._entries["-CFOLDER-"].cget(
            "placeholder_text") == "Where the card images are"
        assert window._entries["-PDFFOLDER-"].cget(
            "placeholder_text") == "Folder for the PDF"
    finally:
        window.destroy()


def test_pdf_file_picks_the_pdf(gui, tmp_path):
    module, _ = gui
    window = _window(module, None)
    try:
        folder = tmp_path / "out"
        folder.mkdir()
        pdf = folder / "cards.pdf"
        pdf.write_bytes(b"%PDF-1.4")

        assert window._pdf_file(folder) == pdf

        empty = tmp_path / "empty"
        empty.mkdir()
        assert window._pdf_file(empty) is None

        assert window._pdf_file(pdf) == pdf
    finally:
        window.destroy()
