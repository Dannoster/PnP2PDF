import threading
import time

import pytest

from pnp2pdf import preview
from pnp2pdf.errors import InputError
from pnp2pdf.layout import create_sheets
from pnp2pdf.model import Card
from pnp2pdf.ui.preview_controller import (
    PreviewController,
    PreviewResult,
    preview_from_state,
)


def _form(**overrides):
    values = {
        '-CFOLDER-': '',
        '-PDFFOLDER-': 'out',
        '-PWIDTH-': '210',
        '-PHEIGHT-': '297',
        '-CWIDTH-': '62,5',
        '-CHEIGHT-': '88,7',
        '-ROWS-': '3',
        '-COLUMNS-': '3',
        '-GAP-': '2',
        '-HSZONES-': '0',
        '-VSZONES-': '0',
        '-SIZECHECKBOX-': False,
        '-DOUBLECHECKBOX-': False,
    }
    values.update(overrides)
    return values


def _sheets_with(count):
    cards = [Card(f"c{i}.png", 62.5, 88.7) for i in range(count)]
    return create_sheets(cards, 210, 297, 3, 3, 2, 0, 0)


class FakeClock:
    def __init__(self):
        self.t = 0.0

    def __call__(self):
        return self.t

    def advance(self, seconds):
        self.t += seconds


def _wait_idle(controller, timeout=5.0):
    deadline = time.monotonic() + timeout
    while controller.is_running() and time.monotonic() < deadline:
        time.sleep(0.001)


def test_submit_then_due_waits_for_debounce():
    clock = FakeClock()
    controller = PreviewController(compute=lambda v, t: ([], ""),
                                   debounce=0.25, now=clock)

    controller.submit({'-CFOLDER-': 'A'})
    assert controller.due() is False

    clock.advance(0.24)
    assert controller.due() is False

    clock.advance(0.01)
    assert controller.due() is True


def test_start_only_when_due_and_uses_latest_values():
    clock = FakeClock()
    calls = []
    started = threading.Event()

    def compute(values, target_px):
        calls.append((values, target_px))
        started.set()
        return [values['-CFOLDER-']], "summary"

    controller = PreviewController(compute=compute, debounce=0.25, now=clock)
    controller.submit({'-CFOLDER-': 'A'})

    assert controller.start() is False
    clock.advance(0.25)
    assert controller.start() is True
    assert started.wait(5)
    _wait_idle(controller)

    assert calls == [({'-CFOLDER-': 'A'}, 560)]


def test_take_returns_one_result_per_call_and_none_when_empty():
    controller = PreviewController(compute=lambda v, t: ([1, 2], "s"),
                                   debounce=0.0)
    controller.submit({'-CFOLDER-': 'A'})
    assert controller.start() is True
    _wait_idle(controller)

    result = controller.take()
    assert isinstance(result, PreviewResult)
    assert result.images == [1, 2]
    assert result.summary == "s"
    assert result.error is None
    assert controller.take() is None


def test_stale_result_is_dropped_then_new_one_delivered():
    clock = FakeClock()
    calls = []
    a_running = threading.Event()
    release_a = threading.Event()

    def compute(values, target_px):
        name = values['-CFOLDER-']
        calls.append(name)
        if name == 'A':
            a_running.set()
            release_a.wait(5)
        return [name], name

    controller = PreviewController(compute=compute, debounce=0.25, now=clock)
    controller.submit({'-CFOLDER-': 'A'})
    clock.advance(0.25)
    assert controller.start() is True
    assert a_running.wait(5)

    controller.submit({'-CFOLDER-': 'B'})
    release_a.set()
    _wait_idle(controller)

    assert controller.take() is None

    clock.advance(0.25)
    assert controller.start() is True
    _wait_idle(controller)

    result = controller.take()
    assert result is not None
    assert result.images == ['B']
    assert calls == ['A', 'B']


def test_errors_become_preview_result_error():
    def compute(values, target_px):
        raise RuntimeError("boom")

    controller = PreviewController(compute=compute, debounce=0.0)
    controller.submit({'-CFOLDER-': 'A'})
    assert controller.start() is True
    _wait_idle(controller)

    result = controller.take()
    assert result.error == "boom"
    assert result.images == []


def test_stop_prevents_new_work_and_is_safe_when_never_started():
    controller = PreviewController(compute=lambda v, t: ([], ""), debounce=0.0)
    controller.stop()
    controller.submit({'-CFOLDER-': 'A'})

    assert controller.due() is False
    assert controller.start() is False
    controller.stop()


def test_stop_lets_a_running_worker_queue_its_result():
    gate = threading.Event()

    def compute(values, target_px):
        gate.wait(5)
        return [1], "s"

    controller = PreviewController(compute=compute, debounce=0.0)
    controller.submit({'-CFOLDER-': 'A'})
    assert controller.start() is True
    controller.stop()
    gate.set()
    _wait_idle(controller)

    result = controller.take()
    assert result is not None
    assert result.images == [1]


def test_preview_from_state_real_folder(card_folder):
    folder = card_folder(names=[f"{i}.png" for i in range(1, 13)])

    images, summary = preview_from_state(_form(**{'-CFOLDER-': str(folder)}))

    assert len(images) == 2
    assert summary == "12 cards · 2 sheets · 2 pages"


def test_preview_from_state_empty_string_folder_raises_input_error():
    with pytest.raises(InputError):
        preview_from_state(_form(**{'-CFOLDER-': ''}))


def test_preview_from_state_empty_folder_raises_input_error(card_folder):
    folder = card_folder(names=[])

    with pytest.raises(InputError):
        preview_from_state(_form(**{'-CFOLDER-': str(folder)}))


def test_describe_singular_and_plural_forms():
    assert preview.describe(_sheets_with(1), False) == "1 card · 1 sheet · 1 page"
    assert preview.describe(_sheets_with(1), True) == "1 card · 1 sheet · 2 pages"
    assert preview.describe(_sheets_with(9), False) == "9 cards · 1 sheet · 1 page"
    assert preview.describe(_sheets_with(12), False) == "12 cards · 2 sheets · 2 pages"


def test_describe_double_sided_doubles_pages():
    assert preview.describe(_sheets_with(12), True) == "12 cards · 2 sheets · 4 pages"
