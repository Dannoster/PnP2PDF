"""Debounced background previews, with no GUI imports so they can be tested.

The window feeds widget values in with :meth:`PreviewController.submit` and
polls for finished pictures with :meth:`PreviewController.take`. Everything
here is plain Python: no ``customtkinter``, no ``tkinter``, no PySimpleGUI.
"""

import queue
import threading
import time
from dataclasses import dataclass

from pnp2pdf.cards import create_cards
from pnp2pdf.errors import InputError
from pnp2pdf.layout import create_sheets
from pnp2pdf.options import JobOptions
from pnp2pdf.preview import describe, preview_pages


@dataclass(frozen=True)
class PreviewResult:
    images: list
    summary: str
    error: str | None
    generation: int
    elapsed_ms: float


def preview_from_state(values: dict, target_px: int = 560) -> tuple[list, str]:
    """Turn widget values into ``(images, summary)`` without writing files.

    ``InputError`` (and anything else the pipeline raises) is left for the
    caller to report; card/layout problems are translated to ``InputError``
    because they come straight from user input.
    """
    options = JobOptions.from_form(values)
    try:
        cards = create_cards(options.cards_folder, options.card_width,
                             options.card_height, options.smaller_size,
                             options.double_sided)
        sheets = create_sheets(cards, options.paper_width, options.paper_height,
                               options.rows, options.columns, options.gap,
                               options.ud_szone, options.lr_szone)
    except InputError:
        raise
    except ValueError as error:
        raise InputError(str(error)) from error
    images = preview_pages(sheets, target_px)
    return images, describe(sheets, options.double_sided)


class PreviewController:
    """Runs ``compute(values, target_px)`` off-thread, at most one at a time."""

    def __init__(self, compute=None, debounce=0.25, now=time.monotonic):
        self._compute = compute if compute is not None else preview_from_state
        self._debounce = debounce
        self._now = now
        self._lock = threading.Lock()
        self._results = queue.Queue()
        self._values = None
        self._last_submit = None
        self._pending = False
        self._running = False
        self._stopped = False
        self._generation = 0

    def submit(self, values) -> None:
        """Remember the latest values; never blocks and never computes."""
        with self._lock:
            self._values = dict(values)
            self._last_submit = self._now()
            self._pending = True
            self._generation += 1

    def due(self) -> bool:
        """True once the debounce has elapsed with nothing already running."""
        with self._lock:
            if not self._pending or self._running or self._stopped:
                return False
            return self._now() - self._last_submit >= self._debounce

    def start(self) -> bool:
        """Launch a worker if ``due()``; returns whether one was started."""
        with self._lock:
            if self._stopped or not self._pending or self._running:
                return False
            if self._now() - self._last_submit < self._debounce:
                return False
            values = self._values
            generation = self._generation
            self._pending = False
            self._running = True
        thread = threading.Thread(target=self._work, args=(values, generation),
                                  daemon=True)
        thread.start()
        return True

    def _work(self, values, generation) -> None:
        started = time.perf_counter()
        try:
            images, summary = self._compute(values, 560)
            result = PreviewResult(images=list(images), summary=summary,
                                   error=None, generation=generation,
                                   elapsed_ms=(time.perf_counter() - started) * 1000)
        except Exception as error:  # noqa: BLE001 - reported through the result
            result = PreviewResult(images=[], summary="", error=str(error),
                                   generation=generation,
                                   elapsed_ms=(time.perf_counter() - started) * 1000)
        self._results.put(result)
        with self._lock:
            self._running = False

    def take(self):
        """Pop the newest fresh result, dropping any that are stale."""
        while True:
            try:
                result = self._results.get_nowait()
            except queue.Empty:
                return None
            with self._lock:
                newest = self._generation
            if result.generation < newest:
                continue
            return result

    def stop(self) -> None:
        """Forbid new work; a running worker may still queue its result."""
        with self._lock:
            self._stopped = True

    def is_running(self) -> bool:
        with self._lock:
            return self._running
