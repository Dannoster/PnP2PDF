"""Small, fast pictures of pages.

This module deliberately has no GUI imports: the window can ask it for a
preview without paying for (or depending on) PySimpleGUI.
"""

import math
import os

from PIL import Image

from pnp2pdf.render import render_back, render_front

MAX_SIDE_PX = 1600          # preview images are never wider or taller
MAX_PIXELS = 2_000_000      # nor bigger by area

_cache: dict = {}


def fit_px_per_mm(width_mm, height_mm, target_px=560) -> float:
    """Pixels per millimetre so the longer side is ``target_px``.

    Never exceeds ``MAX_SIDE_PX`` on a side or ``MAX_PIXELS`` by area, and is
    always finite and positive, even for absurd or non-positive paper sizes.
    """
    longest = max(width_mm, height_mm)
    if longest <= 0 or target_px <= 0:
        return 1.0
    px_per_mm = target_px / longest
    px_per_mm = min(px_per_mm, MAX_SIDE_PX / longest)
    area = width_mm * height_mm
    if area > 0:
        px_per_mm = min(px_per_mm, math.sqrt(MAX_PIXELS / area))
    return float(px_per_mm)


def _cached_loader(card, size):
    """Load a card picture, reusing and downscaling decoded images."""
    path = str(card.dir)
    size = tuple(size)
    key = (path, os.path.getmtime(path), size)
    cached = _cache.get(key)
    if cached is not None:
        return cached

    image = Image.open(path)
    if path.lower().endswith((".jpg", ".jpeg")):
        image.draft("RGB", size)
    image = image.resize(size, Image.LANCZOS)
    image = image.convert("RGB")
    _cache[key] = image
    return image


def _has_backs(sheet) -> bool:
    return any(getattr(card, "back", None) for card in sheet.cards)


def preview_pages(sheets, target_px=560, limit=8) -> list[Image.Image]:
    """Front pages for every sheet, plus backs, at most ``limit`` images."""
    images = []
    for sheet in sheets:
        if len(images) >= limit:
            break
        px_per_mm = fit_px_per_mm(sheet.width, sheet.height, target_px)
        images.append(render_front(sheet, px_per_mm, mode="RGB", profile=False,
                                   low_detail=True, loader=_cached_loader))
        if len(images) >= limit:
            break
        if _has_backs(sheet):
            images.append(render_back(sheet, px_per_mm, mode="RGB",
                                      profile=False, low_detail=True,
                                      loader=_cached_loader))
    return images[:limit]


def page_count(sheets, double_sided) -> int:
    """How many pages ``PdfFile.save`` would write for these sheets."""
    return len(sheets) * (2 if double_sided else 1)


def _plural(count: int, word: str) -> str:
    return word if count == 1 else word + "s"


def describe(sheets, double_sided) -> str:
    """A short English summary of what these sheets would produce."""
    cards = sum(len(sheet.cards) for sheet in sheets)
    sheets_count = len(sheets)
    pages = page_count(sheets, double_sided)
    return (
        f"{cards} {_plural(cards, 'card')} · "
        f"{sheets_count} {_plural(sheets_count, 'sheet')} · "
        f"{pages} {_plural(pages, 'page')}"
    )


def clear_cache() -> None:
    """Forget every cached card picture (call when the folder changes)."""
    _cache.clear()
