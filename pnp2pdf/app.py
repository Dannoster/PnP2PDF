from pnp2pdf import __version__ as PROGRAM_VERSION

# Re-exported for backwards compatibility, so that pnp2pdf.app.<name>
# keeps resolving for code that imports it from here.
from pnp2pdf.errors import InputError, OperationCancelled
from pnp2pdf.model import Card, CardBack, Sheet
from pnp2pdf.pdf import PdfFile
from pnp2pdf.cards import (
    IMAGE_EXTENSIONS, create_cards, find_back_dirs, find_card_dirs, natural_key,
)
from pnp2pdf.layout import create_sheets
from pnp2pdf.options import JobOptions
from pnp2pdf.job import create_pdf
from pnp2pdf.render import (
    CUT_LINE_COLOUR, PXLS_IN_MM, create_mirrored_frame, draw_cutting_corners,
    draw_lines, draw_szones, place_cropped_back,
)
from pnp2pdf.resources import icc_profile, resource_path, settings_path
from pnp2pdf.settings import DEFAULT_SETTINGS, create_save, load_save

from pnp2pdf.ui import window as _window


def build_pdf(values: dict, progress=None, cancelled=None) -> str:
    """Translate the window's widget values into a job and run it."""
    return create_pdf(JobOptions.from_form(values), progress=progress,
                      cancelled=cancelled)


def ui_window(settings):
    return _window.ui_window(settings)


def main():
    _window.main()


if __name__ == "__main__":
    main()
