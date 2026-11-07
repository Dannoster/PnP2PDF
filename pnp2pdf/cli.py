"""Command line entry point for PnP2PDF."""

import argparse
import sys

from pnp2pdf import __version__
from pnp2pdf.errors import InputError, OperationCancelled
from pnp2pdf.job import create_pdf
from pnp2pdf.options import JobOptions

_SEPARATORS = ("×", "X", "x")


class _UsageError(Exception):
    """Raised by the parser instead of exiting so `main` can return 2."""


class _ArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        raise _UsageError(message)


def _split_pair(text: str, width_name: str, height_name: str) -> list[str]:
    normalised = text
    for separator in _SEPARATORS:
        normalised = normalised.replace(separator, "x")
    parts = normalised.split("x")
    if len(parts) != 2 or not parts[0] or not parts[1]:
        raise argparse.ArgumentTypeError(
            f'{width_name} and {height_name} must be in the form WxH '
            f'(got "{text}")')
    return parts


def _parse_number(text: str, name: str) -> float:
    try:
        return float(text.replace(",", "."))
    except ValueError:
        raise argparse.ArgumentTypeError(
            f'{name} must be a number (got "{text}")')


def _parse_size(text: str, width_name: str,
                height_name: str) -> tuple[float, float]:
    width, height = _split_pair(text, width_name, height_name)
    return (_parse_number(width, width_name),
            _parse_number(height, height_name))


def _parse_grid(text: str) -> tuple[int, int]:
    rows, columns = _split_pair(text, "Rows", "Columns")
    return (_parse_int(rows, "Rows"), _parse_int(columns, "Columns"))


def _parse_int(text: str, name: str) -> int:
    try:
        return int(text)
    except ValueError:
        raise argparse.ArgumentTypeError(
            f'{name} must be a whole number (got "{text}")')


def _build_parser() -> argparse.ArgumentParser:
    parser = _ArgumentParser(
        prog="pnp2pdf",
        description="Assemble print-and-play card images into a PDF.",
    )
    parser.add_argument("--cards", required=True, metavar="DIR",
                        help="folder containing the card images")
    parser.add_argument("--out", required=True, metavar="DIR",
                        help="folder where the PDF is written")
    parser.add_argument("--paper", metavar="WxH",
                        type=lambda text: _parse_size(
                            text, "Paper width", "Paper height"),
                        default=(210.0, 297.0),
                        help="paper size in mm (default: 210x297)")
    parser.add_argument("--card", metavar="WxH",
                        type=lambda text: _parse_size(
                            text, "Card width", "Card height"),
                        default=(62.5, 88.7),
                        help="card size in mm (default: 62.5x88.7)")
    parser.add_argument("--grid", metavar="RxC", type=_parse_grid,
                        default=(3, 3),
                        help="rows by columns (default: 3x3)")
    parser.add_argument("--gap", metavar="MM",
                        type=lambda text: _parse_number(text, "Gap"),
                        default=2.0,
                        help="gap between cards in mm (default: 2)")
    parser.add_argument("--safe-ud", metavar="MM", dest="safe_ud",
                        type=lambda text: _parse_number(
                            text, "Top/bottom safe zone"),
                        default=0.0,
                        help="top/bottom safe zone in mm (default: 0)")
    parser.add_argument("--safe-lr", metavar="MM", dest="safe_lr",
                        type=lambda text: _parse_number(
                            text, "Left/right safe zone"),
                        default=0.0,
                        help="left/right safe zone in mm (default: 0)")
    parser.add_argument("--smaller-size", action="store_true",
                        help="scale oversized cards down to fit")
    parser.add_argument("--double-sided", action="store_true",
                        help="also print the backs of double-sided cards")
    parser.add_argument("--quiet", action="store_true",
                        help="suppress the progress report")
    parser.add_argument("--version", action="version",
                        version=f"pnp2pdf {__version__}")
    return parser


class _Progress:
    """Writes ``\\rPage n of m`` to stderr, cleared when the run ends."""

    def __init__(self, quiet: bool):
        self.quiet = quiet
        self.wrote = False

    def __call__(self, done: int, total: int, text: str) -> None:
        if self.quiet:
            return
        sys.stderr.write(f"\r{text}")
        sys.stderr.flush()
        self.wrote = True

    def finish(self) -> None:
        if self.wrote:
            sys.stderr.write("\n")
            sys.stderr.flush()


def main(argv=None) -> int:
    parser = _build_parser()
    try:
        args = parser.parse_args(argv)
    except _UsageError as error:
        print(f"error: {error}", file=sys.stderr)
        return 2
    except SystemExit as exit_code:
        return exit_code.code if isinstance(exit_code.code, int) else 0

    options = JobOptions(
        cards_folder=args.cards,
        output_folder=args.out,
        paper_width=args.paper[0],
        paper_height=args.paper[1],
        card_width=args.card[0],
        card_height=args.card[1],
        rows=args.grid[0],
        columns=args.grid[1],
        gap=args.gap,
        ud_szone=args.safe_ud,
        lr_szone=args.safe_lr,
        smaller_size=args.smaller_size,
        double_sided=args.double_sided,
    )

    progress = _Progress(args.quiet)
    try:
        create_pdf(options, progress=progress, cancelled=lambda: False)
    except InputError as error:
        progress.finish()
        print(f"error: {error}", file=sys.stderr)
        return 2
    except (KeyboardInterrupt, OperationCancelled):
        progress.finish()
        print("cancelled", file=sys.stderr)
        return 130
    except Exception as error:
        progress.finish()
        print(str(error), file=sys.stderr)
        return 1
    progress.finish()
    return 0


if __name__ == "__main__":
    sys.exit(main())
