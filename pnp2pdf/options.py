from dataclasses import dataclass

from pnp2pdf.errors import InputError


def _number(value, label: str) -> float:
    try:
        return float(str(value).replace(',', '.'))
    except (TypeError, ValueError):
        raise InputError(f'{label} must be a number (got "{value}").')


def _integer(value, label: str) -> int:
    number = _number(value, label)
    if not number.is_integer():
        raise InputError(f'{label} must be a whole number (got "{value}").')
    return int(number)


@dataclass(frozen=True)
class JobOptions:
    cards_folder: str
    output_folder: str
    paper_width: float = 210
    paper_height: float = 297
    card_width: float = 62.5
    card_height: float = 88.7
    rows: int = 3
    columns: int = 3
    gap: float = 2
    ud_szone: float = 0
    lr_szone: float = 0
    smaller_size: bool = False
    double_sided: bool = False

    @classmethod
    def from_form(cls, values: dict) -> "JobOptions":
        return cls(
            cards_folder=values['-CFOLDER-'],
            output_folder=values['-PDFFOLDER-'],
            paper_width=_number(values['-PWIDTH-'], "Paper width"),
            paper_height=_number(values['-PHEIGHT-'], "Paper height"),
            card_width=_number(values['-CWIDTH-'], "Card width"),
            card_height=_number(values['-CHEIGHT-'], "Card height"),
            rows=_integer(values['-ROWS-'], "Rows"),
            columns=_integer(values['-COLUMNS-'], "Columns"),
            gap=_number(values['-GAP-'], "Gap"),
            ud_szone=_number(values['-HSZONES-'], "Hor. safe zones"),
            lr_szone=_number(values['-VSZONES-'], "Vert. safe zones"),
            smaller_size=values['-SIZECHECKBOX-'],
            double_sided=values['-DOUBLECHECKBOX-'],
        )

    def validate(self) -> None:
        if not self.cards_folder or not self.cards_folder.strip():
            raise InputError("Cards folder must not be empty.")
        if not self.output_folder or not self.output_folder.strip():
            raise InputError("Output folder must not be empty.")
        if self.paper_width <= 0 or self.paper_height <= 0:
            raise InputError(
                f"Paper size must be positive (got {self.paper_width} x "
                f"{self.paper_height} mm)."
            )
        if self.card_width <= 0 or self.card_height <= 0:
            raise InputError(
                f"Card size must be positive (got {self.card_width} x "
                f"{self.card_height} mm)."
            )
        if self.rows < 1 or self.columns < 1:
            raise InputError(
                f"Rows and columns must be at least 1 (got {self.rows} rows, "
                f"{self.columns} columns)."
            )
        if self.gap < 0 or self.ud_szone < 0 or self.lr_szone < 0:
            raise InputError(
                f"Gap and margins must not be negative (got gap={self.gap}, "
                f"top/bottom={self.ud_szone}, left/right={self.lr_szone} mm)."
            )
