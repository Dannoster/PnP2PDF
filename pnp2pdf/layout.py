from pnp2pdf.model import Card, CardBack, Sheet


def create_sheets(cards: list[Card|CardBack], sheet_width: float, sheet_height: float, 
                  rows: int, columns: int, gap: int, ud_szone: int, lr_szone: int):
    if rows < 1 or columns < 1:
        raise ValueError(
            f"Rows and columns must be at least 1 (got {rows} rows, "
            f"{columns} columns)."
        )
    if sheet_width <= 0 or sheet_height <= 0:
        raise ValueError(
            f"Paper size must be positive (got {sheet_width} x {sheet_height} mm)."
        )
    if gap < 0 or ud_szone < 0 or lr_szone < 0:
        raise ValueError(
            f"Gap and margins must not be negative (got gap={gap}, "
            f"top/bottom={ud_szone}, left/right={lr_szone} mm)."
        )
    remaining = list(cards)
    sheets = []
    while remaining:
        selected_cards = remaining[:rows*columns]
        del remaining[:rows*columns]
        sheet = Sheet(sheet_width, sheet_height, rows, columns, gap, ud_szone, lr_szone, selected_cards)
        sheets.append(sheet)
    
    return sheets
