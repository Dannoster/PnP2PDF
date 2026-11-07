from dataclasses import dataclass


class CardBack:
    def __init__(self, dir: str, width: float, height: float):
        """
        'width' and 'height' in mm
        """
        self.dir = dir
        self.width = width
        self.height = height

    def __repr__(self):
        return f"CardBack({self.dir!r}, {self.width}x{self.height})"


class Card:
    """A printable card face, optionally with a back side."""
    def __init__(self, dir: str, width: float, height: int | float, back=None):
        self.dir = dir
        self.width = width
        self.height = height
        self.back = back

    def __repr__(self):
        return f"Card({self.dir!r}, {self.width}x{self.height}, back={self.back!r})"


@dataclass
class Sheet:
    """
    'width', 'height', 'gap', 'ud_szone', 'lr_szone' in mm
    """ 
    width: float
    height: float
    rows: int
    columns: int
    gap: float
    ud_szone: float
    lr_szone: float
    cards: list[Card|CardBack]
