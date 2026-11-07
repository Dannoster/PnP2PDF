import os
import re

from pnp2pdf.model import Card, CardBack


IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".tif", ".tiff")


def natural_key(name: str) -> list:
    """Sort key that orders 1, 2, ..., 10 instead of 1, 10, 2."""
    parts = re.split(r"(\d+)", os.path.basename(name).lower())
    return [int(part) if part.isdigit() else part for part in parts]

def find_back_dirs(folder: str) -> list[str]:
    back_dirs = []
    for file in os.listdir(folder):
        if file.lower().endswith(IMAGE_EXTENSIONS):
            back_dirs.append(os.path.join(folder, file))
    back_dirs.sort(key=natural_key)
    return back_dirs

def find_card_dirs(folder):
    card_dirs = []
    for path, subdirs, files in os.walk(folder):
        for name in files:
            if name.lower().endswith(IMAGE_EXTENSIONS) and "page_" not in name:
                card_dirs.append(os.path.join(path, name))
    card_dirs.sort(key=natural_key)
    return card_dirs

def create_cards(root_folder, card_width, card_height, smaller_size: bool, double_sided: bool):

    if card_width <= 0 or card_height <= 0:
        raise ValueError(
            f"Card size must be positive (got {card_width} x {card_height} mm)."
        )

    if smaller_size:
        koef = .98
    else:
        koef = 1
    card_width *= koef
    card_height *= koef
        
    all_cards = []
    single_side_folder = root_folder

    if double_sided:
       
        backs = []
        for dir in find_back_dirs(root_folder):
            try:
                stem = os.path.splitext(os.path.basename(dir))[0]
                parts = stem.split("_")
                back = CardBack(dir,
                                koef * float(parts[-2].replace(',', '.')),
                                koef * float(parts[-1].replace(',', '.')))
            except (IndexError, ValueError):
                back = CardBack(dir, card_width, card_height)
            backs.append(back)

        for back in backs:
            folder_name = os.path.basename(back.dir).split("_")[0].split(".")[0]
            card_dirs = find_card_dirs(os.path.join(root_folder, folder_name))
            cards = [Card(dir, card_width, card_height, back) for dir in card_dirs]
            all_cards.extend(cards)

        double_card_dirs = find_card_dirs(os.path.join(root_folder, "Double Sided"))
        if len(double_card_dirs) % 2 != 0:
            raise ValueError(
                f'Found an odd number of files ({len(double_card_dirs)}) in '
                f'"{os.path.join(root_folder, "Double Sided")}".\n'
                'Files in that folder are paired in sorted order as front, back, '
                'front, back, ... - add or remove one file so the count is even.'
            )
        for i in range(0, len(double_card_dirs), 2):
            back = CardBack(double_card_dirs[i+1], card_width, card_height)
            all_cards.append(Card(double_card_dirs[i], card_width, card_height, back))
        single_side_folder = os.path.join(single_side_folder, "Single Sided")

    single_cards = [Card(dir, card_width, card_height) for dir in find_card_dirs(single_side_folder)]
    all_cards.extend(single_cards)
    all_cards.sort(key=lambda card: natural_key(card.dir))

    if not all_cards:
        raise ValueError(
            f"No card images found in \"{root_folder}\".\n"
            "Supported formats: .png, .jpg, .jpeg, .tif"
        )

    return all_cards
