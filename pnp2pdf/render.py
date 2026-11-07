import os

from PIL import Image, ImageCms, ImageDraw

from pnp2pdf.resources import icc_profile

PXLS_IN_MM = 11.81102362204 # pixels per millimetre at 300 dpi

# Cut lines are drawn straight onto a CMYK page, so this is a CMYK colour:
# (0, 0, 0, 47) is black at 47%, a light grey. An alpha value would be
# silently ignored on a CMYK image.
CUT_LINE_COLOUR = (0, 0, 0, 47)


def draw_lines(draw: ImageDraw, paper_width, paper_height, 
               rows, columns, gap, card_width, card_height, number_of_cards):
    """all in pixels"""

    x0 = (paper_width - (columns-1)*gap - columns*card_width) // 2
    y0 = (paper_height - (rows-1)*gap - rows*card_height) // 2

    if number_of_cards >= columns:
        x_num = columns
    else:
        x_num = number_of_cards

    if number_of_cards % columns != 0:
        y_num = number_of_cards // columns + 1
    else:
        y_num = number_of_cards // columns

    for i in range(x_num):
        x = x0 + i * (card_width + gap) - 1
        draw.line(((x, 0), (x, paper_height)), fill=CUT_LINE_COLOUR, width=2)
        x = x0 + card_width + i * (card_width + gap) + 1
        draw.line(((x - 2, 0), (x - 2, paper_height)), fill=CUT_LINE_COLOUR, width=2)
    for i in range(y_num):
        y = y0 + i * (card_height + gap) - 1
        draw.line(((0, y), (paper_width, y)), fill=CUT_LINE_COLOUR, width=2)
        y = y0 + card_height + i * (card_height + gap) + 1
        draw.line(((0, y - 2), (paper_width, y - 2)), fill=CUT_LINE_COLOUR, width=2)

def draw_szones(draw: ImageDraw, paper_width, paper_height, hszone, vszone):
    """paper_width, paper_height, hszone, vszone in pxls"""
    color = (0,143,156,56)#(215,0,0)
    line_width = 4
    if hszone:
        draw.line(((0, hszone-1), (paper_width, hszone-1)), fill=color, width=line_width)
        draw.line(((0, paper_height-hszone-1), (paper_width, paper_height-hszone-1)), fill=color, width=line_width)
    if vszone:
        draw.line(((vszone-1, 0), (vszone-1, paper_height)), fill=color, width=line_width)
        draw.line(((paper_width-vszone-1, 0), (paper_width-vszone-1, paper_height)), fill=color, width=line_width)

def draw_cutting_corners(x: int, y: int, draw: ImageDraw, card_width, card_height):
    """Draw L-shaped corner marks in the band just outside a card. All in pixels."""
    thickness = 2
    arm = 14
    # the marks sit in a 2 px band outside the card's pixel box
    left, top = x - thickness, y - thickness
    right, bottom = x + card_width + thickness - 1, y + card_height + thickness - 1

    # horizontal arms
    draw.rectangle((left, top, left + arm, top + thickness - 1), fill=CUT_LINE_COLOUR)
    draw.rectangle((right - arm, top, right, top + thickness - 1), fill=CUT_LINE_COLOUR)
    draw.rectangle((left, bottom - thickness + 1, left + arm, bottom), fill=CUT_LINE_COLOUR)
    draw.rectangle((right - arm, bottom - thickness + 1, right, bottom), fill=CUT_LINE_COLOUR)

    # vertical arms
    draw.rectangle((left, top, left + thickness - 1, top + arm), fill=CUT_LINE_COLOUR)
    draw.rectangle((right - thickness + 1, top, right, top + arm), fill=CUT_LINE_COLOUR)
    draw.rectangle((left, bottom - arm, left + thickness - 1, bottom), fill=CUT_LINE_COLOUR)
    draw.rectangle((right - thickness + 1, bottom - arm, right, bottom), fill=CUT_LINE_COLOUR)

def create_mirrored_frame(x: int, y: int, sheet: Image, image: Image, gap, card_width, card_height):
    """all in pixels"""
    # Placing mirrored card frame
    left_right = image.transpose(Image.FLIP_LEFT_RIGHT)
    top_bottom = image.transpose(Image.FLIP_TOP_BOTTOM)
    sheet.paste(left_right.crop((card_width - gap//2,0, 
                                 card_width,card_height)),
                (x-gap//2, y))
    sheet.paste(left_right.crop((0,0, 
                                 gap//2,card_height)),
                (x+card_width, y))
    sheet.paste(top_bottom.crop((0,card_height - gap//2, 
                                 card_width,card_height)),
                (x,y-gap//2))
    sheet.paste(top_bottom.crop((0,0, 
                                 card_width,gap//2)),
                (x,y+card_height))

def place_cropped_back(x: int, y: int, sheet: Image, image: Image, gap, 
                       card_width, card_height, back_width, back_height):
    """all in pixels"""
    target_width = gap + card_width
    target_height = gap + card_height
    cropped_w = (back_width - target_width)//2
    cropped_h = (back_height - target_height)//2
    image = image.crop((cropped_w,cropped_h,
                        back_width - cropped_w,back_height - cropped_h))
    sheet.paste(image, (x - gap//2 - 1, y - gap//2 - 1))


def page_geometry(sheet, px_per_mm) -> dict:
    """The pixel numbers ``PdfFile.save`` used to compute for a sheet."""
    card = sheet.cards[0]
    swidth = int(sheet.width*px_per_mm)
    sheight = int(sheet.height*px_per_mm)
    hszone = int(sheet.ud_szone*px_per_mm)
    vszone = int(sheet.lr_szone*px_per_mm)
    cwidth = int(card.width*px_per_mm)
    cheight = int(card.height*px_per_mm)
    rows = sheet.rows
    columns = sheet.columns
    gap = int(sheet.gap*px_per_mm)
    x0 = (swidth - (columns-1)*gap - columns*cwidth) // 2
    y0 = (sheight - (rows-1)*gap - rows*cheight) // 2
    return {
        "swidth": swidth, "sheight": sheight, "hszone": hszone,
        "vszone": vszone, "cwidth": cwidth, "cheight": cheight,
        "rows": rows, "columns": columns, "gap": gap, "x0": x0, "y0": y0,
    }


def _default_loader(profile: bool):
    """The picture loader that reproduces ``PdfFile.save`` exactly."""

    def load(card, size):
        card_image = Image.open(card.dir).resize(size)
        if profile:
            try:
                card_image = ImageCms.profileToProfile(
                    card_image, icc_profile("sRGB-IEC61966-2.1.icc"),
                    icc_profile("USWebCoatedSWOP.icc"), outputMode='CMYK')
            except Exception as error:
                print(f"WARNING: colour profile not applied to "
                      f"{os.path.basename(card.dir)}: {error}")
        return card_image

    return load


def _flat_block(mode, width, height):
    fill = (128, 128, 128, 128) if mode == "CMYK" else (128, 128, 128)
    return Image.new(mode, (width, height), fill)


def _blank_page(mode, width, height):
    fill = (0, 0, 0, 0) if mode == "CMYK" else (255, 255, 255)
    return Image.new(mode, (width, height), fill)


def render_front(sheet, px_per_mm, *, mode="CMYK", profile=True,
                 low_detail=False, loader=None) -> Image.Image:
    """Draw the front of ``sheet`` at ``px_per_mm`` pixels per millimetre."""
    if loader is None:
        loader = _default_loader(profile)
    geom = page_geometry(sheet, px_per_mm)
    swidth, sheight = geom["swidth"], geom["sheight"]
    hszone, vszone = geom["hszone"], geom["vszone"]
    cwidth, cheight = geom["cwidth"], geom["cheight"]
    rows, columns, gap = geom["rows"], geom["columns"], geom["gap"]
    x0, y0 = geom["x0"], geom["y0"]

    page = _blank_page(mode, swidth, sheight)
    draw = ImageDraw.Draw(page)
    draw_lines(draw, swidth, sheight, rows, columns, gap,
               cwidth, cheight, number_of_cards=len(sheet.cards))
    for j, card in enumerate(sheet.cards):
        size = (int(card.width*px_per_mm), int(card.height*px_per_mm))
        if low_detail and (size[0] < 2 or size[1] < 2):
            card_image = _flat_block(mode, size[0], size[1])
        else:
            card_image = loader(card, size)
        x = x0 + j%columns * (cwidth + gap)
        y = y0 + j//columns * (cheight + gap)
        page.paste(card_image, (x, y))
        create_mirrored_frame(x, y, page, card_image, gap, cwidth, cheight)
        draw_cutting_corners(x, y, draw, cwidth, cheight)
    draw_szones(draw, swidth, sheight, hszone, vszone)
    return page


def render_back(sheet, px_per_mm, *, mode="CMYK", profile=True,
                low_detail=False, loader=None) -> Image.Image:
    """Draw the mirrored back of ``sheet`` at ``px_per_mm``."""
    if loader is None:
        loader = _default_loader(profile)
    geom = page_geometry(sheet, px_per_mm)
    swidth, sheight = geom["swidth"], geom["sheight"]
    cwidth, cheight = geom["cwidth"], geom["cheight"]
    columns, gap = geom["columns"], geom["gap"]
    x0, y0 = geom["x0"], geom["y0"]

    page = _blank_page(mode, swidth, sheight)
    for j, card in enumerate(sheet.cards):
        if not card.back:
            continue
        back = card.back
        back_width = int(back.width*px_per_mm)
        back_height = int(back.height*px_per_mm)
        if low_detail and (back_width < 2 or back_height < 2):
            back_image = _flat_block(mode, back_width, back_height)
        else:
            back_image = loader(back, (back_width, back_height))
        x = x0 + (columns-1-j)%columns * (cwidth + gap)
        y = y0 + j//columns * (cheight + gap)
        if back.width == card.width and back.height == card.height:
            page.paste(back_image, (x, y))
            create_mirrored_frame(x, y, page, back_image, gap,
                                  back_width, back_height)
        else:
            place_cropped_back(x, y, page, back_image, gap, cwidth, cheight,
                               back_width, back_height)
    return page
