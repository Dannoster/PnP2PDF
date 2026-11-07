from PIL import Image, ImageDraw

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
