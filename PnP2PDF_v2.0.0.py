import os
import webbrowser
import pickle
import datetime

import subprocess, sys
def install(package): subprocess.check_call([sys.executable, "-m", "pip", "install", package])
install("PySimpleGUI")
install("pillow")

import PySimpleGUI as sg
from PIL import Image, ImageDraw, ImageCms

PROGRAM_VERSION = "2.0.0"
PXLS_IN_MM = 11.81102362204 # need changes
DEFAULT_SETTINGS = {"paper_width"   : 210,
            "paper_height"  : 297,
            "card_width"    : 62.5,
            "card_height"   : 88.7,
            "resize"        : False,
            "double_sided"  : False,
            "rows"          : 3,
            "columns"       : 3,
            "gap"           : 2,
            "ud_bleed"      : 0, 
            "lr_bleed"      : 0,
            "cards_folder"  : "",
            "save_to"       : ""}
curr_path = os.path.dirname(os.path.realpath(__file__))

class CardBack:
    def __init__(self, dir: str, width: float, height: float):
        """
        'width' and 'height' in mm
        """
        self.dir = dir
        self.width = width
        self.height = height

class Card(CardBack):
    """
    Card class
    """    
    def __init__(self, dir: str, width: float, height: float, back: CardBack = None):
        super().__init__(dir, width, height)
        self.back = back


class Sheet:
    def __init__(self, width: float, height: float, 
                 rows: int, columns: int, gap: int, 
                 ud_bleed: int, lr_bleed: int, 
                 cards: list[Card|CardBack]):
        """
        'width', 'height', 'gap', 'ud_bleed', 'lr_bleed' in mm
        """ 
        self.width = width
        self.height = height

        self.rows = rows
        self.columns = columns
        self.gap = gap

        self.ud_bleed = ud_bleed
        self.lr_bleed = lr_bleed

        self.cards = cards

def draw_lines(draw: ImageDraw, paper_width, paper_height, 
               rows, columns, gap, card_width, card_height, number_of_cards):
    """all in pixels"""
    color = (0,0,0,47)#"darkgray"

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
        draw.line(((x, 0), (x, paper_height)), fill=color, width=2)
        x = x0 + card_width + i * (card_width + gap) + 1
        draw.line(((x - 2, 0), (x - 2, paper_height)), fill=color, width=2)
    for i in range(y_num):
        y = y0 + i * (card_height + gap) - 1
        draw.line(((0, y), (paper_width, y)), fill=color, width=2)
        y = y0 + card_height + i * (card_height + gap) + 1
        draw.line(((0, y - 2), (paper_width, y - 2)), fill=color, width=2)

def draw_bleeds(draw: ImageDraw, paper_width, paper_height, hbleed, vbleed):
    """paper_width, paper_height, hbleed, vbleed in pxls"""
    color = (0,143,156,56)#(215,0,0)
    line_width = 4
    if hbleed:
        draw.line(((0, hbleed-1), (paper_width, hbleed-1)), fill=color, width=line_width)
        draw.line(((0, paper_height-hbleed-1), (paper_width, paper_height-hbleed-1)), fill=color, width=line_width)
    if vbleed:
        draw.line(((vbleed-1, 0), (vbleed-1, paper_height)), fill=color, width=line_width)
        draw.line(((paper_width-vbleed-1, 0), (paper_width-vbleed-1, paper_height)), fill=color, width=line_width)

def draw_cutting_corners(x: int, y: int, draw: ImageDraw, card_width, card_height):
    """all in pixels"""
    color = (0,0,0,47)#"darkgray"
    X_LEN = 14
    # upper left corner
    draw.line(((x - 1, y - 1), 
            (x + X_LEN, y - 1)), 
            fill=color, 
            width=2)
    draw.line(((x - 1, y - 1), 
            (x - 1, y + X_LEN)), 
            fill=color, 
            width=2)
    # upper right corner
    draw.line(((x + card_width - 1, y), 
            (x + card_width - X_LEN - 1, y)), 
            fill=color, 
            width=2)
    draw.line(((x + card_width - 1, y), 
            (x + card_width - 1, y + X_LEN)), 
            fill=color, 
            width=2)
    # lower left corner
    draw.line(((x, y + card_height - 1), 
            (x + X_LEN, y + card_height - 1)), 
            fill=color, 
            width=2)
    draw.line(((x, y + card_height - 1), 
            (x, y + card_height - X_LEN - 1)), 
            fill=color, 
            width=2)
    # lower right corner
    draw.line(((x + card_width, y + card_height), 
            (x + card_width - X_LEN - 1, y + card_height)), 
            fill=color, 
            width=2)
    draw.line(((x + card_width, y + card_height), 
            (x + card_width, y + card_height - X_LEN - 1)), 
            fill=color, 
            width=2)

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
    
class PDF_file:
    def __init__(self, sheets: list[Sheet]):
        self.sheets = sheets

    def save(self, dir: str, smaller_size: bool, double_sided: bool, format="jpeg"): # need changes

        time = datetime.datetime.today().strftime("%d%m%y_%H%M")
        if not os.path.exists(f"{dir}/{time}"):
            try:
                os.mkdir(f"{dir}/{time}")
            except:
                os.mkdir(f"{dir}")
                os.mkdir(f"{dir}/{time}")

        sht = self.sheets[0]
        crd = sht.cards[0]
        swidth, sheight = int(sht.width*PXLS_IN_MM), int(sht.height*PXLS_IN_MM)
        hbleed, vbleed  = int(sht.ud_bleed*PXLS_IN_MM), int(sht.lr_bleed*PXLS_IN_MM), 
        cwidth, cheight = int(crd.width*PXLS_IN_MM), int(crd.height*PXLS_IN_MM)
        rows, columns, gap = sht.rows, sht.columns, int(sht.gap*PXLS_IN_MM)
        x0 = (swidth - (columns-1)*gap - columns*cwidth) // 2
        y0 = (sheight - (rows-1)*gap - rows*cheight) // 2

        for i, sheet in enumerate(self.sheets):
            page_a = Image.new('CMYK',#'RGB', 
                               (int(sheet.width*PXLS_IN_MM),int(sheet.height*PXLS_IN_MM)), 
                               (0,0,0,0))#(255,255,255))
            draw_a = ImageDraw.Draw(page_a)
            draw_lines(draw_a, swidth, sheight, rows, columns, gap, 
                       cwidth, cheight, number_of_cards=len(sheet.cards))
            for j, card in enumerate(sheet.cards):
                card_image = Image.open(card.dir).resize((int(card.width*PXLS_IN_MM), 
                                                          int(card.height*PXLS_IN_MM)))
                try:
                    card_image = ImageCms.profileToProfile(card_image, 
                                                           'color_profiles/sRGB-IEC61966-2.1.icc', 
                                                           'color_profiles/USWebCoatedSWOP.icc', 
                                                           outputMode='CMYK')
                except:
                    pass
                # print(card.dir.split("/")[-1], card_image.mode, card_image.load()[0,0])
                x = x0 + j%columns * (cwidth + gap)
                y = y0 + j//columns * (cheight + gap)
                page_a.paste(card_image, (x, y))
                create_mirrored_frame(x, y, page_a, card_image, gap, cwidth, cheight)
                draw_cutting_corners(x, y, draw_a, cwidth, cheight)
            draw_bleeds(draw_a, swidth, sheight, hbleed, vbleed)
            
            savedir = f"{dir}/{time}/{time}_page_{i+1:03}"
            if double_sided:
                savedir += "a"
            savedir += f".{format}"
            page_a.save(savedir, dpi=(300, 300), compression=None, quality=100)
            page_a.close()
        
        if double_sided:
            for i, sheet in enumerate(self.sheets):            
                page_b = Image.new('CMYK',#'RGB', 
                                   (int(sheet.width*PXLS_IN_MM),int(sheet.height*PXLS_IN_MM)), 
                                   (0,0,0,0))#(255,255,255))
                
                for j, card in enumerate(sheet.cards):
                    if card.back:
                        back_image = Image.open(card.back.dir).resize((int(card.back.width*PXLS_IN_MM),
                                                                       int(card.back.height*PXLS_IN_MM)))
                        try:
                            back_image = ImageCms.profileToProfile(back_image, 
                                                                   'color_profiles/sRGB-IEC61966-2.1.icc', 
                                                                   'color_profiles/USWebCoatedSWOP.icc', 
                                                                   outputMode='CMYK')
                        except: 
                            pass
                        x = x0 + (columns-1-j)%columns * (cwidth + gap)
                        y = y0 + j//columns * (cheight + gap)
                        if card.back.width == card.width and card.back.height == card.height:
                            page_b.paste(back_image, (x, y))
                            create_mirrored_frame(x, y, page_b, back_image, gap, 
                                                  int(card.back.width*PXLS_IN_MM), 
                                                  int(card.back.height*PXLS_IN_MM))  
                        else:
                            place_cropped_back(x, y, page_b, back_image, gap, cwidth, cheight,
                                               int(card.back.width*PXLS_IN_MM), 
                                               int(card.back.height*PXLS_IN_MM))

                page_b.save(f"{dir}/{time}/{time}_page_{i+1:03}b.{format}", 
                            dpi=(300, 300), compression=None, quality=100)
                page_b.close()
                
        pages_dirs = sorted(os.listdir(f"{dir}/{time}"))
        pages = [Image.open(f"{dir}/{time}/{directory}") for directory in pages_dirs]
        pages[0].save(f"{dir}/{time}/{time}.pdf", "PDF", optimize=False, 
                        save_all=True, append_images=pages[1:], 
                        dpi=(300, 300), compression=None, quality=100)

            
            




def load_save(dir: str) -> dict:
    if os.path.exists(dir):
        with open(dir, "rb") as file:
            return pickle.load(file)
    else:
        return DEFAULT_SETTINGS

def create_save(dir: str, save: dict):
    with open(dir, 'wb') as file:     
        pickle.dump(save, file)

def find_back_dirs(folder: str) -> list[str]:
    back_dirs = []
    for file in os.listdir(folder):
        if file.lower().endswith(".jpg") or file.lower().endswith(".jpeg") \
                or file.lower().endswith(".png") or file.lower().endswith(".tif"):
            back_dirs.append(os.path.join(folder, file))
    back_dirs.sort()
    return back_dirs

def find_card_dirs(folder):
    card_dirs = []
    for path, subdirs, files in os.walk(folder):
        for name in files:
            if ".jpg" in name.lower() or ".jpeg" in name.lower() \
                    or ".png" in name.lower() or ".tif" in name.lower():
                if not "page_" in name:
                    card_dirs.append(f"{path}/{name}")
    card_dirs.sort()
    return card_dirs

def create_cards(root_folder, card_width, card_height, smaller_size: bool, double_sided: bool):

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
                back = CardBack(dir, 
                                koef * float(dir.split("_")[-2].replace(',','.')),
                                koef * float("".join(dir.split("_")[-1].split(".")[:-1]).replace(',','.')))
            except:
                back = CardBack(dir, card_width, card_height)
            backs.append(back)

        for back in backs:
            folder_name = back.dir.split("/")[-1].split("_")[0].split(".")[0]
            card_dirs = find_card_dirs(f"{root_folder}/{folder_name}")
            cards = [Card(dir, card_width, card_height, back) for dir in card_dirs]
            all_cards.extend(cards)

        double_card_dirs = find_card_dirs(f"{root_folder}/Double Sided")
        for i in range(0, len(double_card_dirs), 2):
            back = CardBack(double_card_dirs[i+1], card_width, card_height)
            all_cards.append(Card(double_card_dirs[i], card_width, card_height, back))
        single_side_folder += "/Single Sided"

    single_cards = [Card(dir, card_width, card_height) for dir in find_card_dirs(single_side_folder)]
    all_cards.extend(single_cards)
    all_cards.sort(key=lambda x: x.dir.split("/")[-1])

    return all_cards

def create_sheets(cards: list[Card|CardBack], sheet_width: float, sheet_height: float, 
                  rows: int, columns: int, gap: int, ud_bleed: int, lr_bleed: int):
    # if len(cards) % (rows*columns) == 0:    
    #     sheets_total = len(cards) // (rows*columns)
    # else:
    #     sheets_total = len(cards) // (rows*columns) + 1
    
    sheets = []
    while cards:
        selected_cards = cards[:rows*columns]
        del cards[:rows*columns]
        sheet = Sheet(sheet_width, sheet_height, rows, columns, gap, ud_bleed, lr_bleed, selected_cards)
        sheets.append(sheet)
    
    return sheets


    

def ui_window(settings: dict):
    form_rows = [
                 [sg.Text('Paper size', size=(14, 1)), 
                  sg.InputText(settings["paper_width"], size=(6, 1), key="-PWIDTH-"), sg.Text('x'), 
                  sg.InputText(settings["paper_height"], size=(6, 1), key="-PHEIGHT-"), sg.Text('mm'), 
                  sg.Push(), 
                  sg.Text('Contact me', enable_events = True, 
                          font = ('Courier New', 12, 'underline'), 
                          key='URL https://www.t.me/nNekoChan')],

                 [sg.Text('Card size', size=(14, 1)), 
                  sg.InputText(f"{settings['card_width']}", size=(6, 1), key="-CWIDTH-"), sg.Text('x'), 
                  sg.InputText(f"{settings['card_height']}", size=(6, 1), key="-CHEIGHT-"), sg.Text('mm'), 
                  sg.Push(),
                  sg.Checkbox("Smaller size (98%)", default=settings["resize"], key='-SIZECHECKBOX-')],

                 [sg.Text("Pattern", size=(14, 1)),
                  sg.InputText(f"{settings['rows']}", size=(6, 1), key="-ROWS-"), sg.Text('x'), 
                  sg.InputText(f"{settings['columns']}", size=(6, 1), key="-COLUMNS-"), 
                  sg.Text('(Raws x Columns)'),
                  sg.Push(),
                  sg.Checkbox("Double Sided Cards", default=settings["double_sided"], key='-DOUBLECHECKBOX-')],

                 [sg.Text('Space between', size=(14, 1)), 
                  sg.InputText(f"{settings['gap']}", size=(6, 1), key="-GAP-"), sg.Text('mm'),
                  sg.Push(),
                  sg.Text('Prepare files for DS print', enable_events = True, 
                          font = ('Courier New', 10, 'underline'), 
                          key='URL https://github.com/Dannoster/PnP2PDF/tree/main#how-to-prepare-double-sided-cards')],

                 [sg.Text("Horizontal bleeds", size=(14, 1)),
                  sg.InputText(f"{settings['ud_bleed']}", size=(6, 1), key="-HBLEEDS-"), sg.Text('mm\t'),
                  sg.Text("Vertical bleeds", size=(14, 1)),
                  sg.InputText(f"{settings['lr_bleed']}", size=(6, 1), key="-VBLEEDS-"), sg.Text('mm')], 

                 [sg.Text('Cards folder', size=(14, 1)), 
                  sg.InputText(f"{settings['cards_folder']}", key='-CFOLDER-'), 
                  sg.FolderBrowse(initial_folder=settings['cards_folder'])],

                 [sg.Text('Save PDF to', size=(14, 1)), 
                  sg.InputText(f"{settings['save_to']}", key='-PDFFOLDER-'), 
                  sg.FolderBrowse(initial_folder=settings['save_to'])],

                 [sg.Submit("Create PDF", size=(14, 1)), sg.Push(), sg.Text(f"v{PROGRAM_VERSION}")]
                ]

    window = sg.Window('PnP2PDF', form_rows)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event.startswith("URL "):
            url = event.split(' ')[1]
            webbrowser.open(url)
        elif event.startswith("Create PDF"):
            cards = create_cards(values['-CFOLDER-'], 
                                 float(values['-CWIDTH-'].replace(',','.')), 
                                 float(values['-CHEIGHT-'].replace(',','.')),
                                 values["-SIZECHECKBOX-"],
                                 values['-DOUBLECHECKBOX-'])
            sheets = create_sheets(cards, 
                                   float(values['-PWIDTH-'].replace(',','.')), 
                                   float(values['-PHEIGHT-'].replace(',','.')), 
                                   int(values['-ROWS-']), 
                                   int(values['-COLUMNS-']), 
                                   float(values['-GAP-'].replace(',','.')), 
                                   float(values['-HBLEEDS-'].replace(',','.')), 
                                   float(values['-VBLEEDS-'].replace(',','.')))
            pdf = PDF_file(sheets)
            pdf.save(values['-PDFFOLDER-'],
                     values["-SIZECHECKBOX-"], # added recently
                     values['-DOUBLECHECKBOX-'])
            window.close()
            return values
            
        
        
def main():
    save_folder = "presets.pickle"
    settings = load_save(save_folder)
    ui_output = ui_window(settings)
    settings = {"paper_width"   : ui_output['-PWIDTH-'],
                "paper_height"  : ui_output['-PHEIGHT-'],
                "card_width"    : ui_output['-CWIDTH-'],
                "card_height"   : ui_output['-CHEIGHT-'],
                "resize"        : ui_output['-SIZECHECKBOX-'],
                "double_sided"  : ui_output['-DOUBLECHECKBOX-'],
                "rows"          : ui_output['-ROWS-'],
                "columns"       : ui_output['-COLUMNS-'],
                "gap"           : ui_output['-GAP-'],
                "ud_bleed"      : ui_output['-HBLEEDS-'], 
                "lr_bleed"      : ui_output['-VBLEEDS-'],
                "cards_folder"  : ui_output['-CFOLDER-'],
                "save_to"       : ui_output['-PDFFOLDER-']}
    create_save(save_folder, settings)

main()