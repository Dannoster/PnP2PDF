import datetime
import os
import webbrowser

import subprocess
import sys
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
install("pillow")
install("pathlib")
install("img2pdf")
install("pypdf")
install("PySimpleGUI")

from PIL import Image, ImageDraw
import img2pdf
from pathlib import Path
from pypdf import PdfMerger
import PySimpleGUI as sg

# Constants in px
PXLS_IN_MM  = 11.8095238095
A4_WIDTH = 210 # 2480px
A4_HEIGHT = 297 # 3508px
ARKHAM_CARD_WIDTH = "62.5" # 738px
ARKHAM_CARD_HEIGHT = "88.7" # 1048px

# standard values
paper_width = A4_WIDTH
paper_height = A4_HEIGHT
card_width = ARKHAM_CARD_WIDTH
card_height = ARKHAM_CARD_HEIGHT
resize = 0
rows = 3
columns = 3
gap = 2 # 24px
curr_path = os.path.dirname(os.path.realpath(__file__))
cards_folder = "" #curr_path.replace("PnP2PDF.app/Contents/Frameworks", "")
result_folder = "" #curr_path.replace("PnP2PDF.app/Contents/Frameworks", "")

def ui_window():
    form_rows = [
                 [sg.Text('Paper size', size=(12, 1)), 
                  sg.InputText(f"{paper_width}", size=(6, 1), key="-PWIDTH-"), 
                  sg.Text('x'), 
                  sg.InputText(f"{paper_height}", size=(6, 1), key="-PHEIGHT-"), 
                  sg.Text('mm'), 
                  sg.Push(), 
                  sg.Text('Contact me', enable_events = True, 
                          font = ('Courier New', 12, 'underline'), 
                          key='URL https://www.t.me/nNekoChan')],

                 [sg.Text('Card size', size=(12, 1)), 
                  sg.InputText(f"{card_width}", size=(6, 1), key="-CWIDTH-"), 
                  sg.Text('x'), 
                  sg.InputText(f"{card_height}", size=(6, 1), key="-CHEIGHT-"), 
                  sg.Text('mm'), 
                  sg.Push(),
                  sg.Checkbox("Smaller size (98%)", default=bool(resize), key='-CHECKBOX-')],

                 [sg.Text("Pattern", size=(12, 1)),
                  sg.InputText(f"{rows}", size=(6, 1), key="-ROWS-"), 
                  sg.Text('x'), 
                  sg.InputText(f"{columns}", size=(6, 1), key="-COLUMNS-"),
                  sg.Text('(Raws x Columns)')],

                 [sg.Text('Space between', size=(12, 1)), 
                  sg.InputText(f"{gap}", size=(6, 1), key="-GAP-"),
                  sg.Text('mm')],

                 [sg.Text('Cards folder', size=(12, 1)), 
                  sg.InputText(f"{cards_folder}", key='-FOLDER1-'), 
                  sg.FolderBrowse(initial_folder=cards_folder)],

                 [sg.Text('Save PDF to', size=(12, 1)), 
                  sg.InputText(f"{result_folder}", key='-FOLDER2-'), 
                  sg.FolderBrowse(initial_folder=result_folder)],

                 [sg.Submit("Create PDF", size=(12, 1)), 
                  sg.Push(), 
                  sg.Text("v1.0.0")]
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
            event, values = window.read()
            window.close()
            return values

    event, values = window.read()
    window.close()
    return values

def draw_lines(draw: ImageDraw):
    x0 = (paper_width - (columns-1)*gap - columns*card_width) // 2
    y0 = (paper_height - (rows-1)*gap - rows*card_height) // 2
    for i in range(columns):
        x = x0 + i * (card_width + gap) - 1
        draw.line(((x, 0), (x, paper_height)), fill="darkgray", width=2)
        x = x0 + card_width + i * (card_width + gap) + 1
        draw.line(((x - 2, 0), (x - 2, paper_height)), fill="darkgray", width=2)
    for i in range(rows):
        y = y0 + i * (card_height + gap) - 1
        draw.line(((0, y), (paper_width, y)), fill="darkgray", width=2)
        y = y0 + card_height + i * (card_height + gap) + 1
        draw.line(((0, y - 2), (paper_width, y - 2)), fill="darkgray", width=2)

def draw_cutting_corners(x: int, y: int, draw: ImageDraw):
    X_LEN = 14
    # upper left corner
    draw.line(((x - 1, y - 1), 
            (x + X_LEN, y - 1)), 
            fill="darkgray", 
            width=2)
    draw.line(((x - 1, y - 1), 
            (x - 1, y + X_LEN)), 
            fill="darkgray", 
            width=2)
    # upper right corner
    draw.line(((x + card_width - 1, y), 
            (x + card_width - X_LEN - 1, y)), 
            fill="darkgray", 
            width=2)
    draw.line(((x + card_width - 1, y), 
            (x + card_width - 1, y + X_LEN)), 
            fill="darkgray", 
            width=2)
    # lower left corner
    draw.line(((x, y + card_height - 1), 
            (x + X_LEN, y + card_height - 1)), 
            fill="darkgray", 
            width=2)
    draw.line(((x, y + card_height - 1), 
            (x, y + card_height - X_LEN - 1)), 
            fill="darkgray", 
            width=2)
    # lower right corner
    draw.line(((x + card_width, y + card_height), 
            (x + card_width - X_LEN - 1, y + card_height)), 
            fill="darkgray", 
            width=2)
    draw.line(((x + card_width, y + card_height), 
            (x + card_width, y + card_height - X_LEN - 1)), 
            fill="darkgray", 
            width=2)

def create_mirrored_frame(x: int, y: int, sheet: Image, image: Image):
    # Placing mirrored card frame
    left_right = image.transpose(Image.FLIP_LEFT_RIGHT)
    top_bottom = image.transpose(Image.FLIP_TOP_BOTTOM)
    sheet.paste(left_right.crop((card_width - gap//2, 
                                0, 
                                card_width, 
                                card_height)),
                (x-gap//2, y))
    sheet.paste(left_right.crop((0, 
                                0, 
                                gap//2, 
                                card_height)),
                (x+card_width, y))
    sheet.paste(top_bottom.crop((0, 
                                card_height - gap//2, 
                                card_width, 
                                card_height)),
                (x,y-gap//2))
    sheet.paste(top_bottom.crop((0, 
                                0, 
                                card_width, 
                                gap//2)),
                                (x,y+card_height))
    
def merge_sheets_into_pdf(sheets_dirs: list[str], time: str):
    merger = PdfMerger()
    for dir in sheets_dirs:
        merger.append(dir)
        # os.remove(dir) # bug on windows
    merger.write(f"{result_folder}/{time}/result.pdf")
    merger.close()

def create_pdf():
    time = datetime.datetime.today().strftime("%d%m%y_%H%M")

    # Search cards in folder and subfolders
    imgs_dirs = []
    for path, subdirs, files in os.walk(cards_folder):
        for name in files:
            if ".jpg" in name.lower() or ".png" in name.lower():
                if not "page_" in name:
                    imgs_dirs.append(f"{path}/{name}")
    if not imgs_dirs:
        print("Couldn't find any image")
        return
    imgs_dirs.sort()
    images = [Image.open(dir) for dir in imgs_dirs]

    cards_total = len(images)
    sheets_total = 0

    while imgs_dirs:
        # Taking cards
        sheets_total += 1
        if len(imgs_dirs) >= columns*rows:
            one_sheet_imgs = [Image.open(imgs_dirs[i]) for i in range(columns*rows)]
        else:
            one_sheet_imgs = [Image.open(dir) for dir in imgs_dirs]
        del imgs_dirs[:columns*rows]

        # Creating A4 sheet
        sheet = Image.new('RGB',(paper_width, paper_height), (255,255,255))
        draw = ImageDraw.Draw(sheet)

        draw_lines(draw)
    
        x0 = (paper_width - (columns-1)*gap - columns*card_width) // 2
        y0 = (paper_height - (rows-1)*gap - rows*card_height) // 2

        # Placing cards
        for i, image in enumerate(one_sheet_imgs):
            x = x0 + i%columns * (card_width + gap)
            y = y0 + i//columns * (card_height + gap)
            image = image.resize((card_width, card_height))
            sheet.paste(image,(x,y))
            create_mirrored_frame(x, y, sheet, image)
            draw_cutting_corners(x, y, draw)

        # Saving A4 sheet as PNG
        if not os.path.exists(f"{result_folder}/{time}"):
            try:
                os.mkdir(f"{result_folder}/{time}")
            except:
                os.mkdir(f"{result_folder}")
                os.mkdir(f"{result_folder}/{time}")
            
        sheet.save(f"{result_folder}/{time}/page_{sheets_total:03}.png", dpi=(300, 300))
        k = cards_total//(columns*rows)
        if cards_total%(columns*rows) != 0:
            k += 1
        print(f"{sheets_total}/{2*k+1} is ready")

    # Converting each sheet into PDF
    sheets_dirs = []
    for i in range(1, sheets_total + 1):
        Path(f"{result_folder}/{time}/page_{i:03}.pdf").write_bytes(img2pdf.convert(f"{result_folder}/{time}/page_{i:03}.png"))
        sheets_dirs.append(f"{result_folder}/{time}/page_{i:03}.pdf")
        print(f"{k+i}/{2*k+1} is ready")

    merge_sheets_into_pdf(sheets_dirs, time)
    print(f"{2*k+1}/{2*k+1} is ready")

def main():
    global paper_width, paper_height, card_width, card_height, \
        resize, rows, columns, gap, cards_folder, result_folder

    # Getting preset from file
    if os.path.exists(f"{curr_path}/PnP2PDF_Preferences.txt"):
        with open(f"{curr_path}/PnP2PDF_Preferences.txt", mode="r") as file:
            paper_width, paper_height, card_width, card_height, resize, rows, \
                columns, gap, cards_folder, result_folder = [line.strip() for line in file]
            resize = int(resize)

    # Dialog with user
    ui_output = ui_window()

    # Updating values from user input
    cards_folder = ui_output['-FOLDER1-']
    result_folder = ui_output['-FOLDER2-']
    paper_width = round(float(ui_output['-PWIDTH-'].replace(",", ".")) * PXLS_IN_MM)
    paper_height = round(float(ui_output['-PHEIGHT-'].replace(",", ".")) * PXLS_IN_MM)
    card_width = round(float(ui_output['-CWIDTH-'].replace(",", ".")) * PXLS_IN_MM)
    card_height = round(float(ui_output['-CHEIGHT-'].replace(",", ".")) * PXLS_IN_MM)
    gap = round(float(ui_output['-GAP-'].replace(",", ".")) * PXLS_IN_MM)
    resize = ui_output['-CHECKBOX-']
    if resize == True:
        card_height = round(0.98 * card_height)
        card_width = round(0.98 * card_width)
    rows = int(ui_output['-ROWS-'])
    columns = int(ui_output['-COLUMNS-'])

    # Main part
    if os.path.exists(cards_folder) and os.path.exists(result_folder):
        create_pdf()
    else:
        print("No such file or directory")

    # Saving preset
    with open(f"{curr_path}/PnP2PDF_Preferences.txt", mode="w") as file:
        file.write(f"{ui_output['-PWIDTH-']}\n{ui_output['-PHEIGHT-']}\n")
        file.write(f"{ui_output['-CWIDTH-']}\n{ui_output['-CHEIGHT-']}\n")
        file.write(f"{int(resize)}\n")
        file.write(f"{rows}\n{columns}\n")
        file.write(f"{ui_output['-GAP-']}\n")
        file.write(f"{cards_folder}\n{result_folder}")

main()