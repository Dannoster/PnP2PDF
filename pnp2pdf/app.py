import os
import webbrowser
import threading
# pip install PySimpleGUI
# pip install pillow
import PySimpleGUI as sg

from pnp2pdf import __version__ as PROGRAM_VERSION

# Re-exported for backwards compatibility, so that pnp2pdf.app.<name>
# keeps resolving for code that imports it from here.
from pnp2pdf.errors import OperationCancelled
from pnp2pdf.model import Card, CardBack, Sheet
from pnp2pdf.pdf import PdfFile
from pnp2pdf.cards import (
    IMAGE_EXTENSIONS, create_cards, find_back_dirs, find_card_dirs, natural_key,
)
from pnp2pdf.layout import create_sheets
from pnp2pdf.render import (
    CUT_LINE_COLOUR, PXLS_IN_MM, create_mirrored_frame, draw_cutting_corners,
    draw_lines, draw_szones, place_cropped_back,
)
from pnp2pdf.resources import ICC_DIR, curr_path, icc_profile
from pnp2pdf.settings import DEFAULT_SETTINGS, create_save, load_save


def build_pdf(values: dict, progress=None, cancelled=None) -> str:
    cards = create_cards(values['-CFOLDER-'],
                         float(values['-CWIDTH-'].replace(',', '.')),
                         float(values['-CHEIGHT-'].replace(',', '.')),
                         values["-SIZECHECKBOX-"],
                         values['-DOUBLECHECKBOX-'])
    sheets = create_sheets(cards,
                           float(values['-PWIDTH-'].replace(',', '.')),
                           float(values['-PHEIGHT-'].replace(',', '.')),
                           int(values['-ROWS-']),
                           int(values['-COLUMNS-']),
                           float(values['-GAP-'].replace(',', '.')),
                           float(values['-HSZONES-'].replace(',', '.')),
                           float(values['-VSZONES-'].replace(',', '.')))
    pdf = PdfFile(sheets)
    return pdf.save(values['-PDFFOLDER-'], values['-DOUBLECHECKBOX-'],
                    progress=progress, cancelled=cancelled)


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
                  sg.Text('Prepare files for double sided print', enable_events = True, 
                          font = ('Courier New', 10, 'underline'), 
                          key='URL https://github.com/Dannoster/PnP2PDF/tree/main#how-to-prepare-double-sided-cards')],

                 [sg.Text("Hor. safe zones", size=(14, 1)),
                  sg.InputText(f"{settings['ud_szone']}", size=(6, 1), key="-HSZONES-"), sg.Text('mm\t'),
                  sg.Text("Vert. safe zones", size=(14, 1)),
                  sg.InputText(f"{settings['lr_szone']}", size=(6, 1), key="-VSZONES-"), sg.Text('mm')], 

                 [sg.Text('Cards folder', size=(14, 1)), 
                  sg.InputText(f"{settings['cards_folder']}", key='-CFOLDER-'), 
                  sg.FolderBrowse(initial_folder=settings['cards_folder'])],

                 [sg.Text('Save PDF to', size=(14, 1)), 
                  sg.InputText(f"{settings['save_to']}", key='-PDFFOLDER-'), 
                  sg.FolderBrowse(initial_folder=settings['save_to'])],

                 [sg.ProgressBar(1000, orientation="h", size=(20, 20), key="-PROGRESS-")],
                 [sg.Text("", size=(40, 1), key="-STATUS-")],
                 [sg.Button("Create PDF", key="-RUN-", size=(14, 1)),
                  sg.Button("Cancel", key="-CANCEL-", disabled=True),
                  sg.Push(), sg.Text(f"v{PROGRAM_VERSION}")]
                ]

    window = sg.Window('PnP2PDF', form_rows, icon="icon.png")

    cancel_event = threading.Event()
    worker = None
    job_values = {}

    while True:
        event, values = window.read(timeout=100)
        if event is None:
            continue
        if event == sg.WINDOW_CLOSED:
            if worker is not None and worker.is_alive():
                cancel_event.set()
                worker.join(timeout=5)
            return "Window closed"
        elif event.startswith("URL "):
            url = event.split(' ')[1]
            webbrowser.open(url)
        elif event == "-RUN-":
            if worker is not None and worker.is_alive():
                continue
            cancel_event.clear()
            window["-RUN-"].update(disabled=True)
            window["-CANCEL-"].update(disabled=False)
            window["-STATUS-"].update("")
            window["-PROGRESS-"].update(0)
            job_values = dict(values)

            def run_job(job_values):
                try:
                    folder = build_pdf(
                        job_values,
                        progress=lambda done, total, text: window.write_event_value(
                            "-PROGRESS-", (done, total, text)),
                        cancelled=cancel_event.is_set)
                    window.write_event_value("-DONE-", folder)
                except OperationCancelled:
                    window.write_event_value("-CANCELLED-", None)
                except Exception as error:
                    window.write_event_value("-ERROR-", str(error))

            worker = threading.Thread(target=run_job, args=(job_values,), daemon=True)
            worker.start()
        elif event == "-PROGRESS-":
            done, total, text = values["-PROGRESS-"]
            window["-PROGRESS-"].update(done * 1000 // max(total, 1))
            window["-STATUS-"].update(text)
        elif event == "-DONE-":
            window.close()
            return job_values
        elif event == "-ERROR-":
            sg.popup_error(f"Could not create the PDF:\n\n{values['-ERROR-']}")
            window["-RUN-"].update(disabled=False)
            window["-CANCEL-"].update(disabled=True)
            window["-STATUS-"].update("Failed")
        elif event == "-CANCELLED-":
            window["-RUN-"].update(disabled=False)
            window["-CANCEL-"].update(disabled=True)
            window["-PROGRESS-"].update(0)
            window["-STATUS-"].update("Cancelled")
        elif event == "-CANCEL-":
            cancel_event.set()
            window["-CANCEL-"].update(disabled=True)
            window["-STATUS-"].update("Cancelling…")

        
def main():
    save_folder = os.path.join(curr_path, "presets.json")
    settings = load_save(save_folder)
    ui_output = ui_window(settings)
    if ui_output != "Window closed":
        settings = {"paper_width"   : ui_output['-PWIDTH-'],
                    "paper_height"  : ui_output['-PHEIGHT-'],
                    "card_width"    : ui_output['-CWIDTH-'],
                    "card_height"   : ui_output['-CHEIGHT-'],
                    "resize"        : ui_output['-SIZECHECKBOX-'],
                    "double_sided"  : ui_output['-DOUBLECHECKBOX-'],
                    "rows"          : ui_output['-ROWS-'],
                    "columns"       : ui_output['-COLUMNS-'],
                    "gap"           : ui_output['-GAP-'],
                    "ud_szone"      : ui_output['-HSZONES-'], 
                    "lr_szone"      : ui_output['-VSZONES-'],
                    "cards_folder"  : ui_output['-CFOLDER-'],
                    "save_to"       : ui_output['-PDFFOLDER-']}
        create_save(save_folder, settings)


if __name__ == "__main__":
    main()
