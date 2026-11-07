from pnp2pdf.cards import create_cards
from pnp2pdf.layout import create_sheets
from pnp2pdf.options import JobOptions
from pnp2pdf.pdf import PdfFile


def create_pdf(options: JobOptions, progress=None, cancelled=None) -> str:
    options.validate()
    cards = create_cards(options.cards_folder,
                         options.card_width,
                         options.card_height,
                         options.smaller_size,
                         options.double_sided)
    sheets = create_sheets(cards,
                           options.paper_width,
                           options.paper_height,
                           options.rows,
                           options.columns,
                           options.gap,
                           options.ud_szone,
                           options.lr_szone)
    pdf = PdfFile(sheets)
    return pdf.save(options.output_folder, options.double_sided,
                    progress=progress, cancelled=cancelled)
