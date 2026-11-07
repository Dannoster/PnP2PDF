import os
import datetime

from PIL import Image, ImageDraw, ImageCms

from pnp2pdf.errors import OperationCancelled
from pnp2pdf.model import Sheet
from pnp2pdf.render import (
    PXLS_IN_MM, create_mirrored_frame, draw_cutting_corners, draw_lines,
    draw_szones, place_cropped_back,
)
from pnp2pdf.resources import icc_profile


class PdfFile:
    def __init__(self, sheets: list[Sheet]):
        self.sheets = sheets

    def save(self, folder: str, double_sided: bool, progress=None, cancelled=None) -> str:

        stamp = datetime.datetime.today().strftime("%d%m%y_%H%M")
        outdir = os.path.join(folder, stamp)
        counter = 2
        while os.path.exists(outdir):
            outdir = os.path.join(folder, f"{stamp}_{counter}")
            counter += 1
        os.makedirs(outdir, exist_ok=True)

        sht = self.sheets[0]
        crd = sht.cards[0]
        swidth, sheight = int(sht.width*PXLS_IN_MM), int(sht.height*PXLS_IN_MM)
        hszone, vszone = int(sht.ud_szone*PXLS_IN_MM), int(sht.lr_szone*PXLS_IN_MM)
        cwidth, cheight = int(crd.width*PXLS_IN_MM), int(crd.height*PXLS_IN_MM)
        rows, columns, gap = sht.rows, sht.columns, int(sht.gap*PXLS_IN_MM)
        x0 = (swidth - (columns-1)*gap - columns*cwidth) // 2
        y0 = (sheight - (rows-1)*gap - rows*cheight) // 2

        total = len(self.sheets) * (2 if double_sided else 1)
        pages = []
        done = 0

        def check_cancelled():
            if cancelled is not None and cancelled():
                for page in pages:
                    page.close()
                try:
                    os.rmdir(outdir)
                except OSError:
                    pass
                raise OperationCancelled()

        if progress is not None:
            progress(0, total, f"Page 0 of {total}")

        for i, sheet in enumerate(self.sheets):
            check_cancelled()
            page_a = Image.new('CMYK',
                               (int(sheet.width*PXLS_IN_MM), int(sheet.height*PXLS_IN_MM)),
                               (0, 0, 0, 0))
            draw_a = ImageDraw.Draw(page_a)
            draw_lines(draw_a, swidth, sheight, rows, columns, gap,
                       cwidth, cheight, number_of_cards=len(sheet.cards))
            for j, card in enumerate(sheet.cards):
                card_image = Image.open(card.dir).resize((int(card.width*PXLS_IN_MM),
                                                          int(card.height*PXLS_IN_MM)))
                try:
                    card_image = ImageCms.profileToProfile(card_image,
                                                           icc_profile("sRGB-IEC61966-2.1.icc"),
                                                           icc_profile("USWebCoatedSWOP.icc"),
                                                           outputMode='CMYK')
                except Exception as error:
                    print(f"WARNING: colour profile not applied to "
                          f"{os.path.basename(card.dir)}: {error}")
                x = x0 + j%columns * (cwidth + gap)
                y = y0 + j//columns * (cheight + gap)
                page_a.paste(card_image, (x, y))
                create_mirrored_frame(x, y, page_a, card_image, gap, cwidth, cheight)
                draw_cutting_corners(x, y, draw_a, cwidth, cheight)
            draw_szones(draw_a, swidth, sheight, hszone, vszone)
            pages.append(page_a)
            done += 1
            if progress is not None:
                progress(done, total, f"Page {done} of {total}")

            if double_sided:
                check_cancelled()
                page_b = Image.new('CMYK',
                                   (int(sheet.width*PXLS_IN_MM), int(sheet.height*PXLS_IN_MM)),
                                   (0, 0, 0, 0))
                for j, card in enumerate(sheet.cards):
                    if card.back:
                        back_image = Image.open(card.back.dir).resize(
                            (int(card.back.width*PXLS_IN_MM),
                             int(card.back.height*PXLS_IN_MM)))
                        try:
                            back_image = ImageCms.profileToProfile(back_image,
                                                                   icc_profile("sRGB-IEC61966-2.1.icc"),
                                                                   icc_profile("USWebCoatedSWOP.icc"),
                                                                   outputMode='CMYK')
                        except Exception as error:
                            print(f"WARNING: colour profile not applied to "
                                  f"{os.path.basename(card.back.dir)}: {error}")
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
                pages.append(page_b)
                done += 1
                if progress is not None:
                    progress(done, total, f"Page {done} of {total}")

        pdf_path = os.path.join(outdir, f"{os.path.basename(outdir)}.pdf")
        pages[0].save(pdf_path, "PDF", optimize=False,
                      save_all=True, append_images=pages[1:],
                      dpi=(300, 300), compression=None, quality=100)
        for page in pages:
            page.close()
        return outdir
