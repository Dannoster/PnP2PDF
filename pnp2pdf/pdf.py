import os
import datetime

from pnp2pdf.errors import OperationCancelled
from pnp2pdf.model import Sheet
from pnp2pdf.render import PXLS_IN_MM, render_back, render_front


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
            page_a = render_front(sheet, PXLS_IN_MM)
            pages.append(page_a)
            done += 1
            if progress is not None:
                progress(done, total, f"Page {done} of {total}")

            if double_sided:
                check_cancelled()
                page_b = render_back(sheet, PXLS_IN_MM)
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
