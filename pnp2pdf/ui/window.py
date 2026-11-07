"""CustomTkinter window with a live preview panel.

The module is imported by the test suite through a stub ``customtkinter`` and
on machines without Tk, so it must not touch ``tkinter`` at import time. Every
``tkinter`` use (file dialogs, message boxes) happens inside a method.
"""

import os
import queue
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

import customtkinter as ctk
from PIL import Image

from pnp2pdf import __version__ as PROGRAM_VERSION
from pnp2pdf import crash
from pnp2pdf.crash import install_handlers, log_exception, log_message, log_path
from pnp2pdf.errors import InputError
from pnp2pdf.options import JobOptions
from pnp2pdf.resources import resource_path, settings_path
from pnp2pdf.settings import create_save, load_save
from pnp2pdf.ui import form
from pnp2pdf.ui.job_runner import BuildJob
from pnp2pdf.ui.preview_controller import PreviewController

CONTACT_URL = "https://www.t.me/nNekoChan"
GUIDE_URL = ("https://github.com/Dannoster/PnP2PDF/tree/main"
             "#how-to-prepare-double-sided-cards")

PAPER_PRESETS = {
    "A4 210×297": (210, 297),
    "A3 297×420": (297, 420),
    "Letter 216×279": (216, 279),
    "Legal 216×356": (216, 356),
}

_ENTRY_KEYS = [
    "-PWIDTH-", "-PHEIGHT-", "-CWIDTH-", "-CHEIGHT-",
    "-ROWS-", "-COLUMNS-", "-GAP-", "-HSZONES-", "-VSZONES-",
    "-CFOLDER-", "-PDFFOLDER-",
]

_PREVIEW_WIDTH = 390
_PREVIEW_PICTURE_WIDTH = 340
_PREVIEW_PICTURE_HEIGHT = 420
_THUMB_WIDTH = 62
_THUMB_HEIGHT = 88

_GREY = ("gray40", "gray70")
_ERROR = ("#b23a2e", "#e06c5a")
_OK = ("#1f7a3d", "#3fb56a")

_FOLDER_CAPTIONS = {
    "-CFOLDER-": "Cards folder",
    "-PDFFOLDER-": "Save the PDF to",
}

_BLANK_PICTURE = None


def _blank_picture():
    """A 1x1 white CTkImage, so the label never points at a deleted Tk image."""
    global _BLANK_PICTURE
    if _BLANK_PICTURE is None:
        _BLANK_PICTURE = ctk.CTkImage(
            light_image=Image.new("RGB", (1, 1), (255, 255, 255)),
            dark_image=Image.new("RGB", (1, 1), (255, 255, 255)),
            size=(1, 1))
    return _BLANK_PICTURE


def _section(parent, title):
    """A rounded card holding a bold header; returns the frame for the body."""
    card = ctk.CTkFrame(parent, corner_radius=12)
    card.pack(fill="x", pady=(0, 10))
    header = ctk.CTkLabel(card, text=title,
                          font=ctk.CTkFont(size=13, weight="bold"))
    header.pack(anchor="w", padx=14, pady=(12, 6))
    body = ctk.CTkFrame(card, fg_color="transparent")
    body.pack(fill="x", padx=14, pady=(0, 12))
    return body


def _entry(parent, key, value, width=6, placeholder=""):
    """A fixed-width text entry pre-filled with ``value``."""
    entry = ctk.CTkEntry(parent, width=max(60, width * 12),
                         font=ctk.CTkFont(size=12),
                         placeholder_text=placeholder)
    entry.insert(0, str(value))
    return entry


def _fit_size(image, max_width, max_height):
    """Fit a PIL image into the box while keeping its aspect ratio."""
    width, height = image.size
    if not width or not height:
        return (max_width, max_height)
    scale = min(max_width / width, max_height / height)
    return (max(1, int(width * scale)), max(1, int(height * scale)))


class PnP2PDFWindow(ctk.CTk):
    """The one window: settings on the left, a live preview on the right."""

    def __init__(self, settings: dict):
        super().__init__()
        self.title("PnP2PDF")
        self.geometry("1080x720")
        self.minsize(1000, 680)
        try:
            self.iconbitmap(resource_path("icon.png"))
        except Exception:
            pass

        self.result = "Window closed"
        self._closed = False
        self._last_pdf = None
        self._entries = {}
        self._job = None
        self._build_started = 0.0
        self._build_events = queue.Queue()
        self._preview = PreviewController()
        self._preview_pending = False
        self._page_views = []
        self._page_thumbs = []
        self._thumb_buttons = []
        self._retired_images = []
        self._ui_failures = 0  # consecutive preview redraw failures
        self._page_index = 0

        self._build_ui(settings)
        self._log_path = log_path()
        self.report_callback_exception = self._report_callback_exception
        if self._entries["-PDFFOLDER-"].get().strip():
            self._output_switch.select()
            self._toggle_output_folder()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._tick)
        self._schedule_preview()

    # -- construction ---------------------------------------------------

    def _build_ui(self, settings):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0, minsize=_PREVIEW_WIDTH)
        self.grid_rowconfigure(0, weight=1)

        left = ctk.CTkScrollableFrame(self, corner_radius=0,
                                      fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(12, 6), pady=12)

        right = ctk.CTkFrame(self, width=_PREVIEW_WIDTH, corner_radius=12)
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 12), pady=12)

        self._build_left(left, settings)
        self._build_preview(right)

        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.grid(row=1, column=0, columnspan=2, sticky="ew",
                    padx=12, pady=(0, 12))
        self._build_bottom(bottom)

    def _register_entry(self, key, value, parent, width=6, placeholder=""):
        entry = _entry(parent, key, value, width, placeholder)
        entry.bind("<KeyRelease>", lambda event: self._schedule_preview())
        self._entries[key] = entry
        return entry

    def _build_left(self, parent, settings):
        initial = form.initial_values(settings)

        paper = _section(parent, "Paper")
        row = ctk.CTkFrame(paper, fg_color="transparent")
        row.pack(fill="x")
        self._register_entry("-PWIDTH-", initial["-PWIDTH-"], row).pack(side="left")
        ctk.CTkLabel(row, text="×").pack(side="left", padx=6)
        self._register_entry("-PHEIGHT-", initial["-PHEIGHT-"], row).pack(side="left")
        ctk.CTkLabel(row, text="mm").pack(side="left", padx=6)
        preset = ctk.CTkOptionMenu(paper, values=list(PAPER_PRESETS),
                                   command=self._apply_preset)
        preset.pack(fill="x", pady=(8, 0))
        preset.set("Presets")

        cards = _section(parent, "Cards")
        row = ctk.CTkFrame(cards, fg_color="transparent")
        row.pack(fill="x")
        self._register_entry("-CWIDTH-", initial["-CWIDTH-"], row).pack(side="left")
        ctk.CTkLabel(row, text="×").pack(side="left", padx=6)
        self._register_entry("-CHEIGHT-", initial["-CHEIGHT-"], row).pack(side="left")
        ctk.CTkLabel(row, text="mm").pack(side="left", padx=6)
        self._size_check = ctk.CTkCheckBox(cards, text="Smaller size (98%)",
                                           command=self._toggle_size)
        self._size_check.pack(anchor="w", pady=(8, 0))
        if initial["-SIZECHECKBOX-"]:
            self._size_check.select()

        layout = _section(parent, "Layout")
        row = ctk.CTkFrame(layout, fg_color="transparent")
        row.pack(fill="x")
        self._register_entry("-ROWS-", initial["-ROWS-"], row).pack(side="left")
        ctk.CTkLabel(row, text="×").pack(side="left", padx=6)
        self._register_entry("-COLUMNS-", initial["-COLUMNS-"], row).pack(side="left")
        ctk.CTkLabel(row, text="rows × columns").pack(side="left", padx=6)
        self._labeled_number(layout, "Space between", "-GAP-", initial["-GAP-"])
        self._labeled_number(layout, "Horizontal safe zones", "-HSZONES-",
                             initial["-HSZONES-"])
        self._labeled_number(layout, "Vertical safe zones", "-VSZONES-",
                             initial["-VSZONES-"])
        self._double_check = ctk.CTkCheckBox(layout, text="Double sided cards",
                                             command=self._toggle_double)
        self._double_check.pack(anchor="w", pady=(8, 0))
        if initial["-DOUBLECHECKBOX-"]:
            self._double_check.select()

        files = _section(parent, "Files")
        self._folder_row(files, "-CFOLDER-", initial["-CFOLDER-"],
                         placeholder="Where the card images are")

        self._output_switch = ctk.CTkCheckBox(
            files, text="Save the PDF to a different folder",
            command=self._toggle_output)
        self._output_switch.pack(anchor="w", pady=(8, 0))

        self._output_hint = ctk.CTkLabel(
            files, text="The PDF will be saved next to the cards.",
            anchor="w", text_color=_GREY)
        self._output_hint.pack(fill="x")

        self._output_row = self._folder_row(
            files, "-PDFFOLDER-", initial["-PDFFOLDER-"],
            placeholder="Folder for the PDF", pack=False)

    def _labeled_number(self, parent, label, key, value):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=(8, 0))
        ctk.CTkLabel(row, text=label).pack(side="left")
        self._register_entry(key, value, row).pack(side="left", padx=6)
        ctk.CTkLabel(row, text="mm").pack(side="left")

    def _folder_row(self, parent, key, value, placeholder="", pack=True):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        ctk.CTkLabel(row, text=_FOLDER_CAPTIONS[key], anchor="w").pack(fill="x")
        line = ctk.CTkFrame(row, fg_color="transparent")
        line.pack(fill="x")
        entry = self._register_entry(key, value, line, width=30,
                                     placeholder=placeholder)
        entry.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(line, text="Choose…", width=90,
                      command=lambda: self._choose_folder(key)).pack(
            side="left", padx=(6, 0))
        if pack:
            row.pack(fill="x", pady=(0, 8))
        return row

    def _toggle_output_folder(self):
        if self._output_switch.get():
            self._output_hint.pack_forget()
            self._output_row.pack(fill="x", pady=(0, 8))
        else:
            self._output_row.pack_forget()
            self._output_hint.pack(fill="x")

    def _log_toggle(self, key, variable):
        crash.log_event("toggle", name=key, value=bool(variable.get()))

    def _toggle_size(self):
        self._log_toggle("size_check", self._size_check)
        self._schedule_preview()

    def _toggle_double(self):
        self._log_toggle("double_check", self._double_check)
        self._schedule_preview()

    def _toggle_output(self):
        self._log_toggle("output_switch", self._output_switch)
        self._toggle_output_folder()

    def _build_bottom(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        bar = ctk.CTkFrame(parent, corner_radius=12)
        bar.grid(row=0, column=0, sticky="ew")
        bar.grid_columnconfigure(0, weight=1)

        self._progress = ctk.CTkProgressBar(bar, height=8)
        self._progress.set(0)
        self._progress.grid(row=0, column=0, columnspan=5, sticky="ew",
                            padx=12, pady=(12, 4))

        self._status = ctk.CTkLabel(bar, text="", anchor="w", text_color=_GREY)
        self._status.grid(row=1, column=0, sticky="w", padx=12, pady=(0, 12))

        self._run_button = ctk.CTkButton(
            bar, text="Create PDF", width=140, command=self._run,
            fg_color=("#3a7ebf", "#1f538d"))
        self._run_button.grid(row=1, column=1, padx=6, pady=(0, 12))
        self._cancel_button = ctk.CTkButton(
            bar, text="Cancel", width=90, state="disabled",
            command=self._cancel)
        self._cancel_button.grid(row=1, column=2, padx=6, pady=(0, 12))
        self._show_button = ctk.CTkButton(
            bar, text="Show PDF", width=100, state="disabled",
            command=self._reveal_pdf)
        self._show_button.grid(row=1, column=3, padx=6, pady=(0, 12))

        links = ctk.CTkFrame(bar, fg_color="transparent")
        links.grid(row=1, column=4, sticky="e", padx=12, pady=(0, 12))
        self._link(links, "Contact", CONTACT_URL)
        self._link(links, "Double-sided guide", GUIDE_URL)
        ctk.CTkLabel(links, text=f"v{PROGRAM_VERSION}").pack(
            side="left", padx=(8, 0))

    def _link(self, parent, text, url):
        label = ctk.CTkLabel(parent, text=text, text_color=("#1f6feb", "#4c9aff"))
        label.pack(side="left", padx=4)
        label.bind("<Button-1>", lambda event, u=url: webbrowser.open(u))

    def _build_preview(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 4))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Preview",
                     font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w")
        self._prev_button = ctk.CTkButton(header, text="‹", width=32,
                                          command=lambda: self._step_page(-1))
        self._prev_button.grid(row=0, column=1, padx=2)
        self._page_label = ctk.CTkLabel(header, text="1 / 1", width=48)
        self._page_label.grid(row=0, column=2)
        self._next_button = ctk.CTkButton(header, text="›", width=32,
                                          command=lambda: self._step_page(1))
        self._next_button.grid(row=0, column=3, padx=2)

        self._picture = ctk.CTkLabel(parent, text="")
        self._picture.grid(row=1, column=0, padx=14, pady=6)

        self._summary = ctk.CTkLabel(parent, text="", text_color=_GREY)
        self._summary.grid(row=2, column=0, padx=14, pady=(4, 8))

        self._thumbs = ctk.CTkFrame(parent, fg_color="transparent")
        self._thumbs.grid(row=3, column=0, padx=14, pady=(0, 8), sticky="ew")

        self._preview_status = ctk.CTkLabel(parent, text="", anchor="w",
                                            text_color=_GREY)
        self._preview_status.grid(row=4, column=0, sticky="ew", padx=14,
                                  pady=(0, 12))

        self._set_nav(1)

    # -- widget values ---------------------------------------------------

    def values(self):
        """The window contract dict, with checkboxes normalised to bool."""
        raw = {key: self._entries[key].get() for key in _ENTRY_KEYS}
        raw["-SIZECHECKBOX-"] = bool(self._size_check.get())
        raw["-DOUBLECHECKBOX-"] = bool(self._double_check.get())
        return form.check_values(raw)

    def _apply_preset(self, label):
        preset = PAPER_PRESETS.get(label)
        if preset is None:
            return
        for key, value in zip(("-PWIDTH-", "-PHEIGHT-"), preset):
            entry = self._entries[key]
            entry.delete(0, "end")
            entry.insert(0, str(value))
        self._schedule_preview()

    def _choose_folder(self, key):
        from tkinter import filedialog

        crash.log_event("click", button="choose-folder", row=key)
        current = self._entries[key].get()
        chosen = filedialog.askdirectory(initialdir=current or None)
        if chosen:
            crash.log_event("folder", row=key, path=chosen)
            entry = self._entries[key]
            entry.delete(0, "end")
            entry.insert(0, chosen)
            self._ui_failures = 0  # a new folder deserves a fresh attempt
            self._schedule_preview()

    # -- preview ---------------------------------------------------------

    def _schedule_preview(self):
        if self._closed:
            return
        if self._job is not None and self._job.is_running():
            self._preview_pending = True
            return
        if not self._entries["-CFOLDER-"].get().strip():
            self._clear_preview()
            self._preview_status.configure(
                text="Choose a cards folder to see the preview.",
                text_color=_GREY)
            return
        self._preview_status.configure(text="Rendering…", text_color=_GREY)
        self._preview.submit(self.values())

    def _tick(self):
        self._drain_build()
        self._drain_preview()
        if self._preview_pending and not (
                self._job is not None and self._job.is_running()):
            self._preview_pending = False
            self._schedule_preview()
        if self._preview.due():
            self._preview.start()
        if not self._closed:
            self.after(100, self._tick)

    def _drain_preview(self):
        result = self._preview.take()
        if result is None:
            return
        if result.error:
            crash.log_event("preview-failed", error=result.error)
            self._clear_preview()
            self._preview_status.configure(text=result.error, text_color=_ERROR)
            return
        self._show_pages(result.images, result.summary)

    def _show_image(self, image):
        if self._ui_failures > 2:
            self._preview_status.configure(
                text=f"Preview unavailable — details were written to "
                     f"{crash.log_path()}", text_color=_ERROR)
            return False
        try:
            self._picture.configure(image=image, text="")
        except Exception as error:
            # tkinter is not importable in the test environment, so catching
            # Exception here (rather than tkinter.TclError) is deliberate.
            self._ui_failures += 1
            crash.log_event("ui-failed", widget="preview",
                            attempt=self._ui_failures, error=error)
            self._retired_images = list(self._page_views) + list(self._page_thumbs)
            self._page_views = []
            self._page_thumbs = []
            try:
                self._picture.configure(image=_blank_picture(), text="")
            except Exception:
                pass
            if self._ui_failures <= 2:
                if self._entries["-CFOLDER-"].get().strip():
                    self._schedule_preview()
            else:
                self._preview_status.configure(
                    text=f"Preview unavailable — details were written to "
                         f"{crash.log_path()}", text_color=_ERROR)
            return False
        self._ui_failures = 0
        return True

    def _clear_preview(self):
        self._show_image(_blank_picture())
        self._retired_images = self._page_views + self._page_thumbs
        for button in self._thumb_buttons:
            button.destroy()
        self._thumb_buttons = []
        self._page_views = []
        self._page_thumbs = []
        self._summary.configure(text="")
        self._show_page(0)

    def _show_pages(self, images, summary):
        crash.log_event("preview", pages=len(images or []), summary=summary)
        self._show_image(_blank_picture())
        self._retired_images = self._page_views + self._page_thumbs
        self._preview_status.configure(text="", text_color=_GREY)
        self._summary.configure(text=summary)
        for button in self._thumb_buttons:
            button.destroy()
        self._thumb_buttons = []
        self._page_views = []
        self._page_thumbs = []
        for index, image in enumerate(images or []):
            self._page_views.append(self._ctk_image(
                image, _PREVIEW_PICTURE_WIDTH, _PREVIEW_PICTURE_HEIGHT))
            thumb = self._ctk_image(image, _THUMB_WIDTH, _THUMB_HEIGHT)
            self._page_thumbs.append(thumb)
            button = ctk.CTkButton(self._thumbs, text="", image=thumb,
                                   width=_THUMB_WIDTH, height=_THUMB_HEIGHT,
                                   command=lambda i=index: self._page_clicked(i))
            button.pack(side="left", padx=3)
            self._thumb_buttons.append(button)
        self._page_index = 0
        self._show_page(0)

    def _ctk_image(self, image, max_width, max_height):
        return ctk.CTkImage(light_image=image, dark_image=image,
                            size=_fit_size(image, max_width, max_height))

    def _show_page(self, index):
        if not self._page_views:
            self._show_image(_blank_picture())
            self._page_label.configure(text="1 / 1")
            self._set_nav(1)
            return
        index = max(0, min(index, len(self._page_views) - 1))
        self._page_index = index
        self._show_image(self._page_views[index])
        self._page_label.configure(text=f"{index + 1} / {len(self._page_views)}")
        self._set_nav(len(self._page_views))

    def _step_page(self, delta):
        self._page_clicked(self._page_index + delta)

    def _page_clicked(self, index):
        if self._page_views:
            index = max(0, min(index, len(self._page_views) - 1))
        else:
            index = 0
        crash.log_event("click", button="page", index=index)
        self._show_page(index)

    def _set_nav(self, count):
        state = "disabled" if count <= 1 else "normal"
        self._prev_button.configure(state=state)
        self._next_button.configure(state=state)

    # -- build job -------------------------------------------------------

    def _run(self):
        crash.log_event("click", button="create")
        try:
            if self._job is not None and self._job.is_running():
                return
            try:
                options = JobOptions.from_form(self.values())
                options.validate()
            except InputError as error:
                self._show_error(str(error))
                return

            crash.log_event("build-start", cards=options.cards_folder,
                            double_sided=options.double_sided)
            self._build_started = time.monotonic()
            self._run_button.configure(state="disabled")
            self._cancel_button.configure(state="normal")
            self._status.configure(text="")
            self._progress.set(0)

            self._job = BuildJob(
                options,
                lambda name, payload: self._build_events.put((name, payload)))
            self._job.start()
        except Exception:
            self._log_failure("start a build")

    def _drain_build(self):
        try:
            while True:
                try:
                    name, payload = self._build_events.get_nowait()
                except queue.Empty:
                    break
                if name == "progress":
                    done, total, text = payload
                    self._progress.set(done / total if total else 0)
                    self._status.configure(text=text)
                elif name == "done":
                    self.result = self.values()
                    self._last_pdf = self._pdf_file(payload)
                    shown = self._last_pdf.name if self._last_pdf else str(payload)
                    crash.log_event(
                        "build-done", pdf=self._last_pdf,
                        seconds=round(time.monotonic() - self._build_started, 1))
                    self._finish_run(reset_progress=False)
                    self._progress.set(1)
                    self._status.configure(text=f"Saved to {shown}",
                                           text_color=_OK)
                    self._show_button.configure(state="normal")
                elif name == "error":
                    crash.log_event("build-failed", error=payload)
                    self._finish_run()
                    self._status.configure(text="Failed")
                    self._show_error(payload)
                elif name == "cancelled":
                    self._finish_run()
                    self._status.configure(text="Cancelled")
        except Exception:
            self._finish_run()
            self._log_failure("event pump")

    def _log_failure(self, where):
        path = log_exception(where, *sys.exc_info())
        self._status.configure(text=f"Failed — details were written to {path}")
        return path

    def _finish_run(self, reset_progress=True):
        self._run_button.configure(state="normal")
        self._cancel_button.configure(state="disabled")
        if reset_progress:
            self._progress.set(0)

    def _pdf_file(self, folder):
        path = Path(folder)
        if path.is_file():
            return path
        matches = sorted(path.glob("*.pdf")) or sorted(path.glob("*/*.pdf"))
        return matches[0] if matches else None

    def _reveal_pdf(self):
        crash.log_event("click", button="show-pdf")
        path = self._last_pdf
        if not path or not os.path.exists(path):
            return
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", "-R", str(path)])
            elif os.name == "nt":
                subprocess.Popen(["explorer", "/select,", str(path)])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(str(path))])
        except OSError:
            pass

    def _cancel(self):
        crash.log_event("click", button="cancel")
        if self._job is not None:
            self._job.cancel()
        self._cancel_button.configure(state="disabled")
        self._status.configure(text="Cancelling…")

    def _report_callback_exception(self, *args):
        log_exception("tk callback", *sys.exc_info())
        error = sys.exc_info()[1]
        self._show_error(str(error) if error else "An unexpected error occurred.")

    def _show_error(self, message):
        from tkinter import messagebox

        path = getattr(self, "_log_path", None)
        if path:
            message = f"{message}\n\nDetails were written to {path}"
        messagebox.showerror("PnP2PDF", message)

    # -- life cycle ------------------------------------------------------

    def _mark_closed(self):
        if not self._closed:
            crash.log_event("window-close")
        self._closed = True

    def _close(self):
        self._mark_closed()
        self._preview.stop()
        self.destroy()

    def _on_close(self):
        self._mark_closed()
        self._preview.stop()
        if self._job is not None and self._job.is_running():
            self._job.cancel()
            self._job.join(5)
        self.result = "Window closed"
        self.destroy()


def ui_window(settings: dict):
    ctk.set_appearance_mode("system")
    window = PnP2PDFWindow(settings)
    window.mainloop()
    return window.result


def main():
    log_file = install_handlers()
    save_folder = settings_path()
    log_message(
        f"PnP2PDF {PROGRAM_VERSION} starting | "
        f"frozen={getattr(sys, 'frozen', False)} | "
        f"settings={save_folder} | log={log_file}")
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
