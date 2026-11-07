import threading

from pnp2pdf.errors import OperationCancelled


class BuildJob:
    """Runs one `create_pdf` call on a worker thread, reporting through `post`."""

    def __init__(self, options, post, build=None):
        self.options = options
        self.post = post
        self._build = build
        self._cancelled = threading.Event()
        self._thread = None
        self.path = None

    def start(self):
        if self.is_running():
            return False
        build = self._build
        if build is None:
            from pnp2pdf.job import create_pdf
            build = create_pdf
        self._thread = threading.Thread(target=self._run, args=(build,),
                                        daemon=True)
        self._thread.start()
        return True

    def _run(self, build):
        try:
            self.path = build(self.options, progress=self._progress,
                              cancelled=self._cancelled.is_set)
            self.post("done", self.path)
        except OperationCancelled:
            self.post("cancelled", None)
        except Exception as error:
            self.post("error", str(error))

    def _progress(self, done, total, text):
        self.post("progress", (done, total, text))

    def cancel(self):
        self._cancelled.set()

    def is_running(self):
        return self._thread is not None and self._thread.is_alive()

    def join(self, timeout=None):
        if self._thread is not None:
            self._thread.join(timeout)
