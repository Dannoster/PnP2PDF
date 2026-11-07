"""Record fatal errors in a file the user can send us."""

import datetime
import os
import sys
import threading
import traceback

from pnp2pdf import resources

LOG_NAME = "crash.log"

_installed = False


def log_path() -> str:
    """Absolute path of the crash log, inside the application data directory."""
    return os.path.join(resources.data_dir(), LOG_NAME)


def _stamp() -> str:
    now = datetime.datetime.now().isoformat(sep=" ", timespec="seconds")
    return f"{now} [pid {os.getpid()}]"


def _append(header: str, body: str = "") -> str:
    """Append one timestamped entry; never raise, whatever the filesystem says."""
    path = log_path()
    text = f"{_stamp()} {header}\n{body}"
    try:
        directory = os.path.dirname(path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        with open(path, "a", encoding="utf-8") as file:
            file.write(text)
    except Exception:
        stream = sys.__stderr__
        if stream is not None:
            try:
                stream.write(text)
                stream.flush()
            except Exception:
                pass
    return path


def log_exception(where: str, exc_type, exc, tb) -> str:
    """Append one timestamped traceback; return the path of the log file."""
    body = "".join(traceback.format_exception(exc_type, exc, tb))
    return _append(where, body)


def log_message(text: str) -> str:
    """Append one plain line (used at startup); return the path."""
    return _append(text)


def install_handlers() -> str:
    """Hook sys.excepthook and threading.excepthook; return the log path."""
    global _installed
    if _installed:
        return log_path()

    previous_excepthook = sys.excepthook
    previous_threading_hook = getattr(threading, "excepthook", None)

    def _excepthook(exc_type, exc, tb):
        log_exception("uncaught exception", exc_type, exc, tb)
        if previous_excepthook is not None:
            previous_excepthook(exc_type, exc, tb)

    def _threading_excepthook(args):
        log_exception("uncaught in worker thread",
                      args.exc_type, args.exc_value, args.exc_traceback)
        if previous_threading_hook is not None:
            previous_threading_hook(args)

    sys.excepthook = _excepthook
    if hasattr(threading, "excepthook"):
        threading.excepthook = _threading_excepthook
    _installed = True
    return log_path()
