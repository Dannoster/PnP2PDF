"""Record fatal errors and the trail of actions in a file the user can send us.

The log holds two kinds of entry: the actions the user took before something
went wrong (``log_event``) and the tracebacks of the errors themselves
(``log_exception``). It never grows past ``MAX_BYTES``: once an append would
pass the cap the live file is renamed aside to ``crash.log.1``, so at most two
files exist.
"""

import datetime
import os
import sys
import threading
import traceback

from pnp2pdf import resources

LOG_NAME = "crash.log"
MAX_BYTES = 320 * 1024
BACKUP_SUFFIX = ".1"

_MAX_FIELD = 120

_installed = False


def log_path() -> str:
    """Absolute path of the crash log, inside the application data directory."""
    return os.path.join(resources.data_dir(), LOG_NAME)


def _stamp() -> str:
    now = datetime.datetime.now().isoformat(sep=" ", timespec="seconds")
    return f"{now} [pid {os.getpid()}]"


def _rotate(path: str, incoming: int) -> None:
    """Rename the live log aside once an append would pass ``MAX_BYTES``."""
    try:
        if os.path.exists(path) and os.path.getsize(path) + incoming > MAX_BYTES:
            os.replace(path, path + BACKUP_SUFFIX)
    except OSError:
        pass


def _append(header: str, body: str = "") -> str:
    """Append one timestamped entry; never raise, whatever the filesystem says."""
    path = log_path()
    text = f"{_stamp()} {header}\n{body}"
    try:
        directory = os.path.dirname(path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        _rotate(path, len(text))
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


def _field(value) -> str:
    """One field value: no line breaks, and short enough not to blow up a line."""
    text = str(value)
    for char in ("\r\n", "\n", "\r", "\t"):
        text = text.replace(char, " ")
    if len(text) > _MAX_FIELD:
        text = text[:_MAX_FIELD] + "…"
    return text


def log_event(event: str, /, **fields) -> str:
    """Append one action line, e.g. ``click button=create``; return the path."""
    try:
        parts = [str(event)]
        parts.extend(f"{key}={_field(value)}" for key, value in fields.items())
        line = " ".join(parts)
    except Exception:
        return ""
    return _append(line)


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
