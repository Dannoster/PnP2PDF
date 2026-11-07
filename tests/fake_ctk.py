"""A tiny in-memory stand-in for CustomTkinter, for driving the window in CI.

The real toolkit needs a display (and Tk), so it cannot run here. This module
installs dumb widget objects into ``sys.modules`` so the genuine
``pnp2pdf.ui.window`` code can be exercised without a screen.

Everything is deliberately minimal: widgets remember the options they were
given, ``after``/``update`` provide a fake event loop, and the handful of
helpers at the bottom exist only so tests can inspect the tree.
"""

import sys
import time
import types


def _walk(widget):
    yield widget
    for child in list(getattr(widget, "children", [])):
        yield from _walk(child)


class _Widget:
    """The common parts of every fake widget: options, children, geometry."""

    def __init__(self, master=None, **kwargs):
        self.master = master
        self.options = dict(kwargs)
        self.children = []
        self.text = kwargs.get("text", "")
        self.image = kwargs.get("image", None)
        self.state = kwargs.get("state", "normal")
        self.command = kwargs.get("command", None)
        self.bindings = {}
        self.destroyed = False
        if isinstance(master, _Widget):
            master.children.append(self)

    # -- geometry --------------------------------------------------------

    def pack(self, **kwargs):
        self.options["_pack"] = kwargs
        return self

    def grid(self, **kwargs):
        self.options["_grid"] = kwargs
        return self

    def grid_propagate(self, flag=True):
        self.options["_grid_propagate"] = flag

    def grid_rowconfigure(self, *args, **kwargs):
        pass

    def grid_columnconfigure(self, *args, **kwargs):
        pass

    def grid_size(self):
        return (0, 0)

    # -- options ---------------------------------------------------------

    def configure(self, **kwargs):
        if self.destroyed:
            return
        self.options.update(kwargs)
        if "text" in kwargs:
            self.text = kwargs["text"]
        if "image" in kwargs:
            self.image = kwargs["image"]
        if "state" in kwargs:
            self.state = kwargs["state"]

    config = configure

    def cget(self, key):
        return self.options.get(key)

    def bind(self, sequence, func, add=None):
        self.bindings[sequence] = func

    def winfo_children(self):
        return list(self.children)

    def destroy(self):
        if self.destroyed:
            return
        self.destroyed = True
        for child in list(self.children):
            child.destroy()
        self.children = []
        master = self.master
        if isinstance(master, _Widget) and self in master.children:
            master.children.remove(self)


class CTk(_Widget):
    """The window: a widget that owns the fake event loop."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title_text = None
        self.geometry_spec = None
        self.min_size = None
        self.icon = None
        self.on_close = None
        self._callbacks = []
        self._next_id = 0

    def title(self, text=None):
        self.title_text = text

    def geometry(self, spec=None):
        self.geometry_spec = spec

    def minsize(self, width=None, height=None):
        self.min_size = (width, height)

    def iconbitmap(self, *args, **kwargs):
        self.icon = args

    # -- fake event loop -------------------------------------------------

    def after(self, ms, func):
        if self.destroyed:
            return None
        callback_id = self._next_id
        self._next_id += 1
        deadline = time.monotonic() + ms / 1000.0
        self._callbacks.append((deadline, callback_id, func))
        return callback_id

    def update(self):
        now = time.monotonic()
        due = [item for item in self._callbacks if item[0] <= now]
        for _, callback_id, func in sorted(due, key=lambda item: item[0]):
            if self.destroyed:
                break
            self._callbacks = [item for item in self._callbacks
                               if item[1] != callback_id]
            func()

    def mainloop(self):
        deadline = time.monotonic() + 10.0
        while not self.destroyed and time.monotonic() < deadline:
            self.update()
            time.sleep(0.005)

    def protocol(self, name, func):
        if name == "WM_DELETE_WINDOW":
            self.on_close = func

    def update_idletasks(self):
        pass


class CTkFrame(_Widget):
    pass


class CTkScrollableFrame(CTkFrame):
    pass


class CTkLabel(_Widget):
    pass


class CTkEntry(_Widget):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self._buffer = ""

    def insert(self, index, text):
        if index == "end":
            self._buffer += text
            return
        position = int(index)
        self._buffer = (self._buffer[:position] + text
                        + self._buffer[position:])

    def delete(self, first, last=None):
        if last is None:
            position = int(first)
            self._buffer = self._buffer[:position] + self._buffer[position + 1:]
            return
        if first == 0 and last == "end":
            self._buffer = ""
            return
        start = int(first)
        end = len(self._buffer) if last == "end" else int(last)
        self._buffer = self._buffer[:start] + self._buffer[end:]

    def get(self):
        return self._buffer


class CTkButton(_Widget):
    def invoke(self):
        if callable(self.command):
            return self.command()


class CTkCheckBox(_Widget):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self._selected = bool(kwargs.get("selected", False))

    def select(self):
        self._selected = True

    def deselect(self):
        self._selected = False

    def get(self):
        return self._selected


class CTkProgressBar(_Widget):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self._value = 0

    def set(self, value):
        self._value = value

    def get(self):
        return self._value


class CTkOptionMenu(_Widget):
    def __init__(self, master=None, values=None, command=None, **kwargs):
        super().__init__(master, command=command, **kwargs)
        self.values = list(values or [])
        self._value = None

    def set(self, value):
        self._value = value

    def get(self):
        return self._value

    def invoke(self, value=None):
        if value is not None:
            self._value = value
        if callable(self.command):
            return self.command(self._value)


class CTkImage:
    def __init__(self, light_image=None, dark_image=None, size=None, **kwargs):
        self.light_image = light_image
        self.dark_image = dark_image
        self.size = size


class CTkFont:
    def __init__(self, size=None, weight=None, **kwargs):
        self.size = size
        self.weight = weight


def set_appearance_mode(mode=None):
    pass


def set_default_color_theme(theme=None):
    pass


def set_widget_scaling(scaling=None):
    pass


def find_by_text(root, text):
    """The first widget (depth-first) whose ``text`` equals ``text``."""
    for widget in _walk(root):
        if getattr(widget, "text", None) == text:
            return widget
    return None


def all_images(root):
    """Every non-``None`` ``image`` held by the widget tree."""
    images = []
    for widget in _walk(root):
        image = getattr(widget, "image", None)
        if image is not None:
            images.append(image)
    return images


def _message_box_module():
    module = types.ModuleType("tkinter.messagebox")
    module.showerror_calls = []

    def showerror(title="", message="", **kwargs):
        module.showerror_calls.append((title, message))
        return "ok"

    module.showerror = showerror
    return module


def _file_dialog_module():
    module = types.ModuleType("tkinter.filedialog")
    module.askdirectory_calls = []
    module.directory = None

    def askdirectory(**kwargs):
        module.askdirectory_calls.append(kwargs)
        return module.directory

    module.askdirectory = askdirectory
    return module


def install():
    """Insert the fake toolkit into ``sys.modules`` and return its handles."""
    ctk_module = types.ModuleType("customtkinter")
    for name, value in {
        "CTk": CTk,
        "CTkFrame": CTkFrame,
        "CTkScrollableFrame": CTkScrollableFrame,
        "CTkLabel": CTkLabel,
        "CTkEntry": CTkEntry,
        "CTkButton": CTkButton,
        "CTkCheckBox": CTkCheckBox,
        "CTkProgressBar": CTkProgressBar,
        "CTkOptionMenu": CTkOptionMenu,
        "CTkImage": CTkImage,
        "CTkFont": CTkFont,
        "set_appearance_mode": set_appearance_mode,
        "set_default_color_theme": set_default_color_theme,
        "set_widget_scaling": set_widget_scaling,
    }.items():
        setattr(ctk_module, name, value)

    filedialog = _file_dialog_module()
    messagebox = _message_box_module()
    tkinter = types.ModuleType("tkinter")
    tkinter.filedialog = filedialog
    tkinter.messagebox = messagebox

    sys.modules["customtkinter"] = ctk_module
    sys.modules["tkinter"] = tkinter
    sys.modules["tkinter.filedialog"] = filedialog
    sys.modules["tkinter.messagebox"] = messagebox

    return {
        "customtkinter": ctk_module,
        "tkinter": tkinter,
        "filedialog": filedialog,
        "messagebox": messagebox,
    }
