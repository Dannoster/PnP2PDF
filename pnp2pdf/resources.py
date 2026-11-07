import os
import sys


def _package_dir() -> str:
    return os.path.dirname(os.path.realpath(__file__))


def _resource_roots() -> list[str]:
    """Candidate directories that may hold bundled data, in priority order."""
    roots = [_package_dir()]

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        roots.append(meipass)

    frozen = getattr(sys, "frozen", False)
    if frozen:
        roots.append(os.path.dirname(sys.executable))

    if meipass:
        roots.append(os.path.normpath(os.path.join(meipass, os.pardir, "Resources")))
    if frozen:
        roots.append(os.path.normpath(
            os.path.join(os.path.dirname(sys.executable), os.pardir, "Resources")))

    unique = []
    for root in roots:
        if root and root not in unique:
            unique.append(root)
    return unique


def resource_path(name: str) -> str:
    """Absolute path to a bundled data file, or raise FileNotFoundError."""
    roots = _resource_roots()
    for root in roots:
        candidate = os.path.join(root, name)
        if os.path.exists(candidate):
            return candidate
    raise FileNotFoundError(
        f"Could not find resource {name!r}; tried:\n  "
        + "\n  ".join(roots))


def icc_profile(name: str) -> str:
    """Absolute path to a bundled ICC profile (works regardless of CWD)."""
    return resource_path(os.path.join("color_profiles", name))


def data_dir(platform: str | None = None, home: str | None = None,
             env: dict | None = None) -> str:
    """Directory where user settings belong, per platform."""
    if platform is None:
        platform = sys.platform
    if home is None:
        home = os.path.expanduser("~")
    if env is None:
        env = os.environ

    if not getattr(sys, "frozen", False):
        return os.path.dirname(_package_dir())

    if platform == "darwin":
        return os.path.join(home, "Library", "Application Support", "PnP2PDF")
    if platform == "win32":
        appdata = env.get("APPDATA")
        if appdata:
            return os.path.join(appdata, "PnP2PDF")
        return os.path.join(home, "AppData", "Roaming", "PnP2PDF")

    xdg = env.get("XDG_DATA_HOME")
    if xdg:
        return os.path.join(xdg, "pnp2pdf")
    return os.path.join(home, ".local", "share", "pnp2pdf")


def settings_path() -> str:
    """Absolute path to the settings file, creating its directory if needed."""
    directory = data_dir()
    os.makedirs(directory, exist_ok=True)
    return os.path.join(directory, "presets.json")
