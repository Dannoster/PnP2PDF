import os


curr_path = os.path.dirname(os.path.realpath(__file__))
ICC_DIR = os.path.join(curr_path, "color_profiles")


def icc_profile(name: str) -> str:
    """Absolute path to a bundled ICC profile (works regardless of CWD)."""
    return os.path.join(ICC_DIR, name)
