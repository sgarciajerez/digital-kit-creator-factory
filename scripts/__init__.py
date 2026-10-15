"""Organized script packages with temporary support for legacy flat imports."""

from pathlib import Path
import sys

_SCRIPT_ROOT = Path(__file__).resolve().parent
for _category in _SCRIPT_ROOT.iterdir():
    if _category.is_dir() and not _category.name.startswith((".", "__")):
        _category_path = str(_category)
        if _category_path not in sys.path:
            sys.path.insert(0, _category_path)
