from pathlib import Path
import sys
import tempfile

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
import scripts  # noqa: F401 - initializes categorized legacy module paths


def pytest_configure(config):
    # Use a unique owned directory instead of a shared pytest-current link
    # potentially created by another Windows user. Preserve it for diagnosis.
    if not config.option.basetemp:
        config.option.basetemp = tempfile.mkdtemp(prefix="excel-factory-tests-")

