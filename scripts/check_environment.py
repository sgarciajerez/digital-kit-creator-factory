"""Read-only checks for a reproducible Windows desktop Excel setup."""
from __future__ import annotations

import argparse
import importlib.util
import json
import sys


def check_environment(check_excel: bool = False) -> dict:
    modules = ["openpyxl", "pydantic", "dotenv", "yaml", "streamlit", "PIL", "imageio", "imageio_ffmpeg", "mss", "numpy", "pyautogui", "pyperclip", "pythoncom", "win32com.client"]
    missing = []
    for module in modules:
        try:
            found = importlib.util.find_spec(module) is not None
        except ModuleNotFoundError:
            found = False
        if not found:
            missing.append(module)
    errors = [f"Missing module: {module}" for module in missing]
    if sys.platform != "win32":
        errors.append("Native generation requires Windows and desktop Excel")
    if check_excel and not missing and sys.platform == "win32":
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            version = str(excel.Version)
        except Exception as exc:
            errors.append(f"Desktop Excel is unavailable: {exc}")
        finally:
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()
    else:
        version = None
    return {"passed": not errors, "python": sys.version.split()[0], "excel_version": locals().get("version"), "errors": errors}


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", action="store_true")
    result = check_environment(parser.parse_args().excel)
    print(json.dumps(result, indent=2))
    raise SystemExit(0 if result["passed"] else 1)
