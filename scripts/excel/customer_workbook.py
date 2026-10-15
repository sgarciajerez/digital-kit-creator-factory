"""Prepare a clean customer copy using Excel to preserve native objects."""
from __future__ import annotations

import shutil
from pathlib import Path

from model_quality import run_functional_contract


def create_clean_workbook(source: Path, destination: Path, contract: dict) -> dict:
    import pythoncom
    import win32com.client
    shutil.copy2(source, destination)
    pythoncom.CoInitialize()
    excel = None
    book = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(destination.resolve()), UpdateLinks=0)
        for table in contract["tables"]:
            if table["sheet"] in {"Settings", "Dashboard"}:
                continue
            sheet = book.Worksheets(table["sheet"])
            for native in sheet.ListObjects:
                for cell in native.DataBodyRange.Cells:
                    if not cell.HasFormula:
                        cell.ClearContents()
        settings = book.Worksheets("Settings")
        settings.Range("C8").Value = "My project"
        settings.Range("C9:C10").Value = 0
        settings.Range("C11:C12").ClearContents()
        book.Application.CalculateFullRebuild()
        book.Save()
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
    return run_functional_contract(destination, {"cases": [{
        "name": "Clean workbook starts at zero", "inputs": [],
        "assertions": [{"sheet": "Dashboard", "cell": address, "operator": "equals", "value": 0} for address in ("B7", "E7", "K7")],
    }]})
