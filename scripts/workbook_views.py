"""Apply explicitly declared Excel view adaptations after visual overlays."""
from openpyxl.utils.cell import coordinate_to_tuple


def apply_native_charts(path, spec):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    excel = None
    book = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0)
        for name, screen in spec.get("screen_contracts", {}).items():
            if "chart" not in screen.get("components", []):
                continue
            definition = screen["chart"]
            sheet = book.Worksheets(name)
            anchor = sheet.Range(definition.get("anchor", "H9"))
            chart = sheet.ChartObjects().Add(anchor.Left, anchor.Top, float(definition.get("width", 13)) * 28.3465, float(definition.get("height", 7)) * 28.3465).Chart
            types = {"bar": 57, "column": 51, "line": 4, "pie": 5, "doughnut": -4120}
            chart.ChartType = types[definition.get("type", "column")]
            values = definition["values"]
            first_row, last_row = values["min_row"], values["max_row"]
            for col in range(values["min_col"], values.get("max_col", values["min_col"]) + 1):
                series = chart.SeriesCollection().NewSeries()
                header = bool(definition.get("titles_from_data", True))
                series.Values = sheet.Range(sheet.Cells(first_row + int(header), col), sheet.Cells(last_row, col))
                if header:
                    series.Name = str(sheet.Cells(first_row, col).Value)
                categories = definition.get("categories")
                if categories:
                    series.XValues = sheet.Range(sheet.Cells(categories["min_row"], categories["min_col"]), sheet.Cells(categories["max_row"], categories.get("max_col", categories["min_col"])))
            chart.HasTitle = True
            chart.ChartTitle.Text = definition.get("title", "Overview")
            chart.HasLegend = definition.get("legend", True)
        book.Save()
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def apply_workbook_views(path, spec):
    panes = spec.get("freeze_panes", {})
    if not panes:
        return
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    excel = None
    book = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        book = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0)
        for name, anchor in panes.items():
            row, col = coordinate_to_tuple(anchor)
            book.Worksheets(name).Activate()
            window = excel.ActiveWindow
            window.FreezePanes = False
            window.SplitRow = row - 1
            window.SplitColumn = col - 1
            window.FreezePanes = True
        book.Worksheets(1).Activate()
        book.Save()
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
