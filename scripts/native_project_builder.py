from __future__ import annotations

"""Native Excel project renderer; kept separate from orchestration and legacy builders."""
import re
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent


def master_workbook_builder(spec: dict, path: Path) -> None:
    """Build a product by cloning and adapting the native master surface.

    The previous builder copied the master and then recreated every screen
    from scratch. That made the workbook functionally valid but visually
    unrelated to ``master_workbook.xlsx``. This builder keeps the mapped
    master sheets, merged cells, widths, heights, charts and navigation
    shapes. Only cell contents, ListObjects, formulas and validations inside
    the declared slots are replaced.
    """
    import pythoncom
    import win32com.client as win32
    from openpyxl.utils.datetime import to_excel

    def rgb(value: str) -> int:
        raw = value.replace("#", "")
        return int(raw[4:6] + raw[2:4] + raw[0:2], 16)

    def col_letter(number: int) -> str:
        return get_column_letter(number)

    source = ROOT / spec.get("master_workbook", "templates/excel-first-base/master_workbook.xlsx")
    product_to_source = spec.get("native_sheet_map") or {}
    if list(product_to_source) != list(spec.get("tabs") or []):
        raise ValueError("El builder master requiere native_sheet_map completo y ordenado")
    mapping = {source_name: product_name for product_name, source_name in product_to_source.items()}
    master_sheet_names = [
        "Inicio", "Resumen", "Presupuesto", "Gastos reales", "Movimientos",
        "Flujo de caja", "Metas y fondo", "Deudas", "Listas",
    ]
    deleted_sources = set(master_sheet_names) - set(mapping)
    palette = spec["palette"]
    ink = rgb(palette["ink"])
    input_fill = rgb(palette.get("input_fill", "FFF9E9"))
    formula_fill = rgb(palette.get("formula_fill", "F4F8F6"))
    currency_format = '#,##0.00 "EUR"'
    # Excel's desktop format parser is locale-aware in this COM path. Using
    # the Spanish four-digit year token avoids writing the literal text
    # ``yyyy`` into the visible cell while keeping the exported value a real
    # date.
    date_format = "dd/mm/aaaa"
    mock_profile = spec.get("mock_profile") or {}
    mock_areas = list(mock_profile.get("areas") or ["Kitchen units", "Flooring", "Lighting", "Plumbing", "Paint", "Permits", "Contingency"])
    mock_suppliers = list(mock_profile.get("suppliers") or ["Oak & Co", "Stone House", "Bright Home", "Flow Trade", "Colour Room", "City office", "Project reserve"])
    mock_items = list(mock_profile.get("items") or ["Primary package", "Materials", "Production", "Tools", "Services", "Travel", "Reserve"])
    mock_locations = list(mock_profile.get("locations") or ["Remote", "Studio", "Warehouse", "Online", "Project", "Office", "Project"])
    mock_phases = list(mock_profile.get("phases") or ["Planning", "Production", "Review", "Delivery", "Operations", "Finance", "Close"])
    mock_tasks = list(mock_profile.get("tasks") or ["Confirm scope", "Book supplier", "Review first draft", "Approve delivery", "Check budget", "Complete final review", "Archive records", "Confirm next action", "Review progress", "Close open items", "Prepare handover", "Finish project"])
    mock_quote_notes = list(mock_profile.get("quote_notes") or ["Scope confirmed", "Includes delivery", "Review lead time", "Alternative option", "Includes support", "Deposit required", "Reserve option"])
    path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, path)

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False, AddToMru=False)

        master_geometry = {}
        for source_name in mapping:
            source_sheet = workbook.Worksheets(source_name)
            master_geometry[source_name] = {
                "columns": {col: source_sheet.Columns(col).ColumnWidth for col in range(1, 18)},
                "rows": {row: source_sheet.Rows(row).RowHeight for row in range(1, 40)},
                "charts": [
                    {
                        "left": source_sheet.ChartObjects(index).Left,
                        "top": source_sheet.ChartObjects(index).Top,
                        "width": source_sheet.ChartObjects(index).Width,
                        "height": source_sheet.ChartObjects(index).Height,
                    }
                    for index in range(1, source_sheet.ChartObjects().Count + 1)
                ],
            }

        for sheet_name in deleted_sources:
            workbook.Worksheets(sheet_name).Delete()
        for source_name, product_name in mapping.items():
            workbook.Worksheets(source_name).Name = product_name

        def prepare_sheet(sheet, title: str, subtitle: str) -> None:
            # Delete only model objects and cell contents. Do not unmerge,
            # clear formatting, delete shapes or reset dimensions inherited
            # from the master workbook.
            for index in range(sheet.ListObjects.Count, 0, -1):
                sheet.ListObjects(index).Delete()
            try:
                sheet.UsedRange.Validation.Delete()
            except Exception:
                pass
            sheet.UsedRange.ClearContents()
            sheet.Range("B2").Value = f"RENOVATION PLANNER · {sheet.Name.upper()}"
            sheet.Range("B3").Value = title
            sheet.Range("B4").Value = subtitle
            sheet.Range("B2").Font.Color = rgb(palette["accent"])
            sheet.Range("B3").Font.Color = ink
            sheet.Range("B4").Font.Color = rgb(palette.get("muted", "6D7D89"))
            sheet.Tab.Color = rgb(palette["accent"])
            sheet.Application.ActiveWindow.DisplayGridlines = False
            sheet.Application.ActiveWindow.FreezePanes = False

        def apply_inherited_palette(sheet) -> None:
            """Recolour the copied master surface and remove its old logo.

            Copying the master preserves geometry, but its original teal
            colours and family-budget mark must never leak into a new product.
            The product palette is therefore applied to inherited fills,
            fonts and navigation shapes before the product data is written.
            """
            palette_map = {
                rgb("0F766E"): rgb(palette["accent"]),
                rgb("164E63"): rgb(palette["accent"]),
                rgb("477D8E"): rgb(palette["secondary"]),
                rgb("20343D"): ink,
                rgb("60767D"): rgb(palette.get("muted", "66757B")),
                rgb("F4F8F7"): rgb(palette["background"]),
                rgb("DDF3EF"): rgb(palette.get("input_fill", "E6F0EE")),
                rgb("E7F2EF"): rgb(palette.get("input_fill", "E6F0EE")),
                rgb("E8F0F5"): rgb(palette.get("formula_fill", "EFF3F2")),
                rgb("EDF4F5"): rgb(palette.get("formula_fill", "EFF3F2")),
                rgb("C6D4D8"): rgb(palette.get("line", "D4DEDB")),
            }
            used = sheet.UsedRange
            first_row = int(used.Row)
            first_col = int(used.Column)
            for row in range(first_row, first_row + int(used.Rows.Count)):
                for col in range(first_col, first_col + int(used.Columns.Count)):
                    cell = sheet.Cells(row, col)
                    try:
                        old_fill = int(cell.Interior.Color)
                        if old_fill in palette_map:
                            cell.Interior.Color = palette_map[old_fill]
                    except Exception:
                        pass
                    try:
                        old_font = int(cell.Font.Color)
                        if old_font in palette_map:
                            cell.Font.Color = palette_map[old_font]
                    except Exception:
                        pass
            for index in range(sheet.Shapes.Count, 0, -1):
                shape = sheet.Shapes(index)
                name = str(shape.Name)
                if name == "FamilyBudgetIcon":
                    shape.Delete()
                    continue
                if name.startswith("Nav_"):
                    try:
                        shape.Fill.ForeColor.RGB = rgb(palette["accent"])
                        shape.Line.ForeColor.RGB = rgb(palette["accent"])
                        shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
                    except Exception:
                        pass

        titles = {
            tab: (
                spec["title"] if tab == "Start here" else f"{spec['title']} · {tab}",
                spec.get("subtitle", "Use the inherited master layout to keep this workbook connected."),
            )
            for tab in spec["tabs"]
        }
        for tab in spec["tabs"]:
            sheet = workbook.Worksheets(tab)
            apply_inherited_palette(sheet)
            prepare_sheet(sheet, *titles[tab])

        def write_table(sheet, start_row: int, headers: list[str], rows: list[list[object]], table_name: str, roles: list[str], formats: list[str | None]) -> None:
            start_col = 2  # The master table frame starts in column B.
            last_col = start_col + len(headers) - 1
            last_letter = col_letter(last_col)
            for offset, header in enumerate(headers):
                cell = sheet.Cells(start_row, start_col + offset)
                cell.Value = header
            for row_offset, values in enumerate(rows, 1):
                row = start_row + row_offset
                for offset, value in enumerate(values):
                    cell = sheet.Cells(row, start_col + offset)
                    if isinstance(value, str) and value.startswith("="):
                        cell.Formula = value
                    elif isinstance(value, (date, datetime)):
                        # Keep the Excel date serial deterministic.  Passing a
                        # Python datetime through COM applies a timezone
                        # conversion on some machines and shifts the day.
                        cell.Value = float(to_excel(value))
                    else:
                        cell.Value = value
                    cell.Font.Color = ink
                    cell.Interior.Color = input_fill if roles[offset] == "input" else formula_fill
                    if formats[offset]:
                        if "d" in formats[offset].casefold():
                            cell.NumberFormatLocal = date_format
                        else:
                            cell.NumberFormat = formats[offset]
                    if formats[offset] and "d" in formats[offset].casefold():
                        cell.HorizontalAlignment = -4152  # xlRight
            end_row = start_row + len(rows)
            table = sheet.ListObjects.Add(1, sheet.Range(f"B{start_row}:{last_letter}{end_row}"), None, 1)
            table.Name = table_name
            table.TableStyle = "TableStyleMedium2"
            sheet.Range(f"B{start_row}:{last_letter}{start_row}").Interior.Color = rgb(palette["accent"])
            sheet.Range(f"B{start_row}:{last_letter}{start_row}").Font.Color = 0xFFFFFF
            sheet.Range(f"B{start_row}:{last_letter}{start_row}").HorizontalAlignment = -4108
            # The master workbook uses one deterministic table-alignment
            # contract. Headers are centered; the first field is a label and
            # stays left-aligned; every other field is right-aligned, even
            # when its value is text. This keeps dates, amounts, statuses and
            # notes visually anchored to the same inherited grid.
            sheet.Range(f"B{start_row}:{last_letter}{start_row}").HorizontalAlignment = -4108  # xlCenter
            if rows:
                sheet.Range(f"B{start_row + 1}:{last_letter}{end_row}").HorizontalAlignment = -4152  # xlRight
                sheet.Range(f"B{start_row + 1}:B{end_row}").HorizontalAlignment = -4131  # xlLeft

        def write_total(sheet, row: int, last_col: int, label: str, formulas: dict[int, str], formats: dict[int, str]) -> None:
            sheet.Cells(row, 2).Value = label
            sheet.Cells(row, 2).Font.Bold = True
            last_letter = col_letter(last_col)
            sheet.Range(f"B{row}:{last_letter}{row}").HorizontalAlignment = -4152  # xlRight
            sheet.Cells(row, 2).HorizontalAlignment = -4131  # xlLeft
            for col, formula in formulas.items():
                cell = sheet.Cells(row, col)
                cell.Formula = formula
                cell.Font.Bold = True
                cell.Interior.Color = formula_fill
                cell.HorizontalAlignment = -4152  # xlRight
                if col in formats:
                    cell.NumberFormat = formats[col]
            sheet.Range(f"B{row}:{last_letter}{row}").VerticalAlignment = -4108  # xlCenter

        def format_table_surface(sheet, table_name: str, header_row: int, last_row: int) -> None:
            """Fit table examples to the inherited master slots.

            Compact fields stay on one line and use ShrinkToFit. Only fields
            whose meaning is genuinely narrative are allowed to wrap.
            """
            table = sheet.ListObjects(table_name)
            header = table.HeaderRowRange
            header.WrapText = True
            header.VerticalAlignment = -4108  # xlCenter
            header.HorizontalAlignment = -4108  # xlCenter
            sheet.Rows(header_row).RowHeight = 30
            table.DataBodyRange.VerticalAlignment = -4108
            # Notes is still a compact table field: keep it on one line like
            # the other example values. Only genuinely narrative fields may
            # wrap inside their inherited slot.
            long_text_headers = {"description", "message", "body", "why it matters"}
            for column in range(1, table.ListColumns.Count + 1):
                header_text = str(table.HeaderRowRange.Cells(1, column).Value or "").strip().casefold()
                allow_wrap = header_text in long_text_headers
                body = table.ListColumns(column).DataBodyRange
                body.WrapText = allow_wrap
                body.ShrinkToFit = not allow_wrap
                body.VerticalAlignment = -4108  # xlCenter

        def populate_settings_reference_panels(sheet) -> None:
            """Populate every visible Settings panel or remove its inherited fill."""
            header_fill = rgb("0F766E")
            panel_fill = rgb("DDF3EF")
            panels = {
                "E": ("WORKFLOW", [
                    "1. Set budget cap",
                    "2. Set contingency",
                    "3. Confirm dates",
                    "4. Add purchases",
                    "5. Compare quotes",
                    "6. Review dashboard",
                ]),
                "G": ("AREAS", mock_areas),
                "I": ("STATUS", ["Planned", "Ordered", "Purchased", "Open", "In progress"]),
                "K": ("PAID?", ["Yes", "No"]),
                "M": ("CURRENCY", ["EUR", "USD", "GBP"]),
            }
            for column, (heading, values) in panels.items():
                sheet.Range(f"{column}6").Value = heading
                sheet.Range(f"{column}6").Interior.Pattern = 1
                sheet.Range(f"{column}6").Interior.Color = header_fill
                sheet.Range(f"{column}6").Font.Color = 0xFFFFFF
                sheet.Range(f"{column}6").Font.Bold = True
                sheet.Range(f"{column}7:{column}18").ClearContents()
                sheet.Range(f"{column}7:{column}18").Interior.Pattern = -4142
                sheet.Range(f"{column}7:{column}18").Borders.LineStyle = 0
                for index, value in enumerate(values, 7):
                    cell = sheet.Range(f"{column}{index}")
                    cell.Value = value
                    cell.Interior.Pattern = 1
                    cell.Interior.Color = panel_fill
                    cell.Font.Color = ink
                    cell.WrapText = True
                    cell.VerticalAlignment = -4108
            note = sheet.Range("E21:M23")
            note.Value = "Use the yellow inputs on the left; the reference panels keep the renovation workflow visible."
            note.Interior.Pattern = 1
            note.Interior.Color = formula_fill
            note.Font.Color = ink
            note.WrapText = True
            note.VerticalAlignment = -4108

        def remove_chart_gridlines(sheet) -> None:
            """Remove gridlines and dark plot borders from every chart on a sheet."""
            for index in range(1, sheet.ChartObjects().Count + 1):
                chart = sheet.ChartObjects(index).Chart
                for axis_type in (1, 2):
                    try:
                        axis = chart.Axes(axis_type)
                        axis.HasMajorGridlines = False
                        axis.HasMinorGridlines = False
                        axis.MajorGridlines.Format.Line.Visible = False
                        axis.MinorGridlines.Format.Line.Visible = False
                    except Exception:
                        pass
                for chart_part in (chart.ChartArea, chart.PlotArea):
                    try:
                        chart_part.Format.Line.Visible = False
                    except Exception:
                        pass

        def configure_timeline_chart(sheet) -> None:
            """Make the inherited timeline chart reflect timeline decisions.

            The master chart plots monthly cash-flow series. Once that screen
            becomes Timeline, keeping those source ranges produces a visually
            flat or semantically unrelated chart. Use the derived Days left
            column against Task instead, with a native column chart.
            """
            if sheet.ChartObjects().Count < 1:
                return
            chart = sheet.ChartObjects(1).Chart
            chart.ChartType = 51  # xlColumnClustered
            try:
                while chart.SeriesCollection().Count > 0:
                    chart.SeriesCollection(1).Delete()
            except Exception:
                pass
            series = chart.SeriesCollection().NewSeries()
            series.Name = "Days left"
            series.Values = f"='{sheet.Name}'!$G$8:$G$19"
            series.XValues = f"='{sheet.Name}'!$B$8:$B$19"
            try:
                series.Format.Fill.ForeColor.RGB = rgb(palette["accent"])
                series.Format.Line.Visible = False
            except Exception:
                pass
            chart.HasTitle = True
            chart.ChartTitle.Text = "Launch tasks by days left"

        def validation(sheet, address: str, values: str) -> None:
            rule = sheet.Range(address).Validation
            rule.Delete()
            rule.Add(3, 1, 1, f'"{values}"')

        def named_validation(sheet, address: str, name: str) -> None:
            """Apply a locale-independent list validation using a defined name."""
            rule = sheet.Range(address).Validation
            rule.Delete()
            rule.Add(3, 1, 1, f"={name}")

        start = workbook.Worksheets("Start here")
        start.Range("B6").Value = "HOW THIS WORKBOOK WORKS"
        start.Range("G6").Value = "QUICK PROJECT VIEW"
        steps = [
            ("01", "Set the project assumptions", "Enter the budget cap, contingency and dates in Settings."),
            ("02", "Log decisions once", "Add purchases and supplier quotes in their dedicated tabs."),
            ("03", "Use the dashboard", "Check remaining budget and the next task before deciding."),
            ("04", "Keep the timeline current", "Update task status and due dates as the work moves."),
        ]
        for row, (number, heading, body) in enumerate(steps, 7):
            start.Cells(row, 2).Value = number
            start.Cells(row, 3).Value = heading
            start.Cells(row, 5).Value = body
        quick = [("Project", "=Settings!$C$8"), ("Budget cap", "=Settings!$C$9"), ("Contingency", "=Settings!$C$10"), ("Start date", "=Settings!$C$11"), ("Target date", "=Settings!$C$12")]
        for row, (label, formula) in enumerate(quick, 7):
            start.Cells(row, 7).Value = label
            start.Cells(row, 8).Formula = formula
        start.Range("H8").NumberFormat = currency_format
        start.Range("H9").NumberFormat = "0%"
        start.Range("H10:H11").NumberFormatLocal = date_format
        start.Range("B14").Value = "WORKBOOK MAP"
        map_rows = [("Dashboard", "Budget, actual spend and next decisions."), ("Budget", "Planned renovation areas and remaining amounts."), ("Purchases", "One operational log for materials and supplier costs."), ("Quotes", "Compare alternatives before committing."), ("Timeline", "Tasks, dates, statuses and days left."), ("Settings", "Project assumptions and dates."), ("Start here", "Setup guide and quick project view.")]
        for row, (label, body) in enumerate(map_rows, 15):
            start.Cells(row, 2).Value = label
            start.Cells(row, 3).Value = body
        start.Range("B25").Value = "Tip: duplicate the file before starting a new renovation. The master grid and formulas remain easy to review and adapt."

        settings = workbook.Worksheets("Settings")
        settings.Range("B6").Value = "PROJECT SETTINGS"
        settings_rows = [["Project name", mock_profile.get("project_name", "Oak Street kitchen refresh")], ["Budget cap", 24000], ["Contingency %", 0.10], ["Start date", date(2026, 10, 20)], ["Target date", date(2027, 2, 28)], ["Currency", "EUR"]]
        write_table(settings, 7, ["Setting", "Value"], settings_rows, "tblRenovationSettings", ["input", "input"], [None, None])
        settings.Range("C8:C13").HorizontalAlignment = -4152  # xlRight
        settings.Range("C8").WrapText = False
        settings.Range("C8").ShrinkToFit = True
        settings.Range("C8:C13").VerticalAlignment = -4108  # xlCenter
        settings.Range("C9").NumberFormat = currency_format
        settings.Range("C10").NumberFormat = "0%"
        settings.Range("C11:C12").NumberFormatLocal = date_format
        populate_settings_reference_panels(settings)

        budget = workbook.Worksheets("Budget")
        budget.Range("B6").Value = "RENOVATION BUDGET · INPUTS IN YELLOW"
        base_area_numbers = [9000, 5200, 3600, 1800, 1400, 1000, 2400]
        base_area_notes = ["Primary creator income or project allocation", "Production and delivery costs", "Editing and post-production", "Tools and subscriptions", "Workspace and equipment", "Travel and creator logistics", "Reserve for tax and surprises"]
        areas = [(area, base_area_numbers[index], base_area_notes[index], 4500 if index == 0 else 0, 2500 if index == 0 else 0, 1500 if index == 0 else 0, 300 if index == 0 else 0, 0, 200 if index == 0 else 0, 0, base_area_numbers[index]) for index, area in enumerate(mock_areas)]
        budget_rows = [[f"=Settings!G{row - 1}", planned, f'=SUMIF(Purchases!$C$8:$C$14,B{row},Purchases!$F$8:$F$14)', f"=C{row}-D{row}", f'=IF(D{row}=0,"Not started",IF(E{row}<0,"Over budget","On track"))', note, materials, labour, fixtures, finishes, permits, design, contingency, forecast] for row, (area, planned, note, materials, labour, fixtures, finishes, permits, design, contingency, forecast) in enumerate(areas, 8)]
        budget_headers = ["Area", "Planned", "Actual spend", "Remaining", "Status", "Notes", "Materials", "Labour", "Fixtures", "Finishes", "Permits", "Design", "Contingency", "Forecast"]
        budget_roles = ["input", "input", "derived", "derived", "derived", "input"] + ["derived"] * 8
        budget_formats = [None, currency_format, currency_format, currency_format, None, None] + [currency_format] * 8
        write_table(budget, 7, budget_headers, budget_rows, "tblRenovationBudget", budget_roles, budget_formats)
        write_total(budget, 15, 15, "TOTAL PROJECT", {3: "=SUM(C8:C14)", 4: "=SUM(D8:D14)", 5: "=SUM(E8:E14)", **{col: f"=SUM({col_letter(col)}8:{col_letter(col)}14)" for col in range(7, 16)}}, {col: currency_format for col in [3, 4, 5, *range(7, 16)]})
        budget.Range("F15").Value = "On track"
        budget.Range("G15").Value = "LINKED TO PURCHASES"
        budget.ListObjects("tblRenovationBudget").Resize(budget.Range("B7:O14"))
        budget.Range("B18").Value = "Planned, actual and remaining amounts stay connected to Purchases across the full renovation budget grid."

        purchases = workbook.Worksheets("Purchases")
        purchases.Range("B6").Value = "PURCHASE LOG · ENTER EACH COST ONCE"
        purchase_rows = [
            [date(2026, 10, 20), "Kitchen units", "Deposit", "Oak & Co", 0, "No", "Planned", "Joinery", "Q-001", "Kitchen", "Transfer", date(2026, 10, 20), "Pending", "Confirm fitting date"],
            [date(2026, 10, 28), "Flooring", "Porcelain tiles", "Stone House", 1800, "Yes", "Purchased", "Materials", "Q-002", "Kitchen", "Card", date(2026, 10, 28), "Saved", "Include delivery"],
            [date(2026, 11, 4), "Lighting", "Pendant lights", "Bright Home", 1200, "No", "Ordered", "Fixtures", "Q-003", "Dining", "Transfer", date(2026, 11, 4), "Pending", "Check lead time"],
            [date(2026, 11, 12), "Plumbing", "Tap and valves", "Flow Trade", 500, "Yes", "Purchased", "Plumbing", "Q-004", "Kitchen", "Card", date(2026, 11, 12), "Saved", "Materials included"],
            [date(2026, 11, 18), "Paint", "Primer and paint", "Colour Room", 400, "No", "Planned", "Finishes", "Q-005", "Hall", "Transfer", date(2026, 11, 18), "Pending", "Confirm colour"],
            [date(2026, 11, 25), "Permits", "Permit fee", "City office", 0, "No", "Planned", "Permits", "Q-006", "Property", "Transfer", date(2026, 11, 25), "Pending", "Submit application"],
            [date(2026, 12, 2), "Contingency", "Reserved amount", "Project reserve", 0, "No", "Planned", "Reserve", "Q-007", "Project", "Transfer", date(2026, 12, 2), "Pending", "Keep untouched"],
        ]
        if mock_profile:
            purchase_rows = [
                [date(2026, 10, 20), mock_areas[0], mock_items[0], mock_suppliers[0], 0, "No", "Planned", mock_phases[0], "Q-001", mock_locations[0], "Transfer", date(2026, 10, 20), "Pending", mock_quote_notes[0]],
                [date(2026, 10, 28), mock_areas[1], mock_items[1], mock_suppliers[1], 1800, "Yes", "Purchased", mock_phases[1], "Q-002", mock_locations[1], "Card", date(2026, 10, 28), "Saved", mock_quote_notes[1]],
                [date(2026, 11, 4), mock_areas[2], mock_items[2], mock_suppliers[2], 1200, "No", "Ordered", mock_phases[2], "Q-003", mock_locations[2], "Transfer", date(2026, 11, 4), "Pending", mock_quote_notes[2]],
                [date(2026, 11, 12), mock_areas[3], mock_items[3], mock_suppliers[3], 500, "Yes", "Purchased", mock_phases[3], "Q-004", mock_locations[3], "Card", date(2026, 11, 12), "Saved", mock_quote_notes[3]],
                [date(2026, 11, 18), mock_areas[4], mock_items[4], mock_suppliers[4], 400, "No", "Planned", mock_phases[4], "Q-005", mock_locations[4], "Transfer", date(2026, 11, 18), "Pending", mock_quote_notes[4]],
                [date(2026, 11, 25), mock_areas[5], mock_items[5], mock_suppliers[5], 0, "No", "Planned", mock_phases[5], "Q-006", mock_locations[5], "Card", date(2026, 11, 25), "Pending", mock_quote_notes[5]],
                [date(2026, 12, 2), mock_areas[6], mock_items[6], mock_suppliers[6], 0, "No", "Planned", mock_phases[6], "Q-007", mock_locations[6], "Transfer", date(2026, 12, 2), "Pending", mock_quote_notes[6]],
            ]
        purchase_headers = ["Purchase date", "Area", "Item", "Supplier", "Amount", "Paid?", "Status", "Phase", "Quote ref", "Room", "Payment method", "Due date", "Receipt", "Notes"]
        purchase_formats = [date_format, None, None, None, currency_format, None, None, None, None, None, None, date_format, None, None]
        write_table(purchases, 7, purchase_headers, purchase_rows, "tblRenovationPurchases", ["input"] * 14, purchase_formats)
        write_total(purchases, 15, 15, "TOTAL PURCHASES", {6: "=SUM(F8:F14)"}, {6: currency_format})
        purchases.Range("H15").Value = "Review"
        purchases.Range("O15").Value = "One operational log"
        purchases.ListObjects("tblRenovationPurchases").Resize(purchases.Range("B7:O14"))
        purchases.Range("B18").Value = "Log each cost once here; Budget and Dashboard update automatically across the full purchase grid."
        # Enum lists are part of the normative model contract.  The builder
        # must render that contract instead of maintaining a second set of
        # hardcoded dropdown values that can silently drift.
        contract = spec.get("model_contract") or contract_for_product(int(spec["product_id"]))
        enum_contract = contract.get("enum_contract", [])
        enum_store_column = 24  # X; outside the visible master grid.
        settings = workbook.Worksheets("Settings")
        created_validation_names: set[str] = set()
        for enum_index, enum in enumerate(enum_contract):
            enum_sheet = workbook.Worksheets(str(enum["sheet"]))
            enum_values = [str(value) for value in enum.get("values", [])]
            if not enum_values:
                raise ValueError(f"Enum vacío: {enum.get('sheet')}!{enum.get('range')}")
            if str(enum.get("column")) == "Area":
                # Settings!G7:G... is the single visible source of truth for
                # areas. Every input table that uses Area points to this same
                # named range, so Budget, Purchases, Quotes and Timeline
                # cannot drift apart.
                storage_column = "G"
                start_row = 7
                end_row = start_row + len(enum_values) - 1
                validation_name = f"dv_{int(spec['product_id'])}_areas"
                settings.Range(f"{storage_column}{start_row}:{storage_column}{end_row}").Value = [[value] for value in enum_values]
                ref = f"='{settings.Name}'!${storage_column}${start_row}:${storage_column}${end_row}"
            else:
                storage_column = col_letter(enum_store_column + enum_index)
                start_row = 2
                end_row = start_row + len(enum_values) - 1
                validation_name = f"dv_{int(spec['product_id'])}_{enum_index + 1}"
                settings.Range(f"{storage_column}1").Value = validation_name
                settings.Range(f"{storage_column}{start_row}:{storage_column}{end_row}").Value = [[value] for value in enum_values]
                settings.Columns(storage_column).Hidden = True
                ref = f"='{settings.Name}'!${storage_column}${start_row}:${storage_column}${end_row}"
            if validation_name not in created_validation_names:
                try:
                    workbook.Names(validation_name).Delete()
                except Exception:
                    pass
                workbook.Names.Add(Name=validation_name, RefersTo=ref)
                created_validation_names.add(validation_name)
            named_validation(enum_sheet, str(enum["range"]), validation_name)

        quotes = workbook.Worksheets("Quotes")
        quotes.Range("B6").Value = "QUOTE COMPARISON"
        quote_rows = [[supplier, area, quote_date, amount, selected, notes, lead_time, decision] for supplier, area, quote_date, amount, selected, notes, lead_time, decision in [
            ["Oak & Co", "Kitchen units", date(2026, 10, 14), 9400, "Yes", "Includes fitting", "3 weeks", "Preferred"],
            ["Northline Joinery", "Kitchen units", date(2026, 10, 16), 10200, "No", "Longer lead time", "5 weeks", "Backup"],
            ["Stone House", "Flooring", date(2026, 10, 18), 5200, "Yes", "Includes installation", "2 weeks", "Preferred"],
            ["Light Lab", "Lighting", date(2026, 10, 21), 3600, "No", "Awaiting revised quote", "4 weeks", "Review"],
            ["Bright Home", "Lighting", date(2026, 10, 23), 3450, "Yes", "Trade discount", "2 weeks", "Preferred"],
            ["Flow Trade", "Plumbing", date(2026, 10, 25), 1800, "Yes", "Materials included", "1 week", "Preferred"],
            ["Pipe & Brass", "Plumbing", date(2026, 10, 27), 2150, "No", "Higher labour allowance", "2 weeks", "Backup"],
            ["Paint House", "Paint", date(2026, 11, 1), 1280, "Yes", "Includes primer", "1 week", "Preferred"],
            ["Colour Works", "Paint", date(2026, 11, 2), 1450, "No", "Premium finish option", "2 weeks", "Review"],
            ["Permit Desk", "Permits", date(2026, 11, 3), 680, "Yes", "Application support", "1 week", "Preferred"],
            ["City Permits", "Permits", date(2026, 11, 4), 520, "No", "Self-service guidance", "3 weeks", "Backup"],
            ["Joinery Studio", "Kitchen units", date(2026, 11, 6), 9800, "No", "Alternative door profile", "4 weeks", "Review"],
            ["Tile & Co", "Flooring", date(2026, 11, 8), 4980, "No", "Excludes delivery", "3 weeks", "Backup"],
            ["Warm Floor", "Flooring", date(2026, 11, 9), 5600, "Yes", "Includes underlay", "2 weeks", "Preferred"],
            ["Switchboard", "Lighting", date(2026, 11, 10), 2950, "No", "Fixtures excluded", "3 weeks", "Review"],
            ["Oakline Surfaces", "Kitchen units", date(2026, 11, 12), 8750, "No", "Laminate alternative", "3 weeks", "Review"],
            ["Solid Oak Co", "Kitchen units", date(2026, 11, 13), 11600, "No", "Solid timber option", "6 weeks", "Backup"],
            ["Step Right", "Flooring", date(2026, 11, 15), 4720, "No", "Engineered boards", "2 weeks", "Review"],
            ["Lumen House", "Lighting", date(2026, 11, 17), 3180, "No", "Includes dimmers", "2 weeks", "Backup"],
            ["Flowline Services", "Plumbing", date(2026, 11, 19), 1980, "No", "Weekend availability", "2 weeks", "Review"],
            ["Fresh Coat", "Paint", date(2026, 11, 21), 1190, "No", "Trade account pricing", "1 week", "Backup"],
        ]]
        if mock_profile:
            quote_rows = [
                [mock_suppliers[index % len(mock_suppliers)], mock_areas[index % len(mock_areas)], date(2026, 10, 14) + timedelta(days=index), 1800 + index * 275, "Yes" if index % 3 == 0 else "No", mock_quote_notes[index % len(mock_quote_notes)], f"{index % 5 + 1} weeks", "Preferred" if index % 3 == 0 else "Review"]
                for index in range(21)
            ]
        write_table(quotes, 7, ["Supplier", "Area", "Quote date", "Amount", "Selected", "Notes", "Lead time", "Decision"], quote_rows, "tblRenovationQuotes", ["input"] * 8, [None, None, date_format, currency_format, None, None, None, None])
        quotes.ListObjects("tblRenovationQuotes").Resize(quotes.Range("B7:I27"))
        for address, text_value in (
            ("K7:M10", "Compare supplier, area and amount before committing."),
            ("K12:M15", "Mark Selected = Yes only for the quote you will carry into Purchases."),
        ):
            card = quotes.Range(address)
            card.Value = text_value
            card.Interior.Pattern = 1
            card.Interior.Color = formula_fill
            card.Font.Color = ink
            card.WrapText = True
            card.VerticalAlignment = -4108

        timeline = workbook.Worksheets("Timeline")
        timeline.Range("B6").Value = "PROJECT TIMELINE"
        tasks = [("Confirm kitchen layout", "Kitchen units", date(2026, 10, 20), date(2026, 11, 3), "Open"), ("Order flooring", "Flooring", date(2026, 10, 28), date(2026, 11, 10), "In progress"), ("Book electrician", "Lighting", date(2026, 11, 4), date(2026, 11, 15), "Open"), ("Complete plumbing rough-in", "Plumbing", date(2026, 11, 12), date(2026, 11, 22), "Open"), ("Confirm permit", "Permits", date(2026, 11, 15), date(2026, 11, 29), "Open"), ("Prepare paintwork", "Paint", date(2027, 1, 5), date(2027, 1, 17), "Open"), ("Paint and finish", "Paint", date(2027, 1, 18), date(2027, 1, 31), "Open"), ("Install kitchen", "Kitchen units", date(2027, 2, 1), date(2027, 2, 14), "Open"), ("Install flooring", "Flooring", date(2027, 2, 2), date(2027, 2, 16), "Open"), ("Final snag list", "Contingency", date(2027, 2, 17), date(2027, 2, 22), "Open"), ("Deep clean", "Contingency", date(2027, 2, 23), date(2027, 2, 25), "Open"), ("Handover", "Contingency", date(2027, 2, 26), date(2027, 2, 28), "Open")]
        timeline_rows = [[task, area, start_date, due_date, status, f"=MAX(E{row}-TODAY(),0)", owner, phase, note] for row, (task, area, start_date, due_date, status, owner, phase, note) in enumerate([
            ("Confirm kitchen layout", "Kitchen units", date(2026, 10, 20), date(2026, 11, 3), "Open", "Santi", "Design", "Approve final layout"),
            ("Order flooring", "Flooring", date(2026, 10, 28), date(2026, 11, 10), "In progress", "Santi", "Materials", "Confirm delivery"),
            ("Book electrician", "Lighting", date(2026, 11, 4), date(2026, 11, 15), "Open", "Contractor", "Fixtures", "Share lighting plan"),
            ("Complete plumbing rough-in", "Plumbing", date(2026, 11, 12), date(2026, 11, 22), "Open", "Contractor", "Plumbing", "Check access"),
            ("Confirm permit", "Permits", date(2026, 11, 15), date(2026, 11, 29), "Open", "Santi", "Permits", "Submit documents"),
            ("Prepare paintwork", "Paint", date(2027, 1, 5), date(2027, 1, 17), "Open", "Contractor", "Finishes", "Protect surfaces"),
            ("Paint and finish", "Paint", date(2027, 1, 18), date(2027, 1, 31), "Open", "Contractor", "Finishes", "Final coat"),
            ("Install kitchen", "Kitchen units", date(2027, 2, 1), date(2027, 2, 14), "Open", "Joiner", "Joinery", "Coordinate delivery"),
            ("Install flooring", "Flooring", date(2027, 2, 2), date(2027, 2, 16), "Open", "Contractor", "Materials", "Allow curing time"),
            ("Final snag list", "Contingency", date(2027, 2, 17), date(2027, 2, 22), "Open", "Santi", "Review", "Walk every room"),
            ("Deep clean", "Contingency", date(2027, 2, 23), date(2027, 2, 25), "Open", "Cleaner", "Handover", "Book after works"),
            ("Handover", "Contingency", date(2027, 2, 26), date(2027, 2, 28), "Open", "Santi", "Handover", "Archive receipts"),
        ], 8)]
        if mock_profile:
            creator_tasks = [
                (mock_tasks[index], mock_areas[index % len(mock_areas)], start_date, due_date, "In progress" if index == 1 else "Open", "Project lead", mock_phases[index % len(mock_phases)], mock_quote_notes[index % len(mock_quote_notes)])
                for index, (start_date, due_date) in enumerate([
                    (date(2026, 10, 20), date(2026, 11, 3)), (date(2026, 10, 28), date(2026, 11, 10)),
                    (date(2026, 11, 4), date(2026, 11, 15)), (date(2026, 11, 12), date(2026, 11, 22)),
                    (date(2026, 11, 15), date(2026, 11, 29)), (date(2027, 1, 5), date(2027, 1, 17)),
                    (date(2027, 1, 18), date(2027, 1, 31)), (date(2027, 2, 1), date(2027, 2, 14)),
                    (date(2027, 2, 2), date(2027, 2, 16)), (date(2027, 2, 17), date(2027, 2, 22)),
                    (date(2027, 2, 23), date(2027, 2, 25)), (date(2027, 2, 26), date(2027, 2, 28)),
                ])
            ]
            timeline_rows = [[task, area, start_date, due_date, status, f"=MAX(E{row}-TODAY(),0)", owner, phase, note] for row, (task, area, start_date, due_date, status, owner, phase, note) in enumerate(creator_tasks, 8)]
        write_table(timeline, 7, ["Task", "Area", "Start date", "Due date", "Status", "Days left", "Owner", "Phase", "Notes"], timeline_rows, "tblRenovationTimeline", ["input", "input", "input", "input", "input", "derived", "input", "input", "input"], [None, None, date_format, date_format, None, "0", None, None, None])
        timeline.Range("B20").Value = "TOTAL TASKS"
        timeline.Range("C20").Formula = "=COUNTA(B8:B19)"
        # Totals must point to the semantic columns: Status is F and Days
        # left is G in the inherited B:J timeline frame. The previous
        # implementation averaged the Owner text instead of Days left.
        timeline.Range("F20").Formula = '=COUNTIF(F8:F19,"Open")'
        timeline.Range("G20").Formula = "=AVERAGE(G8:G19)"
        timeline.Range("J20").Value = "Keep the next action visible"
        for address in ("B20", "J20"):
            timeline.Range(address).HorizontalAlignment = -4131  # xlLeft
        for address in ("C20", "F20", "G20"):
            timeline.Range(address).HorizontalAlignment = -4152  # xlRight
            timeline.Range(address).VerticalAlignment = -4108
        timeline.Range("B20:J20").VerticalAlignment = -4108
        timeline.Range("B20:J20").Font.Bold = True
        timeline.Range("B23").Value = "Keep due dates current: Days left updates automatically and highlights the next action across the full timeline grid."
        timeline.ListObjects("tblRenovationTimeline").Resize(timeline.Range("B7:J19"))

        dashboard = workbook.Worksheets("Dashboard")
        dashboard.Range("B5").Value = "Project"
        dashboard.Range("C5").Formula = "=Settings!C8"
        for label_cell, value_cell, label, formula in (("B6", "B7", "BUDGET CAP", "=Settings!C9"), ("E6", "E7", "ACTUAL SPEND", "=SUM(Purchases!F8:F14)"), ("H6", "H7", "CONTINGENCY", "=Settings!C9*Settings!C10"), ("K6", "K7", "AVAILABLE", "=B7-E7")):
            dashboard.Range(label_cell).Value = label
            dashboard.Range(value_cell).Formula = formula
            dashboard.Range(value_cell).NumberFormat = currency_format
        dashboard.Range("B10").Value = "BUDGET BY RENOVATION AREA"
        dashboard_rows = [[f"=Budget!B{row}", f"=Budget!C{row}", f"=Budget!D{row}", f"=Budget!E{row}", f"=Budget!F{row}"] for row in range(8, 15)]
        write_table(dashboard, 11, ["Area", "Planned", "Actual spend", "Remaining", "Status"], dashboard_rows, "tblRenovationDashboard", ["derived"] * 5, [None, currency_format, currency_format, currency_format, None])
        write_total(dashboard, 19, 6, "TOTAL PROJECT", {3: "=SUM(C12:C18)", 4: "=SUM(D12:D18)", 5: "=SUM(E12:E18)"}, {3: currency_format, 4: currency_format, 5: currency_format})
        dashboard.ListObjects("tblRenovationDashboard").Resize(dashboard.Range("B11:F18"))
        for address, fill in (
            ("B6:D6", rgb("DDF3EF")), ("B7:D8", rgb("DDF3EF")),
            ("E6:G6", rgb("EDF4F5")), ("E7:G8", rgb("EDF4F5")),
            ("H6:J6", rgb("DDF3EF")), ("H7:J8", rgb("DDF3EF")),
            ("K6:M6", rgb("E7F2EF")), ("K7:M8", rgb("E7F2EF")),
        ):
            card = dashboard.Range(address)
            card.Interior.Pattern = 1
            card.Interior.Color = fill
            card.WrapText = True
            card.VerticalAlignment = -4108
        dashboard.Range("B5:C5").Interior.Pattern = 1
        dashboard.Range("B5:C5").Interior.Color = formula_fill
        dashboard.Range("B22:F24").Value = "Positive remaining means the area is still within plan; negative means review the scope or quote."
        dashboard.Range("B22:F24").Interior.Pattern = 1
        dashboard.Range("B22:F24").Interior.Color = formula_fill
        dashboard.Range("B22:F24").Font.Color = ink
        dashboard.Range("B22:F24").WrapText = True
        dashboard.Range("B22:F24").VerticalAlignment = -4108
        remove_chart_gridlines(dashboard)
        # The master also contains a line chart on Timeline.  Its vertical
        # axis gridlines are the source of the dark horizontal bands visible
        # behind the plotted lines, so clean it with the same chart contract.
        configure_timeline_chart(timeline)
        remove_chart_gridlines(timeline)

        nav_labels = {source_name: product_name for product_name, source_name in product_to_source.items()}
        for sheet in workbook.Worksheets:
            for index in range(sheet.Shapes.Count, 0, -1):
                shape = sheet.Shapes(index)
                name = str(shape.Name)
                if not name.startswith("Nav_"):
                    continue
                source_name = name[4:].replace("_", " ")
                if source_name not in mapping:
                    shape.Delete()
                    continue
                shape.TextFrame2.TextRange.Text = nav_labels[source_name]
                target_name = mapping[source_name]
                for link_index in range(sheet.Hyperlinks.Count, 0, -1):
                    link = sheet.Hyperlinks(link_index)
                    try:
                        if link.Shape.Name == name:
                            link.Delete()
                    except Exception:
                        pass
                sheet.Hyperlinks.Add(shape, "", f"'{target_name}'!A1")
            sheet.Application.ActiveWindow.FreezePanes = False

        # Excel may widen a column when a new ListObject is created. Restore
        # the exact master geometry after all writes and table resizes.
        for source_name, product_name in mapping.items():
            sheet = workbook.Worksheets(product_name)
            geometry = master_geometry[source_name]
            for col, width in geometry["columns"].items():
                sheet.Columns(col).ColumnWidth = width
            for row, height in geometry["rows"].items():
                sheet.Rows(row).RowHeight = height

        # Keep the master dimensions exact. Readability is handled with
        # wrapped headers and content designed for the existing columns, not
        # by widening or cutting the master grid.
        table_layouts = [
            ("Settings", "tblRenovationSettings", 7, 13),
            ("Budget", "tblRenovationBudget", 7, 15),
            ("Purchases", "tblRenovationPurchases", 7, 15),
            ("Quotes", "tblRenovationQuotes", 7, 27),
            ("Timeline", "tblRenovationTimeline", 7, 20),
            ("Dashboard", "tblRenovationDashboard", 11, 19),
        ]
        for tab, table_name, header_row, last_row in table_layouts:
            sheet = workbook.Worksheets(tab)
            format_table_surface(sheet, table_name, header_row, last_row)

        # Keep charts in the master viewport even when fitting a table widens
        # columns to accommodate product-specific labels.
        for source_name, product_name in mapping.items():
            sheet = workbook.Worksheets(product_name)
            for index, geometry in enumerate(master_geometry[source_name]["charts"], 1):
                if index > sheet.ChartObjects().Count:
                    break
                chart_object = sheet.ChartObjects(index)
                chart_object.Left = geometry["left"]
                chart_object.Top = geometry["top"]
                chart_object.Width = geometry["width"]
                chart_object.Height = geometry["height"]

        # The product has seven tabs, while the master has nine. Reflow the
        # surviving native buttons so the final Settings button does not keep
        # the deleted Deudas/Metas gap from the source workbook.
        nav_left = 15.8
        nav_step = 72.0
        for sheet in workbook.Worksheets:
            for index, source_name in enumerate(mapping):
                shape_name = f"Nav_{source_name.replace(' ', '_')}"
                try:
                    shape = sheet.Shapes(shape_name)
                except Exception:
                    continue
                shape.Left = nav_left + index * nav_step
                shape.Top = 2.0
                shape.Width = 68.0
                shape.Height = 21.0

        from project_model import apply_project_behaviour
        apply_project_behaviour(workbook, spec)
        workbook.Application.CalculateFullRebuild()
        workbook.Save()
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

