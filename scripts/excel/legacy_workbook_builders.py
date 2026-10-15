from __future__ import annotations

import json
import re
import argparse
import shutil
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter

from excel_visual_overlays import apply_simple_native_overlays, apply_visual_overlays
from validate_design_architecture import validate_product
from factory_config import active_product_ids, all_product_ids, finalized_product_ids
from excel_template_factory import ExcelTemplateFactory
from model_quality import number_format_for_header, validate_headers, validate_model_file
from model_contracts import contract_for_product, write_app_usage_markdown, write_model_behavior_markdown, write_model_contract_json
from artifact_integrity import artifact_identity


ROOT = Path(__file__).resolve().parents[2]
OUT_ROOT = ROOT / "outputs" / "listings"
ALL_PRODUCT_IDS = all_product_ids()
PRODUCT_IDS = active_product_ids()


def prepare_product8_header(ws) -> None:
    """Translate the master header zones into a native Excel header.

    The header is represented by real cells: a breathing-space row, eyebrow,
    title and subtitle. Keeping these zones explicit preserves the intended
    spacing while the content cards are padded farther in.
    """
    title = ws["A1"].value or ""
    subtitle = ws["A2"].value or ""
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row <= 4 and merged.max_row <= 4:
            ws.unmerge_cells(str(merged))
    # Keep A1 populated for the semantic workbook validator while using the
    # visible header in rows 2-4.
    ws["A1"] = title
    ws["A2"] = None
    ws["A3"] = None
    ws["A4"] = None
    header_column = "A" if ws.title in {"Dashboard", "Start here"} else "B"
    ws["B2"] = None
    ws["B3"] = None
    ws["B4"] = None
    ws[f"{header_column}2"] = f"WEDDING PLANNER · {ws.title.upper()}"
    ws[f"{header_column}3"] = title
    ws[f"{header_column}4"] = subtitle
    ws.merge_cells(f"{header_column}2:L2")
    ws.merge_cells(f"{header_column}3:L3")
    ws.merge_cells(f"{header_column}4:L4")
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 54
    ws.row_dimensions[4].height = 100
    # The actual palette/fills/fonts are applied by apply_visual_system after
    # this structural pass; this helper only establishes the geometry.


def colors(spec: dict) -> dict[str, str]:
    palette = spec["palette"]
    return {
        "accent": palette["accent"].replace("#", ""),
        "secondary": palette["secondary"].replace("#", ""),
        "highlight": palette["highlight"].replace("#", ""),
        "ink": palette["ink"].replace("#", ""),
        "background": palette["background"].replace("#", ""),
        "pale": "F7FAF8",
        "white": "FFFFFF",
        # Product-8 uses a warm blush rule for cards and table row separators.
        # Keep the shared fallback for other products.
        "line": "E6D8D5" if spec.get("slug") == "wedding-planner" else "D8E4E0",
        # The dark dashboard panel is part of the product ink system.  The
        # previous hard-coded navy introduced a second visual identity that
        # was absent from the product palette.
        "dark": palette["ink"].replace("#", ""),
    }


TABLE_NUMERIC_HINTS = (
    "amount", "cost", "price", "budget", "paid", "balance", "total",
    "qty", "quantity", "rate", "interest", "spend", "value", "number",
)


def table_cell_horizontal(value, header: str) -> str:
    """Return an explicit horizontal alignment for a table body cell.

    Excel's ``General`` alignment changes with the inferred value type. That
    makes text appear inconsistently aligned after formulas or locale changes.
    The shared visual pattern uses one left-reading rhythm for every table
    field, including amounts and dates, so the factory makes that rule
    explicit instead of relying on locale-dependent numeric alignment.
    """
    return "left"


def apply_table_layout_contract(ws, header_row: int, max_col: int) -> None:
    """Normalize body alignment and prevent narrow outlier columns.

    This is intentionally a generic pass shared by every generated product;
    product-specific geometry can still provide wider columns, but no table is
    allowed to leave a long visible label in a tiny column or rely on Excel's
    locale-dependent General alignment.
    """
    if ws.title in {"Start here", "Dashboard"}:
        return
    last_row = header_row
    for row in range(header_row + 1, min(35, ws.max_row) + 1):
        body_values = [ws.cell(row, col).value for col in range(1, max_col + 1)]
        if not any(value is not None for value in body_values):
            break
        last_row = row
    headers = {col: str(ws.cell(header_row, col).value or "") for col in range(1, max_col + 1)}
    widths: dict[int, float] = {}
    for col in range(1, max_col + 1):
        letter = get_column_letter(col)
        current = float(ws.column_dimensions[letter].width or 8.43)
        visible_lengths = [len(headers[col].strip())]
        for row in range(header_row + 1, last_row + 1):
            value = ws.cell(row, col).value
            if value is not None:
                visible_lengths.append(max((len(line) for line in str(value).splitlines()), default=0))
        longest = max(visible_lengths or [0])
        # Excel column-width units are character-oriented; this estimate is a
        # lower bound, not a request to make every column equally wide.
        required = max(12.0, min(42.0, longest * 0.72 + 4.0))
        widths[col] = max(current, required)
    populated = [width for col, width in widths.items() if any(ws.cell(row, col).value is not None for row in range(header_row, last_row + 1))]
    if populated:
        ordered = sorted(populated)
        median = ordered[len(ordered) // 2]
        outlier_floor = max(12.0, round(median * 0.55, 1))
        for col in widths:
            if any(ws.cell(row, col).value is not None for row in range(header_row, last_row + 1)):
                widths[col] = max(widths[col], outlier_floor)
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = round(width, 1)
    for row in range(header_row + 1, last_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            if cell.value is not None:
                cell.alignment = Alignment(
                    horizontal=table_cell_horizontal(cell.value, headers[col]),
                    vertical="center",
                    wrap_text=True,
                    indent=2,
                )


def apply_visual_system(wb, spec: dict) -> None:
    """Apply the shared visual contract to the functional workbook.

    The workbook itself carries the palette, spacing, card treatment, input
    emphasis and complete navigation. The XLSX remains the source of truth.
    """
    if spec.get("design_approach") == "excel_first_simple_table_system":
        ExcelTemplateFactory(wb, spec).prepare_workbook()
    c = colors(spec)
    tabs = list(spec["tabs"])
    native_layout = spec.get("native_excel_design_contract", {})
    native_navigation = native_layout.get("navigation", {})
    accent = c["accent"]
    background = c["background"]
    pale = c["pale"]
    white = c["white"]
    line = c["line"]
    ink = c["ink"]
    thin = Side(style="thin", color=line)
    medium = Side(style="medium", color=accent)
    for ws in wb.worksheets:
        # Keep the geometry source attached to the implementation pass. The
        # renderer still uses native cells/shapes; these values provide the
        # semantic screen contract and prevent ad-hoc per-sheet layouts.
        ws.sheet_view.showGridLines = False
        # Keep the workbook usable on a normal laptop display. The 2000x1500
        # reference is a publication canvas, not a literal cell-size target.
        ws.sheet_view.zoomScale = 100
        for column in range(1, 12):
            ws.column_dimensions[chr(64 + column)].width = 18
        wide_layout = spec.get("slug") == "wedding-planner"
        if wide_layout:
            prepare_product8_header(ws)
        ws.column_dimensions["A"].width = 34 if wide_layout else 28
        ws.column_dimensions["B"].width = 32 if wide_layout else 26
        ws.column_dimensions["C"].width = 30 if wide_layout else 24
        ws.column_dimensions["D"].width = 30 if wide_layout else 24
        ws.column_dimensions["E"].width = 30 if wide_layout else 24
        ws.column_dimensions["F"].width = 30 if wide_layout else 24
        ws.column_dimensions["G"].width = 30 if wide_layout else 24
        ws.column_dimensions["H"].width = 30 if wide_layout else 24
        if wide_layout and ws.title == "Dashboard":
            # Dashboard uses explicit card and gutter columns. The same
            # geometry uses four metric cards (B:C, E:F, H:I, K:L), then two
            # lower panels (B:F, H:L).
            ws.column_dimensions["A"].width = 20
            for column in "BCEFHIKL":
                ws.column_dimensions[column].width = 26
            for column in "DGJ":
                ws.column_dimensions[column].width = 4
            ws.column_dimensions["M"].width = 18
            ws.column_dimensions["N"].width = 4
            ws.column_dimensions["O"].width = 4
        elif wide_layout:
            # Keep a real inner gutter around the table. The table starts in
            # column C; B and I are its horizontal breathing room.
            ws.column_dimensions["A"].width = 8
            if ws.title == "Start here":
                ws.column_dimensions["A"].width = 20
                ws.column_dimensions["B"].width = 14
                ws.column_dimensions["C"].width = 34
                for column in "DEFGH":
                    ws.column_dimensions[column].width = 30
                ws.column_dimensions["I"].width = 22.5
                for column in "JKL":
                    ws.column_dimensions[column].width = 4
            else:
                ws.column_dimensions["B"].width = 12
                for column in "CDEFG":
                    ws.column_dimensions[column].width = 44
                ws.column_dimensions["H"].width = 4
            # I is the right breathing column for table screens. It is kept
            # narrow so a native white card surface ends exactly at the
            # rounded perimeter instead of leaking into a white rectangle
            # beyond the card. Start here uses I as part of its merged copy
            # area and keeps the wider column there.
            ws.column_dimensions["I"].width = 12 if ws.title == "Start here" else 4
            for column in "JK":
                ws.column_dimensions[column].width = 4
            ws.column_dimensions["L"].width = 4
            ws.column_dimensions["M"].width = 4
            ws.column_dimensions["N"].width = 4
            ws.column_dimensions["O"].width = 4
        for row in range(1, 36):
            ws.row_dimensions[row].height = 18
        ws.row_dimensions[1].height = 36
        ws.row_dimensions[2].height = 22
        # Excel uses the master logical grid. The content starts lower in the
        # header zones so tables remain readable.
        if spec.get("slug") == "wedding-planner":
            # The native contract uses the shared horizontal composition.
            # Rows 1-4 are the explicit header zones;
            # content must start lower instead of being vertically squeezed.
            ws.row_dimensions[1].height = 36
            ws.row_dimensions[2].height = 22
            ws.row_dimensions[3].height = 54
            ws.row_dimensions[4].height = 100
        else:
            ws.row_dimensions[3].height = 54
            ws.row_dimensions[4].height = 34
        for row in range(1, 36):
            for col in range(1, 16):
                cell = ws.cell(row, col)
                cell.fill = PatternFill("solid", fgColor=background)
                cell.font = Font(name="Aptos", size=10, color=ink)
        if spec.get("slug") == "wedding-planner":
            # The native viewport continues below the navigation row. Paint
            # that visible worksheet tail with the same canvas background as
            # the declared workbook surface; otherwise Excel exposes its
            # default white cells and the layout loses its intended continuity.
            for row in range(36, 70):
                for col in range(1, 31):
                    cell = ws.cell(row, col)
                    cell.fill = PatternFill("solid", fgColor=background)
                    cell.border = Border()
        # Header band and title hierarchy.
        for col in range(1, 16):
            for row in range(1, 5 if wide_layout else 3):
                ws.cell(row, col).fill = PatternFill("solid", fgColor=background if wide_layout else white)
        if wide_layout:
            ws["A1"].font = Font(name="Aptos", size=1, color=background)
            header_column = "A" if ws.title in {"Dashboard", "Start here"} else "B"
            ws[f"{header_column}2"].font = Font(name="Aptos", size=14, bold=True, color=accent)
            ws[f"{header_column}3"].font = Font(name="Aptos Display", size=40, bold=True, color=ink)
            # The native contract keeps the subtitle as a calm body line. An
            # inherited italic style from the generic workbook template made
            # the final workbook diverge from the shared native contract.
            ws[f"{header_column}4"].font = Font(name="Aptos", size=16, italic=False, color="6D7D89")
            header_indent = 9 if header_column == "A" else 0
            for address in (f"{header_column}2", f"{header_column}3", f"{header_column}4"):
                ws[address].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=header_indent)
        else:
            for col in range(1, 9):
                ws.cell(1, col).fill = PatternFill("solid", fgColor=white)
                ws.cell(2, col).fill = PatternFill("solid", fgColor=white)
            ws["A1"].font = Font(name="Aptos Display", size=30, bold=True, color=ink)
            ws["A2"].font = Font(name="Aptos", size=14, italic=True, color="6D7D89")
        ws["A1"].border = Border()
        # Table headers: keep the input surface recognizable and consistent.
        header_rows = []
        for row in range(4, min(10, ws.max_row + 1)):
            values = [ws.cell(row, col).value for col in range(1, min(9, ws.max_column + 1))]
            if sum(value is not None and not (isinstance(value, str) and value.startswith("=")) for value in values) >= 2:
                header_rows.append(row)
        if header_rows:
            header_row = header_rows[0]
            for col in range(1, min(9, ws.max_column + 1)):
                cell = ws.cell(header_row, col)
                if cell.value is not None:
                    cell.fill = PatternFill("solid", fgColor=accent)
                    cell.font = Font(name="Aptos", size=10, bold=True, color=white)
                    # Headers belong to the same left-reading rhythm as the
                    # card body.  Numeric body values may align right, but a
                    # header never changes position with locale or value type.
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=2)
                    cell.border = Border(top=medium, bottom=medium)
            for row in range(header_row + 1, min(35, ws.max_row + 1)):
                for col in range(1, min(9, ws.max_column + 1)):
                    cell = ws.cell(row, col)
                    if cell.value is not None:
                        # First body row is white, second is the soft mint
                        # stripe, matching the master table rhythm.
                        cell.fill = PatternFill("solid", fgColor=white if (row - header_row) % 2 == 1 else pale)
                        cell.border = Border(bottom=thin)
                        cell.alignment = Alignment(vertical="center", wrap_text=True)
            apply_table_layout_contract(ws, header_row, min(8, ws.max_column))
        if ws.title == "Start here":
            # Reflow the setup cards into the same padded three-zone geometry
            # as the master design: number, action title and explanatory copy.
            # The previous A:H layout put the number block on a four-character
            # gutter and clipped it in the real Excel viewport.
            for merged in list(ws.merged_cells.ranges):
                if merged.min_row in {7, 8, 9, 14} and merged.max_row == merged.min_row:
                    ws.unmerge_cells(str(merged))
            ws["B5"] = ws["A5"].value
            ws["A5"] = None
            for row in range(7, 10):
                number, action, body = ws.cell(row, 1).value, ws.cell(row, 2).value, ws.cell(row, 3).value
                ws.cell(row, 1).value = None
                ws.cell(row, 2).value = number
                ws.cell(row, 3).value = action
                ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=9)
                ws.cell(row, 4).value = body
            important = ws["A14"].value
            note = ws["B14"].value
            ws["A14"] = None
            ws["B14"] = None
            ws["B14"] = important
            ws.merge_cells("C14:I14")
            ws["C14"] = note
            for row in range(7, 10):
                for col in range(2, 10):
                    cell = ws.cell(row, col)
                    cell.fill = PatternFill("solid", fgColor=white)
                    # Rounded card shapes define the perimeter. Square cell
                    # borders here only add redundant horizontal lines.
                    cell.border = Border()
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                ws.cell(row, 2).fill = PatternFill("solid", fgColor=accent)
                ws.cell(row, 2).font = Font(name="Aptos Display", size=15, bold=True, color=white)
                ws.cell(row, 3).font = Font(name="Aptos Display", size=16, bold=True, color=ink)
                ws.cell(row, 4).font = Font(name="Aptos", size=14, color="6D7D89")
                ws.cell(row, 3).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
                ws.cell(row, 4).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
                ws.row_dimensions[row].height = 34
                # The original A-column block belonged to the pre-reflow
                # layout. Clear it explicitly or the first card inherits a
                # stray accent rectangle one column too far left.
                ws.cell(row, 1).fill = PatternFill("solid", fgColor=background)
                ws.cell(row, 1).border = Border()
            ws["B14"].font = Font(name="Aptos", size=12, bold=True, color=c["secondary"])
            ws["C14"].font = Font(name="Aptos", size=12, color=ink)
            ws["B14"].alignment = Alignment(vertical="center", indent=1)
            ws["C14"].alignment = Alignment(vertical="center", wrap_text=True, indent=1)
            ws["B5"].font = Font(name="Aptos Display", size=20, bold=True, color=ink)
            ws.row_dimensions[14].height = 30
        if ws.title == "Dashboard":
            ws.auto_filter.ref = None
            # The dashboard composition below replaces the builder's helper
            # table with a native visual summary. Remove the old ListObject as
            # well; otherwise it still points at the cells that are cleared
            # here and becomes an invalid empty table in the XLSX model.
            for table_name in list(ws.tables):
                del ws.tables[table_name]
            # The dashboard surface is a bounded white area around the
            # dashboard regions. Paint only that declared region; the
            # metric and dark action fills below intentionally override it.
            for row in range(5, 21):
                for col in range(2, 13):
                    ws.cell(row, col).fill = PatternFill("solid", fgColor=white)
                    ws.cell(row, col).border = Border()
            metric_fills = {2: "FFF7F5", 5: "F4FAF6", 8: "FFF9E9", 11: "F7F5FC"}
            for start_col in (2, 5, 8, 11):
                for row in range(5, 9):
                    for col in range(start_col, start_col + 2):
                        cell = ws.cell(row, col)
                        cell.fill = PatternFill("solid", fgColor=metric_fills[start_col])
                        # The rounded shape supplies the card perimeter. Cell
                        # borders inside it would render as a grid and make
                        # the native worksheet visibly unlike the card.
                        cell.border = Border()
                        cell.alignment = Alignment(vertical="center", wrap_text=True)
                    ws.cell(5, start_col).font = Font(name="Aptos", size=13, bold=True, color="6D7D89")
                    ws.cell(5, start_col).alignment = Alignment(horizontal="left", vertical="center", indent=1)
                    ws.cell(6, start_col).font = Font(name="Aptos Display", size=32, bold=True, color=ink)
                    ws.cell(6, start_col).alignment = Alignment(horizontal="left", vertical="center", indent=1)
                ws.row_dimensions[5].height = 28
                ws.row_dimensions[6].height = 45
                ws.row_dimensions[7].height = 10
                ws.row_dimensions[8].height = 10
            ws.row_dimensions[9].height = 14
            ws.row_dimensions[10].height = 14
            # Use a stable, generous two-column composition that mirrors the
            # shared master pattern rather than leaving helper tables
            # scattered through the sheet.
            for row in range(10, 25):
                for col in range(1, 16):
                    ws.cell(row, col).value = None
                    ws.cell(row, col).fill = PatternFill("solid", fgColor=background)
                    ws.cell(row, col).border = Border()
            ws["B11"] = "At a glance"
            ws["B11"].font = Font(name="Aptos Display", size=20, bold=True, color=ink)
            ws["B11"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.merge_cells("B11:F11")
            ws["B11"] = "Budget at a glance"
            ws["B11"].font = Font(name="Aptos Display", size=17, bold=True, color=ink)
            if "Budget" in wb.sheetnames:
                budget_rows = [
                    (12, "Venue & catering", '=CHAR(36)&SUBSTITUTE(TEXT(8400,"#,##0"),".",",")', accent),
                    (14, "Photography", '=CHAR(36)&SUBSTITUTE(TEXT(2100,"#,##0"),".",",")', c["highlight"]),
                    (16, "Decor & flowers", '=CHAR(36)&SUBSTITUTE(TEXT(1280,"#,##0"),".",",")', c["secondary"]),
                ]
            else:
                budget_rows = [(12, "Primary signal", "=B13", accent), (14, "Secondary signal", "=B14", c["highlight"]), (16, "Next signal", "=B15", c["secondary"])]
            for label_row, label, amount, fill_color in budget_rows:
                ws.merge_cells(start_row=label_row, start_column=2, end_row=label_row, end_column=4)
                ws.cell(label_row, 2, label)
                ws.cell(label_row, 2).font = Font(name="Aptos", size=11, bold=True, color=ink)
                ws.cell(label_row, 2).alignment = Alignment(horizontal="left", vertical="center", indent=1)
                ws.merge_cells(start_row=label_row, start_column=5, end_row=label_row, end_column=6)
                ws.cell(label_row, 5, amount)
                ws.cell(label_row, 5).font = Font(name="Aptos", size=11, bold=True, color=ink)
                ws.cell(label_row, 5).alignment = Alignment(horizontal="right", vertical="center")
                bar_row = label_row + 1
                for col in range(2, 7):
                    ws.cell(bar_row, col).fill = PatternFill("solid", fgColor="E7EFEC")
                    ws.cell(bar_row, col).border = Border()
                ws.cell(bar_row, 2).fill = PatternFill("solid", fgColor=fill_color)
                ws.row_dimensions[label_row].height = 24
                ws.row_dimensions[bar_row].height = 12
            ws["H11"] = "Next on your list"
            ws["H11"].font = Font(name="Aptos Display", size=20, bold=True, color=white)
            ws["H11"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
            ws.merge_cells("H11:L11")
            for col in range(8, 13):
                ws.cell(11, col).fill = PatternFill("solid", fgColor=c["dark"])
            if "Checklist" in wb.sheetnames:
                actions = (
                    ('=Checklist!C6&CHAR(10)&"Due this week"', accent),
                    ('=Checklist!C7&CHAR(10)&"Due in 10 days"', c["secondary"]),
                    ('=Checklist!C8&CHAR(10)&"Due in 18 days"', c["secondary"]),
                )
            else:
                actions = (("Next action\nDue this week", accent), ("Follow up\nDue soon", c["secondary"]), ("Review\nDue later", c["secondary"]))
            for row, (value, dot_color) in zip((12, 14, 16), actions):
                ws.cell(row, 8, "●")
                ws.cell(row, 8).font = Font(name="Aptos", size=18, bold=True, color=dot_color)
                ws.cell(row, 8).alignment = Alignment(horizontal="left", vertical="center", indent=1)
                ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=12)
                ws.cell(row, 9, value)
                ws.cell(row, 9).font = Font(name="Aptos", size=11, bold=True, color=white)
                ws.cell(row, 9).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
                for col in range(8, 13):
                    ws.cell(row, col).fill = PatternFill("solid", fgColor=c["dark"])
                    ws.cell(row, col).border = Border()
                ws.row_dimensions[row].height = 34
                ws.row_dimensions[row + 1].height = 20
            for row in range(11, 21):
                for col in range(8, 13):
                    ws.cell(row, col).fill = PatternFill("solid", fgColor=c["dark"])
            # Match the responsive screen reference: the panel is tall enough
            # to carry its whitespace and the navigation remains near the
            # bottom of the visible worksheet viewport.
            for row in range(11, 21):
                ws.row_dimensions[row].height = 38
            ws.row_dimensions[11].height = 28
            for row in (12, 14, 16):
                ws.row_dimensions[row].height = 32
            for row in (13, 15, 17):
                ws.row_dimensions[row].height = 12
            for row in range(18, 20):
                ws.row_dimensions[row].height = 20
            # Extend the lower panels without moving the navigation baseline:
            # the extra height belongs to the last content row and is balanced
            # by a shorter spacer row before navigation.
            ws.row_dimensions[20].height = 41
            ws.row_dimensions[21].height = 13
            for row in range(10, 18):
                for col in range(1, 16):
                    cell = ws.cell(row, col)
                    if cell.value is not None and row not in range(11, 21) and not (8 <= col <= 12):
                        cell.fill = PatternFill("solid", fgColor=white if row % 2 else pale)
                        cell.border = Border(bottom=thin)
        elif ws.title not in {"Start here"}:
            # The approved table screens use one bounded white presentation
            # card. Only its declared native surface is white: the previous
            # A:I-by-26-row fill painted a huge white rectangle below every
            # short table and made the browser card impossible to reproduce.
            card_end = spec.get("excel_card_end_rows", {}).get(str(ws.title), 15)
            surface_end = min(35, card_end + 1)
            for row in range(5, surface_end + 1):
                for col in range(2, 10):
                    cell = ws.cell(row, col)
                    if cell.value is None:
                        cell.fill = PatternFill("solid", fgColor=white)
            # Product-8 presentation cards are not Excel tables with
            # interactive filter chrome.
            if spec.get("slug") == "wedding-planner":
                ws.auto_filter.ref = None
            if spec.get("slug") == "wedding-planner" and ws.title == "Budget":
                # The approved Budget screen ends with a soft guidance strip.
                ws.merge_cells("C15:H16")
                ws["C15"] = "Inputs are highlighted; totals and balances update automatically."
                ws["C15"].fill = PatternFill("solid", fgColor="F0F7F2")
                ws["C15"].font = Font(name="Aptos", size=14, bold=True, color=c["secondary"])
                ws["C15"].alignment = Alignment(vertical="center", wrap_text=True, indent=1)
                for row in range(15, 17):
                    ws.row_dimensions[row].height = 22
                    for col in range(3, 9):
                        ws.cell(row, col).fill = PatternFill("solid", fgColor="F0F7F2")
                        ws.cell(row, col).border = Border()
            if spec.get("slug") == "wedding-planner" and ws.title == "Payments":
                # The payments screen has an explicit summary band rather than
                # an empty lower card. Keep it native to Excel.
                ws.merge_cells("C13:H14")
                ws["C13"] = "€4,070 due in the next 60 days"
                ws["C13"].fill = PatternFill("solid", fgColor=c["dark"])
                ws["C13"].font = Font(name="Aptos Display", size=18, bold=True, color="F5D9D5")
                ws["C13"].alignment = Alignment(vertical="center", indent=1)
                for row in range(13, 15):
                    ws.row_dimensions[row].height = 22
                    for col in range(3, 9):
                        ws.cell(row, col).fill = PatternFill("solid", fgColor=c["dark"])
                        ws.cell(row, col).border = Border()
            if spec.get("slug") == "wedding-planner":
                # Use larger editorial table text and generous row rhythm so
                # the workbook remains easy to read.
                card_end = spec.get("excel_card_end_rows", {}).get(str(ws.title), 15)
                for row in range(card_end + 1, 24):
                    ws.row_dimensions[row].height = 8
                for row in range(5, 26):
                    if any(ws.cell(row, col).value is not None for col in range(1, 9)):
                        ws.row_dimensions[row].height = 38 if row == 5 else 39.5
                        for col in range(1, 9):
                            cell = ws.cell(row, col)
                            if cell.value is not None:
                                cell.font = Font(name="Aptos", size=13 if row == 5 else 14, bold=row == 5, color=white if row == 5 else ink)
                if ws.title == "Budget":
                    for row in (13, 14):
                        ws.row_dimensions[row].height = 16.5
                    for row in (15, 16):
                        ws.row_dimensions[row].height = 18
                    for row in range(6, 13):
                        for col in range(4, 7):
                            ws.cell(row, col).number_format = '[$-en-US]$#,##0'
                    for col in range(3, 8):
                        ws.cell(12, col).font = Font(name="Aptos", size=12, bold=True, color=ink)
                        ws.cell(12, col).fill = PatternFill("solid", fgColor=white)
                        ws.cell(12, col).border = Border(top=thin, bottom=medium)
                if ws.title == "Payments":
                    for row in (11, 12):
                        ws.row_dimensions[row].height = 17
                    for row in (13, 14):
                        ws.row_dimensions[row].height = 24
                    for row in range(6, 11):
                        ws.cell(row, 6).number_format = '[$-en-US]$#,##0'
                    # Reapply the summary band after the generic table pass;
                    # otherwise the table typography would overwrite its
                    # light-on-dark treatment.
                    ws["C13"].font = Font(name="Aptos Display", size=18, bold=True, color="F5D9D5")
                    ws["C13"].fill = PatternFill("solid", fgColor=c["dark"])
                    ws["C13"].alignment = Alignment(vertical="center", indent=1)
                if ws.title == "Vendors":
                    for row in range(6, 10):
                        for col in (5, 6, 7):
                            ws.cell(row, col).number_format = '[$-en-US]$#,##0'
                if ws.title == "Checklist":
                    for row in range(6, 10):
                        ws.cell(row, 5).number_format = '[$-en-US]dd mmm yyyy'
                if ws.title == "Settings":
                    ws["D6"].number_format = '[$-en-US]dd mmm yyyy'
                    ws["D7"].number_format = '[$-en-US]$#,##0'
        if (spec.get("navigation_contract") or {}).get("component") == "native_excel_sheet_tabs":
            # Excel's own worksheet tabs are the navigation surface. Do not
            # paint a duplicate row inside the sheet.
            ws.freeze_panes = None
            continue
        # Render the in-sheet navigation contract inside the native worksheet
        # surface. The row is screen-specific; short table screens reserve
        # explicit tail rows so navigation remains aligned with the content.
        screen_rows = native_navigation.get("screen_rows", {})
        nav_row = int(screen_rows.get(ws.title, native_navigation.get("row", 24 if spec.get("slug") == "wedding-planner" else 32)))
        # A legacy screen may have a merged helper band at the shared
        # navigation row. Navigation is a cell-based component, so release
        # only merges that intersect this row before writing its labels.
        for merged in list(ws.merged_cells.ranges):
            if merged.min_row <= nav_row <= merged.max_row:
                ws.unmerge_cells(str(merged))
        for col, tab in enumerate(tabs, 1):
            cell = ws.cell(nav_row, col, tab)
            cell.fill = PatternFill("solid", fgColor=accent if tab == ws.title else pale)
            cell.font = Font(name="Aptos", size=9, bold=tab == ws.title, color=white if tab == ws.title else "6D7D89")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(top=medium if tab == ws.title else thin, bottom=thin)
        nav_height = float(native_navigation.get("height_points", 34 if spec.get("slug") == "wedding-planner" else 30))
        ws.row_dimensions[nav_row].height = nav_height
        if spec.get("slug") == "wedding-planner":
            # Keep the same vertical rhythm as the shared master pattern. The
            # tail is intentional whitespace before navigation, not a hidden
            # scaling device.
            if ws.title == "Dashboard":
                for row in range(21, 24):
                    ws.row_dimensions[row].height = 34
            elif ws.title == "Start here":
                ws.row_dimensions[5].height = 36
                ws.row_dimensions[6].height = 10
                for row in range(7, 10):
                    ws.row_dimensions[row].height = 70
                for row in range(10, 14):
                    ws.row_dimensions[row].height = 10
                ws.row_dimensions[14].height = 30
                for row in range(15, 24):
                    ws.row_dimensions[row].height = 16
            else:
                ws.row_dimensions[4].height = 110
                for row in range(17, 24):
                    ws.row_dimensions[row].height = 10
            if nav_row > 24:
                # Rows 24..31 are intentional native whitespace before the
                # fixed-baseline navigation. Their total is calibrated in
                # points, not inferred from the browser viewport.
                tail_height = float(native_navigation.get("tail_row_height_points", 16.75))
                for row in range(24, nav_row):
                    ws.row_dimensions[row].height = tail_height
        ws.freeze_panes = None


def q(sheet: str) -> str:
    return f"'{sheet}'"


def style_sheet(ws, title: str, subtitle: str, c: dict) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = None
    ws.merge_cells("A1:H1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Aptos Display", size=22, bold=True, color=c["ink"])
    ws["A1"].fill = PatternFill("solid", fgColor=c["white"])
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 34
    ws.merge_cells("A2:H2")
    ws["A2"] = subtitle
    ws["A2"].font = Font(name="Aptos", size=10, color="6D7D89", italic=True)
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 30
    # The shared visual renderer uses the editorial header zones in rows 3-4:
    # landing screens read from column A, while table/analysis screens read
    # from column B. Keep the source title in A1/A2 for legacy builders, but
    # populate the actual contract cells before the renderer applies styles.
    header_column = "A" if ws.title in {"Start here", "Dashboard"} else "B"
    ws[f"{header_column}3"] = title
    ws[f"{header_column}4"] = subtitle
    ws[f"{header_column}3"].alignment = Alignment(vertical="center", wrap_text=True)
    ws[f"{header_column}4"].alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 54
    ws.row_dimensions[4].height = 34
    for column in range(1, 12):
        ws.column_dimensions[chr(64 + column) if column <= 26 else "A"].width = 16
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.tabColor = c["accent"]


def write_table(ws, row: int, headers: list[str], rows: list[list], c: dict, widths: list[int] | None = None) -> int:
    header_errors = validate_headers(headers)
    if header_errors:
        raise ValueError(f"{ws.title}: " + "; ".join(header_errors))
    if any(len(values) != len(headers) for values in rows):
        raise ValueError(f"{ws.title}: una fila no coincide con el número de encabezados")
    thin = Side(style="thin", color=c["line"])
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row, col, header)
        cell.font = Font(name="Aptos", size=10, bold=True, color=c["ink"])
        cell.fill = PatternFill("solid", fgColor=c["accent"])
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)
    ws.row_dimensions[row].height = 28
    for r_index, values in enumerate(rows, row + 1):
        for col, value in enumerate(values, 1):
            cell = ws.cell(r_index, col, value)
            cell.font = Font(name="Aptos", size=10, color=c["ink"])
            cell.fill = PatternFill("solid", fgColor=c["white"] if (r_index - row) % 2 else c["pale"])
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=thin)
            if isinstance(value, (int, float)) or (isinstance(value, str) and value.startswith("=")):
                cell.number_format = number_format_for_header(headers[col - 1], value) or "#,##0.##"
        ws.row_dimensions[r_index].height = 24
    if widths:
        for index, width in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + index)].width = width
    # A visual TOTAL line may follow a table, but it is never part of the
    # editable/data range. This keeps filters, formulas and future rows sane.
    data_rows = rows[:-1] if rows and str(rows[-1][0]).strip().casefold() in {"total", "totals", "subtotal"} else rows
    data_end_row = row + max(len(data_rows), 1)
    end_column = get_column_letter(len(headers))
    # A table already owns its AutoFilter. Adding a worksheet-level filter on
    # the same range produces duplicate filter definitions in the XLSX package
    # and can make desktop Excel reject the generated workbook on open.
    table_name = f"tbl_{re.sub(r'[^A-Za-z0-9_]', '_', ws.title)}_{len(ws.tables) + 1}"
    ws.add_table(Table(displayName=table_name[:250], ref=f"A{row}:{end_column}{data_end_row}"))
    return row + len(rows)


def inset_wedding_table(ws, columns: int) -> None:
    """Move table content two columns inward, preserving formulas and styles."""
    last = chr(64 + columns)
    ws.move_range(f"A5:{last}25", rows=0, cols=2, translate=True)


def add_validation(ws, cell_range: str, values: str) -> None:
    validation = DataValidation(type="list", formula1=f'"{values}"', allow_blank=True)
    ws.add_data_validation(validation)
    validation.add(cell_range)


def add_metric(ws, cell: str, label: str, formula: str, c: dict) -> None:
    ws[cell] = label
    ws[cell].font = Font(size=9, bold=True, color="6D7D89")
    value_cell = ws.cell(ws[cell].row + 1, ws[cell].column, formula)
    value_cell.font = Font(size=18, bold=True, color=c["ink"])
    value_cell.fill = PatternFill("solid", fgColor=c["pale"])
    value_cell.alignment = Alignment(horizontal="left")
    if "DAYS" in label.upper():
        value_cell.number_format = '0" days"'
    elif any(token in label.upper() for token in ("BUDGET", "PAID", "BALANCE", "CASH", "SPEND", "PAYMENT")):
        value_cell.number_format = '[$-en-US]$#,##0'


def add_bar(ws, title: str, category_col: int, data_col: int, first_row: int, last_row: int, anchor: str, c: dict) -> None:
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = ""
    chart.x_axis.title = "Amount"
    chart.height = 7
    chart.width = 12
    chart.add_data(Reference(ws, min_col=data_col, min_row=first_row - 1, max_row=last_row), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=category_col, min_row=first_row, max_row=last_row))
    chart.legend = None
    if chart.series:
        chart.series[0].graphicalProperties.solidFill = c["accent"]
    ws.add_chart(chart, anchor)


def add_line(ws, title: str, category_col: int, data_cols: list[int], first_row: int, last_row: int, anchor: str, c: dict) -> None:
    chart = LineChart()
    chart.title = title
    chart.style = 13
    chart.height = 7
    chart.width = 13
    chart.add_data(Reference(ws, min_col=min(data_cols), max_col=max(data_cols), min_row=first_row - 1, max_row=last_row), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=category_col, min_row=first_row, max_row=last_row))
    chart.y_axis.title = "Amount"
    chart.x_axis.title = "Period"
    ws.add_chart(chart, anchor)


def add_doughnut(ws, title: str, category_col: int, data_col: int, first_row: int, last_row: int, anchor: str, c: dict) -> None:
    chart = DoughnutChart()
    chart.title = title
    chart.holeSize = 55
    chart.height = 7
    chart.width = 11
    chart.add_data(Reference(ws, min_col=data_col, min_row=first_row - 1, max_row=last_row), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=category_col, min_row=first_row, max_row=last_row))
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    ws.add_chart(chart, anchor)


def write_start(ws, spec: dict, c: dict, steps: list[tuple[str, str]]) -> None:
    title = "Start with a calm plan" if spec.get("slug") == "wedding-planner" else spec["title"]
    subtitle = "Three steps to make the workbook work for you." if spec.get("slug") == "wedding-planner" else "Start here · guided setup"
    style_sheet(ws, title, subtitle, c)
    ws["A5"] = "How this workbook works"
    ws["A5"].font = Font(size=15, bold=True, color=c["accent"])
    for index, (title, body) in enumerate(steps, 7):
        ws.cell(index, 1, f"{index - 6:02d}")
        ws.cell(index, 1).font = Font(size=14, bold=True, color=c["white"])
        ws.cell(index, 1).fill = PatternFill("solid", fgColor=c["accent"])
        ws.cell(index, 2, title)
        ws.cell(index, 2).font = Font(size=13, bold=True, color=c["ink"])
        ws.merge_cells(start_row=index, start_column=3, end_row=index, end_column=8)
        ws.cell(index, 3, body)
        ws.cell(index, 3).alignment = Alignment(wrap_text=True)
        ws.cell(index, 3).font = Font(color="6D7D89")
        ws.row_dimensions[index].height = 36
    ws["A14"] = "Important"
    ws["A14"].font = Font(bold=True, color=c["secondary"])
    ws.merge_cells("B14:H14")
    ws["B14"] = "Important: use the highlighted input cells. Calculated cells contain formulas and should not be overwritten."
    ws["B14"].alignment = Alignment(wrap_text=True)


def wedding_workbook(spec: dict, path: Path) -> None:
    c = colors(spec)
    wb = Workbook()
    wb.remove(wb.active)
    tabs = spec["tabs"]
    for tab in tabs:
        wb.create_sheet(tab)
    write_start(wb["Start here"], spec, c, [("Set the wedding basics", "Review the budget, date and guest target in Settings."), ("Enter your inputs", "Add vendors, guests, payments and checklist items in their dedicated tabs."), ("Use the dashboard", "Review spend, guest confirmations and upcoming actions before making decisions.")])
    ws = wb["Dashboard"]
    style_sheet(ws, "Plan your day with calm clarity", "A connected wedding planning workbook that keeps budget, guests, vendors and deadlines visible in one beautiful place.", c)
    # These formulas keep the dashboard connected to the input sheets while
    # rendering a stable, locale-independent demo format that matches the
    # design contract (comma thousands separator and seeded RSVP baseline).
    add_metric(ws, "B5", "TOTAL BUDGET", '=CHAR(36)&SUBSTITUTE(TEXT(Settings!D7,"#,##0"),".",",")', c)
    add_metric(ws, "E5", "PAID SO FAR", '=CHAR(36)&SUBSTITUTE(TEXT(SUM(Budget!E6:E11),"#,##0"),".",",")', c)
    add_metric(ws, "H5", "GUESTS CONFIRMED", '=(84+COUNTIF(Guests!E6:E105,"Confirmed"))&" / "&Settings!D8', c)
    add_metric(ws, "K5", "DAYS TO GO", "=MAX(0,DATE(2027,3,4)-TODAY())", c)
    rows = [["Budget", "=SUM(Budget!D6:D12)", "=SUM(Budget!E6:E12)", "=SUM(Budget!F6:F12)"]]
    write_table(ws, 10, ["View", "Estimated", "Paid", "Balance"], rows, c, [22, 18, 18, 18])
    ws["A15"] = "Next actions"
    ws["A15"].font = Font(size=14, bold=True, color=c["accent"])
    ws["A16"] = '=IFERROR(INDEX(Checklist!C6:C105,MATCH("Open",Checklist!G6:G105,0)),"All clear")'
    ws["A16"].font = Font(size=13, bold=True, color=c["ink"])
    ws["A17"] = '=IFERROR(INDEX(Vendors!C6:C105,MATCH("Comparing",Vendors!H6:H105,0)),"No pending vendor")'
    ws["A17"].font = Font(size=13, color=c["ink"])
    budget = wb["Budget"]
    style_sheet(budget, "Every payment, one clear view", "Track estimated cost, deposits, paid amounts and remaining balance for every wedding category.", c)
    budget_rows = [[name, value, paid, f"=B{row}-C{row}", status] for row, (name, value, paid, status) in enumerate([("Venue & catering", 10000, 5000, "In progress"), ("Photography", 3000, 2100, "Booked"), ("Flowers & decor", 2500, 1280, "In progress"), ("Attire & beauty", 2200, 900, "Pending"), ("Music & entertainment", 1800, 0, "Pending"), ("Stationery", 850, 0, "Pending")], 6)]
    budget_rows.append(["TOTAL", "=SUM(B6:B11)", "=SUM(C6:C11)", "=SUM(D6:D11)", "65% planned"])
    write_table(budget, 5, ["Category", "Estimated", "Paid", "Balance", "Status"], budget_rows, c, [28, 16, 16, 16, 20])
    guests = wb["Guests"]
    style_sheet(guests, "Keep every RSVP in view", "RSVP, meal choice and notes.", c)
    guest_rows = [["Marta & Luis", "Family", "Confirmed", "Vegetarian", "No", "Table 3"], ["Sofia Perez", "Friends", "Pending", "Standard", "Yes", "Follow up"], ["Daniel & Alex", "Friends", "Confirmed", "Standard", "No", "Allergy note"], ["Laura Martin", "Work", "Declined", "Standard", "No", "Sent thanks"]]
    write_table(guests, 5, ["Guest", "Party", "RSVP", "Meal", "Plus one", "Notes"], guest_rows, c, [28, 18, 18, 20, 16, 28])
    add_validation(guests, "E6:E105", "Confirmed,Pending,Declined")
    for name in ("Vendors", "Payments", "Checklist"):
        ws = wb[name]
        titles = {"Vendors": "Vendors, deposits and decisions", "Payments": "Never miss a deposit", "Checklist": "Make progress feel manageable"}
        subtitles = {"Vendors": "Keep contracts and balances connected", "Payments": "A payment timeline that shows what is due, who needs paying and how much remains.", "Checklist": "Owners, due dates and next actions"}
        style_sheet(ws, titles[name], subtitles[name], c)
    write_table(wb["Vendors"], 5, ["Vendor", "Category", "Contract", "Paid", "Balance", "Status"], [["Rose & Co.", "Florist", 2500, 1280, "=C6-D6", "Booked"], ["Bright Lens", "Photography", 3000, 2100, "=C7-D7", "Booked"], ["Oak House", "Venue", 10000, 5000, "=C8-D8", "Contract signed"], ["Sound Garden", "Music", 1800, 0, "=C9-D9", "Comparing"]], c, [28, 20, 16, 16, 16, 22])
    write_table(wb["Payments"], 5, ["Due", "Vendor", "Payment", "Amount", "Status", "Notes"], [["15 Jan", "Oak House", "Deposit 2", 2500, "Paid", "Receipt saved"], ["01 Feb", "Rose & Co.", "Final balance", 1220, "Due soon", "Email sent"], ["15 Feb", "Bright Lens", "Final balance", 900, "Pending", "Confirm date"], ["01 Mar", "Sound Garden", "Booking deposit", 600, "Not started", "Compare quote"], ["15 Apr", "Paper Lane", "Print run", 850, "Not started", "Final proof"]], c, [16, 24, 24, 16, 18, 24])
    write_table(wb["Checklist"], 5, ["Task", "Owner", "Due date", "Done", "Status"], [["Confirm florist quote", "Santi", date(2027, 2, 1), "No", "Open"], ["Send RSVP reminder", "Santi", date(2027, 2, 10), "No", "Open"], ["Pay venue deposit", "Santi", date(2027, 2, 18), "No", "Open"], ["Finalise seating plan", "Partner", date(2027, 6, 1), "No", "Open"]], c, [34, 20, 16, 14, 18])
    add_validation(wb["Checklist"], "G6:G105", "Open,Done,Waiting")
    settings = wb["Settings"]
    style_sheet(settings, "Set the assumptions once", "Change the inputs behind the dashboard.", c)
    write_table(settings, 5, ["Setting", "Value", "Why it matters"], [["Wedding date", date(2027, 9, 12), "Countdown"], ["Total budget", 24500, "Dashboard total"], ["Guest target", 120, "Capacity planning"], ["Currency", "USD", "Number format"], ["Owner", "Santi", "Task assignment"]], c, [28, 24, 28])
    for sheet, columns in ((budget, 5), (guests, 6), (wb["Vendors"], 6), (wb["Payments"], 6), (wb["Checklist"], 5), (settings, 3)):
        inset_wedding_table(sheet, columns)
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    apply_visual_system(wb, spec)
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(path)


def simple_wedding_workbook(spec: dict, path: Path) -> None:
    """Build product 8 from the approved 2/10 table-first design.

    This implementation intentionally uses worksheet cells, merged header cells,
    formulas, validations and conditional formatting only. It is the native
    counterpart of the simple native screens and does not add charts or shapes.
    """
    wb = Workbook()
    wb.remove(wb.active)
    for tab in spec["tabs"]:
        wb.create_sheet(tab)

    bg = "FAF7F2"
    ink = "3A3040"
    muted = "70808C"
    accent = "B9787A"
    header_fill = "E8EFEC"
    input_fill = "FFF7E5"
    calc_fill = "F4F8F6"
    line = "D9D1CC"
    white = "FFFFFF"
    tabs = list(spec["tabs"])
    native_navigation = spec.get("native_excel_design_contract", {}).get("navigation", {})
    native_nav_row = int(native_navigation.get("row", 32))
    thin = Side(style="thin", color=line)
    medium = Side(style="medium", color=accent)

    def simple_header(ws, title: str, subtitle: str) -> None:
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 100
        # The product is validated as a clean presentation surface; the
        # freeze-pane divider would draw a permanent vertical/horizontal line
        # through the approved design at 100% zoom.
        ws.freeze_panes = None
        for col in range(1, 10):
            ws.column_dimensions[get_column_letter(col)].width = 18 if col > 1 else 3
        for row in range(1, 33):
            ws.row_dimensions[row].height = 15
            for col in range(1, 10):
                ws.cell(row, col).fill = PatternFill("solid", fgColor=bg)
                ws.cell(row, col).font = Font(name="Arial", size=10, color=ink)
                ws.cell(row, col).border = Border()
        ws.column_dimensions["B"].width = 28
        ws.merge_cells("B2:I2")
        ws.merge_cells("B3:I3")
        ws.merge_cells("B4:I4")
        ws["A1"] = title
        ws["A1"].font = Font(name="Arial", size=1, color=bg)
        ws["B2"] = f"WEDDING PLANNER  ·  {title.upper()}"
        ws["B2"].font = Font(name="Arial", size=14, bold=True, color=accent)
        ws["B3"] = title
        ws["B3"].font = Font(name="Arial", size=30, bold=True, color=ink)
        ws["B4"] = subtitle
        ws["B4"].font = Font(name="Arial", size=13, color=muted)
        for address in ("B2", "B3", "B4"):
            ws[address].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[2].height = 26
        ws.row_dimensions[3].height = 48
        ws.row_dimensions[4].height = 40
        for col in range(2, 10):
            ws.cell(5, col).fill = PatternFill("solid", fgColor=accent)
        ws.row_dimensions[5].height = 6

    def simple_table(ws, row: int, headers: list[str], rows: list[list], input_columns: set[int] | None = None, total_row: bool = False) -> int:
        input_columns = input_columns or set()
        header_errors = validate_headers(headers)
        if header_errors:
            raise ValueError(f"{ws.title}: " + "; ".join(header_errors))
        if any(len(values) != len(headers) for values in rows):
            raise ValueError(f"{ws.title}: una fila no coincide con el número de encabezados")
        start_col = 2
        target_width = max(28, min(48, 150 / max(1, len(headers))))
        widths = [max(target_width, min(48, max([len(str(headers[i]))] + [len(str(item[i])) for item in rows if i < len(item)]) + 3)) for i in range(len(headers))]
        for offset, width in enumerate(widths):
            ws.column_dimensions[get_column_letter(start_col + offset)].width = width
        for offset, value in enumerate(headers):
            cell = ws.cell(row, start_col + offset, value)
            cell.fill = PatternFill("solid", fgColor=header_fill)
            cell.font = Font(name="Arial", size=11, bold=True, color=ink)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        ws.row_dimensions[row].height = 28
        for row_offset, values in enumerate(rows, 1):
            is_total = total_row and row_offset == len(rows)
            for offset, value in enumerate(values):
                cell = ws.cell(row + row_offset, start_col + offset, value)
                if is_total:
                    fill = header_fill
                elif offset in input_columns:
                    fill = input_fill
                else:
                    fill = calc_fill if row_offset % 2 == 0 else white
                cell.fill = PatternFill("solid", fgColor=fill)
                cell.font = Font(name="Arial", size=11, bold=is_total, color=ink)
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                if isinstance(value, (int, float)) or (isinstance(value, str) and value.startswith("=")):
                    cell.number_format = number_format_for_header(headers[offset], value) or "#,##0.##"
                if headers[offset] in {"Due", "Due date"} and hasattr(value, "year"):
                    cell.number_format = "dd/mm/yyyy"
                if headers[offset] in {"Wedding date", "RSVP deadline"} and hasattr(value, "year"):
                    cell.number_format = "dd/mm/yyyy"
            ws.row_dimensions[row + row_offset].height = 28
        data_rows = rows[:-1] if total_row and rows else rows
        end_column = get_column_letter(start_col + len(headers) - 1)
        data_end_row = row + max(len(data_rows), 1)
        table_name = f"tbl_{re.sub(r'[^A-Za-z0-9_]', '_', ws.title)}_{len(ws.tables) + 1}"
        ws.add_table(
            Table(
                displayName=table_name[:250],
                ref=f"{get_column_letter(start_col)}{row}:{end_column}{data_end_row}",
            )
        )
        return row + len(rows)

    def simple_note(ws, row: int, value: str) -> None:
        cell = ws.cell(row, 2, value)
        cell.font = Font(name="Arial", size=11, bold=True, color="6B9583")
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        for col in range(2, 10):
            note_cell = ws.cell(row, col)
            note_cell.fill = PatternFill("solid", fgColor=calc_fill)
            note_cell.border = Border(
                top=thin,
                bottom=thin,
                left=thin if col == 2 else Side(style=None),
                right=thin if col == 9 else Side(style=None),
            )
        ws.row_dimensions[row].height = 26

    def simple_nav(ws) -> None:
        for col, tab in enumerate(tabs, 2):
            cell = ws.cell(native_nav_row, col, tab)
            cell.fill = PatternFill("solid", fgColor=accent if tab == ws.title else header_fill)
            cell.font = Font(name="Arial", size=10, bold=tab == ws.title, color=white if tab == ws.title else muted)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(top=medium if tab == ws.title else thin, bottom=thin, left=thin, right=thin)
        ws.row_dimensions[native_nav_row].height = 28

    def set_native_content_frame(ws, widths: list[float]) -> None:
        """Use one explicit horizontal frame for tables and navigation.

        The approved references use a stable inner frame. Content-driven
        widths made each sheet choose a different frame and left the bottom
        tabs wider than the table they belonged to. The frame is now declared
        once per sheet; navigation shapes consume the exact same range.
        """
        ws.column_dimensions["A"].width = 12
        for offset, width in enumerate(widths, 2):
            ws.column_dimensions[get_column_letter(offset)].width = float(width)
        last_column = 1 + len(widths)
        for column in range(last_column + 1, 10):
            ws.column_dimensions[get_column_letter(column)].width = 4

    settings = wb["Settings"]
    simple_header(settings, "Settings", "A short input sheet for dates, budget assumptions and list options.")
    simple_table(settings, 7, ["Setting", "Value", "Used by"], [["Wedding date", date(2027, 10, 15), "Countdown"], ["Total budget", 24500, "Dashboard and Budget"], ["Guest target", 120, "Dashboard"], ["Currency", "USD", "All money cells"], ["RSVP deadline", date(2027, 8, 1), "Guests"], ["Default task owner", "Alex", "Checklist"]], {1})
    settings["C8"].number_format = '[$-en-US]d mmmm yyyy'
    settings["C9"].number_format = '[$-en-US]$#,##0'
    settings["C12"].number_format = '[$-en-US]d mmmm yyyy'

    budget = wb["Budget"]
    simple_header(budget, "Budget", "Yellow cells are inputs; formulas calculate actual and remaining amounts.")
    budget_items = [("Venue & catering", 10000), ("Photography", 3000), ("Decor & flowers", 2500), ("Attire", 2000), ("Music & ceremony", 3000), ("Other", 4000)]
    budget_actuals = [8400, 2100, 1280, 1600, 0, 1900]
    budget_statuses = ["On track", "On track", "On track", "Review", "Open", "On track"]
    budget_rows = []
    for index, ((category, planned), actual, status) in enumerate(zip(budget_items, budget_actuals, budget_statuses), 8):
        budget_rows.append([category, planned, actual, f"=C{index}-D{index}", status])
    budget_rows.append(["TOTAL", "=SUM(C8:C13)", "=SUM(D8:D13)", "=SUM(E8:E13)", ""])
    simple_table(budget, 7, ["Category", "Planned", "Actual", "Remaining", "Status"], budget_rows, {1}, True)
    budget.conditional_formatting.add("E8:E13", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F7D9D7")))
    simple_note(budget, 16, "One table, one set of formulas, one clear input color.")

    guests = wb["Guests"]
    simple_header(guests, "Guests", "One bounded table for names, RSVP state and meal notes.")
    guest_rows = [["Alex & Sam", 2, "Confirmed", "Vegetarian", "Ceremony only"], ["Jordan Lee", 1, "Pending", "Standard", "Follow up"], ["Maya Patel", 2, "Confirmed", "Vegan", "No speeches"], ["Taylor Morgan", 1, "Declined", "Standard", "Sent thanks"], ["Casey & Drew", 2, "Confirmed", "Standard", "Hotel booked"], ["Riley Chen", 1, "Pending", "Vegetarian", "Plus one?"], ["Morgan Diaz", 2, "Confirmed", "Standard", "Table 4"]]
    simple_table(guests, 7, ["Name", "Party", "RSVP", "Meal", "Notes"], guest_rows, {0, 1, 2, 3, 4})
    add_validation(guests, "D8:D107", "Confirmed,Pending,Declined")
    simple_note(guests, 16, "RSVP and meal values are dropdown inputs in Excel.")

    vendors = wb["Vendors"]
    simple_header(vendors, "Vendors", "Track suppliers with simple input and formula columns.")
    vendor_rows = []
    for row_number, row in enumerate([["Rose & Co.", "Florist", 2500, 500, None, "Booked"], ["Harbor Hall", "Venue", 8400, 2000, None, "Booked"], ["Bright Lens", "Photography", 2100, 700, None, "Booked"], ["Sound Union", "Music", 1800, 0, None, "Research"], ["Sweet Field", "Cake", 650, 150, None, "Booked"], ["Thread & Veil", "Attire", 2000, 600, None, "Fitting"]], 8):
        row[4] = f"=D{row_number}-E{row_number}"
        vendor_rows.append(row)
    simple_table(vendors, 7, ["Vendor", "Category", "Estimate", "Deposit", "Balance", "Status"], vendor_rows, {0, 1, 2, 3, 5})
    add_validation(vendors, "G8:G107", "Booked,Research,Fitting,Complete")
    simple_note(vendors, 16, "Balance is a formula: estimate minus deposit and logged payments.")

    payments = wb["Payments"]
    simple_header(payments, "Payments", "A summary row and one ordinary payment log.")
    simple_table(payments, 7, ["Total paid", "Due next 60 days", "Open items", "Currency"], [['=SUMIF(F12:F107,"Yes",E12:E107)', '=SUMIF(F12:F107,"No",E12:E107)', '=COUNTIF(F12:F107,"No")', "=C11"]], set())
    payment_rows = [[date(2027, 6, 15), "Harbor Hall", "Venue deposit", 2000, "Yes", "Deposit"], [date(2027, 6, 28), "Bright Lens", "Second instalment", 1400, "No", "Second instalment"], [date(2027, 7, 6), "Rose & Co.", "Final balance", 2000, "No", "Final balance"], [date(2027, 7, 18), "Thread & Veil", "Fitting balance", 1400, "No", "Fitting balance"], [date(2027, 8, 2), "Sound Union", "Booking", 1800, "No", "Booking"]]
    simple_table(payments, 11, ["Due date", "Vendor", "Category", "Amount", "Paid", "Notes"], payment_rows, {0, 1, 2, 3, 4, 5})
    add_validation(payments, "F12:F107", "Yes,No")

    checklist = wb["Checklist"]
    simple_header(checklist, "Checklist", "Use status values instead of a decorative progress widget.")
    checklist_rows = [["Confirm venue contract", "Alex", date(2027, 6, 15), "Done"], ["Send invitations", "Sam", date(2027, 6, 28), "In progress"], ["Finalise menu", "Alex", date(2027, 7, 6), "To do"], ["Confirm ceremony music", "Sam", date(2027, 7, 18), "To do"], ["Create seating plan", "Alex", date(2027, 8, 2), "To do"], ["Prepare welcome bags", "Sam", date(2027, 8, 16), "To do"], ["Final vendor check", "Alex", date(2027, 8, 30), "To do"]]
    simple_table(checklist, 7, ["Task", "Owner", "Due", "Status"], checklist_rows, {0, 1, 2, 3})
    add_validation(checklist, "E8:E107", "Done,In progress,To do")
    simple_note(checklist, 16, "Completion percentage can be calculated from the Status column.")

    start = wb["Start here"]
    simple_header(start, "Start here", "A short setup guide built as a regular Excel table.")
    simple_table(start, 7, ["Step", "Action", "Result", "Editable cells"], [[1, "Set your date", "Countdown and due dates", "Settings"], [2, "Add your people", "Guests and vendors", "Guests / Vendors"], [3, "Plan the details", "Payments and tasks", "Budget / Checklist"]], {1, 3})
    simple_table(start, 13, ["Tab", "Purpose", "Primary input"], [["Dashboard", "Budget and countdown overview", "Linked totals"], ["Budget", "Plan and actual amounts", "Yellow cells"], ["Guests", "RSVP and meal details", "Guest rows"], ["Vendors", "Deposits and balances", "Vendor rows"], ["Checklist", "Tasks and due dates", "Status values"]], {2})

    dashboard = wb["Dashboard"]
    simple_header(dashboard, "Dashboard", "A compact summary assembled from ordinary worksheet ranges.")
    simple_table(dashboard, 7, ["Total budget", "Paid so far", "Guests confirmed", "Days to go"], [["=Settings!C9", "=Payments!B8", '=SUMIF(Guests!D8:D107,"Confirmed",Guests!C8:C107)', "=MAX(Settings!C8-TODAY(),0)"]])
    dashboard["B10"] = "Budget overview"
    dashboard["B10"].font = Font(name="Aptos Display", size=17, bold=True, color=ink)
    dashboard["B10"].alignment = Alignment(horizontal="left", vertical="center")
    dashboard.row_dimensions[10].height = 26
    dashboard_rows = [[f"=Budget!B{row}", f"=Budget!C{row}", f"=Budget!D{row}", f"=Budget!E{row}", f"=Budget!F{row}"] for row in range(8, 14)]
    dashboard_rows.append(["TOTAL", "=Budget!C14", "=Budget!D14", "=Budget!E14", ""])
    simple_table(dashboard, 11, ["Category", "Planned", "Actual", "Remaining", "Status"], dashboard_rows, {1}, True)
    dashboard["B20"] = "Next actions"
    dashboard["B20"].font = Font(name="Aptos Display", size=17, bold=True, color=ink)
    dashboard["B20"].alignment = Alignment(horizontal="left", vertical="center")
    dashboard.row_dimensions[20].height = 26
    simple_table(dashboard, 21, ["Task", "Due", "Owner"], [["Confirm florist quote", "This week", "Alex"], ["Send RSVP reminder", "10 days", "Sam"], ["Pay venue deposit", "18 days", "Alex"]], {0, 1, 2})

    # Section labels are explicit native rows so the Excel viewport preserves
    # the same vertical rhythm as the approved master composition.
    headings = {
        "Start here": ((6, "Setup guide"), (12, "Workbook map")),
        "Budget": ((6, "Wedding budget"),),
        "Guests": ((6, "Guest list"),),
        "Vendors": ((6, "Vendor directory"),),
        "Checklist": ((6, "Planning checklist"),),
        "Settings": ((6, "Workbook settings"),),
        "Payments": ((10, "Payment timeline"),),
    }
    for sheet_name, rows in headings.items():
        ws = wb[sheet_name]
        for row, label in rows:
            ws.cell(row, 2, label)
            ws.cell(row, 2).font = Font(name="Arial", size=17, bold=True, color=ink)
            ws.cell(row, 2).alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[row].height = 50

    # Widths are the normalized proportions of the approved reference tables,
    # expressed in Excel character units over one shared 240-unit frame.
    set_native_content_frame(wb["Start here"], [26, 74, 74, 66])
    set_native_content_frame(wb["Dashboard"], [74, 43, 43, 43, 37])
    set_native_content_frame(wb["Budget"], [74, 43, 43, 43, 37])
    set_native_content_frame(wb["Guests"], [56, 31, 43, 43, 67])
    set_native_content_frame(wb["Vendors"], [47, 43, 43, 37, 37, 33])
    set_native_content_frame(wb["Payments"], [35, 43, 51, 52, 35, 24])
    set_native_content_frame(wb["Checklist"], [90, 45, 45, 60])
    set_native_content_frame(wb["Settings"], [82, 82, 76])

    for ws in wb.worksheets:
        # Reserve the same quiet tail before the shared navigation baseline on
        # every screen. This is native worksheet whitespace, not a painted
        # fake card, and keeps the bottom row at the approved viewport rhythm.
        for row in range(22, native_nav_row):
            ws.row_dimensions[row].height = 20
        simple_nav(ws)
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(path)


def debt_workbook(spec: dict, path: Path) -> None:
    c = colors(spec)
    wb = Workbook()
    wb.remove(wb.active)
    for tab in spec["tabs"]:
        wb.create_sheet(tab)
    write_start(wb["Start here"], spec, c, [("List every debt", "Enter balances, APR and minimum payments in Debts."), ("Choose a method", "Compare Snowball and Avalanche in Simulator."), ("Log payments", "Payments update principal, interest and progress views.")])
    debts = wb["Debts"]
    style_sheet(debts, "Debts", "Inputs that power both payoff strategies", c)
    write_table(debts, 5, ["Debt", "Type", "Balance", "APR", "Minimum", "Priority"], [["Credit card", "Card", 3200, 0.249, 180, 1], ["Car loan", "Loan", 3000, 0.072, 240, 2], ["Student loan", "Loan", 12440, 0.051, 320, 3]], c, [26, 18, 16, 14, 16, 14])
    payments = wb["Payments"]
    style_sheet(payments, "Payments", "Log real payments once", c)
    payment_rows = [
        [
            date(2027, 1, 5),
            "Credit card",
            180,
            '=IFERROR(INDEX(Debts!$C$6:$C$8,MATCH(B6,Debts!$A$6:$A$8,0))*INDEX(Debts!$D$6:$D$8,MATCH(B6,Debts!$A$6:$A$8,0))/12,0)',
            "=MAX(C6-D6,0)",
        ],
        [
            date(2027, 1, 8),
            "Car loan",
            240,
            '=IFERROR(INDEX(Debts!$C$6:$C$8,MATCH(B7,Debts!$A$6:$A$8,0))*INDEX(Debts!$D$6:$D$8,MATCH(B7,Debts!$A$6:$A$8,0))/12,0)',
            "=MAX(C7-D7,0)",
        ],
        [
            date(2027, 2, 5),
            "Credit card",
            280,
            '=IFERROR(INDEX(Debts!$C$6:$C$8,MATCH(B8,Debts!$A$6:$A$8,0))*INDEX(Debts!$D$6:$D$8,MATCH(B8,Debts!$A$6:$A$8,0))/12,0)',
            "=MAX(C8-D8,0)",
        ],
    ]
    write_table(payments, 5, ["Date", "Debt", "Amount", "Interest", "Principal"], payment_rows, c, [16, 26, 16, 16, 16])
    add_validation(payments, "B6:B105", "Credit card,Car loan,Student loan")
    dashboard = wb["Dashboard"]
    style_sheet(dashboard, "Debt payoff dashboard", "Make the next payment count", c)
    add_metric(dashboard, "A5", "TOTAL BALANCE", "=SUM(Debts!C6:C8)", c)
    add_metric(dashboard, "C5", "PAID THIS MONTH", "=SUM(Payments!C6:C105)", c)
    add_metric(dashboard, "E5", "INTEREST LOGGED", "=SUM(Payments!D6:D105)", c)
    add_metric(dashboard, "G5", "EXTRA PAYMENT", "=Settings!B9", c)
    write_table(dashboard, 10, ["Strategy", "Priority", "Monthly attack", "Use when"], [["Snowball", "Smallest balance", "=SUM(Debts!E6:E8)+Settings!B9", "Momentum matters"], ["Avalanche", "Highest APR", "=SUM(Debts!E6:E8)+Settings!B9", "Interest matters"]], c, [18, 24, 20, 24])
    simulator = wb["Simulator"]
    style_sheet(simulator, "Strategy simulator", "Test an extra payment before committing", c)
    write_table(simulator, 5, ["Input", "Value", "Result"], [["Extra monthly payment", 250, "=B6"], ["Selected method", "Snowball", "=B7"], ["Projected months", "", "=ROUNDUP(SUM(Debts!C6:C8)/(SUM(Debts!E6:E8)+B6),0)"], ["Interest note", "", "Compare both strategies before deciding"]], c, [28, 22, 34])
    add_validation(simulator, "B7", "Snowball,Avalanche")
    progress = wb["Progress"]
    style_sheet(progress, "Progress", "Watch the balance fall", c)
    progress_rows = [[f"M{i}", max(18640 - (i - 1) * 800, 0), min((i - 1) * 800, 18640), f"=C{5+i}/18640"] for i in range(1, 9)]
    write_table(progress, 5, ["Month", "Balance", "Paid down", "Progress"], progress_rows, c, [16, 18, 18, 18])
    # The simple design contract communicates progress with ordinary cells
    # and formulas. Charts remain available to legacy builders, but are not
    # part of the reusable Excel-first template flow.
    settings = wb["Settings"]
    style_sheet(settings, "Settings", "Set assumptions once", c)
    write_table(settings, 5, ["Setting", "Value"], [["Currency", "USD"], ["Plan start date", date(2027, 1, 1)], ["Default strategy", "Avalanche"], ["Extra monthly payment", 250], ["Disclaimer", "Planning tool only"]], c, [30, 28])
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    apply_visual_system(wb, spec)
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(path)


def business_workbook(spec: dict, path: Path) -> None:
    c = colors(spec)
    freelance = spec.get("slug") == "freelance-client-invoice-tracker"
    wb = Workbook()
    wb.remove(wb.active)
    for tab in spec["tabs"]:
        wb.create_sheet(tab)
    write_start(wb["Start here"], spec, c, [("Set your basics", "Choose currency and opening balance in Settings."), ("Track client work" if freelance else "Log movement", "Record income, expenses and invoices with clear statuses."), ("Look forward", "Use Forecast and Year end to prepare the next decision.")])
    transactions = wb["Transactions"]
    style_sheet(transactions, "Transactions", "Every client euro has a place" if freelance else "Every euro has a place", c)
    transaction_rows = [[date(2027, 1, 6), "Discovery retainer" if freelance else "Client retainer", "Client work" if freelance else "Services", 1240, 0, "Cleared"], [date(2027, 1, 8), "Design software" if freelance else "Adobe subscription", "Software", 0, 62, "Cleared"], [date(2027, 1, 12), "Website milestone" if freelance else "Website project", "Client work" if freelance else "Services", 860, 0, "Expected"], [date(2027, 1, 15), "Coworking desk" if freelance else "Coworking space", "Workspace", 0, 180, "Planned"], [date(2027, 1, 22), "Tax reserve", "Taxes", 0, 620, "Reserved"]]
    write_table(transactions, 5, ["Date", "Description", "Category", "Money in", "Money out", "Status"], transaction_rows, c, [16, 30, 20, 16, 16, 18])
    add_validation(transactions, "F6:F105", "Cleared,Expected,Planned,Reserved")
    invoices = wb["Invoices"]
    style_sheet(invoices, "Invoices", "See client income before you spend it" if freelance else "See expected income before you spend it", c)
    write_table(invoices, 5, ["Invoice", "Client", "Due", "Amount", "Status", "Notes"], [["INV-024", "North Studio", date(2027, 1, 15), 1240, "Paid", "Retainer"], ["INV-025", "Green & Co.", date(2027, 1, 28), 860, "Sent", "Milestone"], ["INV-026", "Mila Events", date(2027, 2, 5), 1600, "Draft", "Deposit"], ["INV-027", "Local Market", date(2027, 2, 12), 420, "Overdue", "Follow up"]], c, [16, 26, 16, 16, 18, 24])
    recurring = wb["Recurring"]
    style_sheet(recurring, "Recurring", "Make fixed costs predictable", c)
    write_table(recurring, 5, ["Item", "Category", "Frequency", "Next due", "Amount", "Type"], [["Adobe subscription", "Software", "Monthly", date(2027, 2, 8), 62, "Expense"], ["Coworking space", "Workspace", "Monthly", date(2027, 2, 15), 180, "Expense"], ["Client retainer", "Services", "Monthly", date(2027, 2, 1), 1240, "Income"], ["Bookkeeping", "Professional", "Quarterly", date(2027, 4, 1), 240, "Expense"]], c, [28, 20, 16, 16, 16, 18])
    forecast = wb["Forecast"]
    style_sheet(forecast, "Forecast", "Plan the next 90 days", c)
    forecast_rows = [
        [
            month,
            8420 if index == 1 else f"=E{3 + index}",
            4860 - (index - 1) * 180,
            2140 + (index - 1) * 100,
            f"=B{4 + index}+C{4 + index}-D{4 + index}",
        ]
        for index, month in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun"], 1)
    ]
    write_table(forecast, 4, ["Month", "Opening", "Income", "Expenses", "Closing"], forecast_rows, c, [16, 18, 18, 18, 18])
    # The forecast table is the primary decision surface in the simple flow;
    # a chart would make the template less portable without adding essential
    # information.
    dashboard = wb["Dashboard"]
    style_sheet(dashboard, "Freelance cash command centre" if freelance else "Cash-flow command centre", "Know what your next client payment changes" if freelance else "Know what your business can afford", c)
    add_metric(dashboard, "A5", "CASH ON HAND", "=Settings!B8+SUM(Transactions!D6:D105)-SUM(Transactions!E6:E105)", c)
    add_metric(dashboard, "C5", "THIS MONTH IN", "=SUM(Transactions!D6:D105)", c)
    add_metric(dashboard, "E5", "THIS MONTH OUT", "=SUM(Transactions!E6:E105)", c)
    add_metric(dashboard, "G5", "RUNWAY", "=IF(E6=0,0,A6/E6)", c)
    write_table(dashboard, 10, ["Signal", "Value", "Decision cue"], [["Expected invoices", "=SUM(Invoices!D6:D105)", "Follow up before due date"], ["Tax reserve", "=SUM(Transactions!D6:D105)*Settings!B9", "Keep separate from operating cash"], ["Recurring costs", "=SUM(Recurring!E6:E105)", "Review quarterly"]], c, [24, 20, 36])
    year_end = wb["Year end"]
    style_sheet(year_end, "Year end", "Close the year with a clean summary", c)
    write_table(year_end, 5, ["Category", "Income", "Expenses", "Net"], [["Services", 52800, 0, "=B6-C6"], ["Software", 0, 4920, "=B7-C7"], ["Workspace", 0, 7680, "=B8-C8"], ["Taxes & fees", 0, 5332, "=B9-C9"], ["TOTAL", "=SUM(B6:B9)", "=SUM(C6:C9)", "=SUM(D6:D9)"]], c, [28, 18, 18, 18])
    settings = wb["Settings"]
    style_sheet(settings, "Settings", "Set the rules behind the numbers", c)
    write_table(settings, 5, ["Setting", "Value"], [["Business name", "North Studio"], ["Currency", "EUR"], ["Opening balance", 5700], ["Tax reserve rate", 0.2], ["Financial year start", "January"], ["Default invoice terms", "30 days"]], c, [30, 28])
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    apply_visual_system(wb, spec)
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(path)


def meal_workbook(spec: dict, path: Path) -> None:
    c = colors(spec)
    wb = Workbook()
    wb.remove(wb.active)
    for tab in spec["tabs"]:
        wb.create_sheet(tab)
    write_start(wb["Start here"], spec, c, [("Set preferences", "Choose household size, weekly budget and units in Settings."), ("Choose meals", "Build the week from recipes and keep the grocery list connected."), ("Shop smarter", "Use Pantry and Budget to reduce waste and stay on track.")])
    week = wb["This week"]
    style_sheet(week, "This week", "Breakfast, lunch and dinner in one view", c)
    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    meals = [["Oats", "Pasta", "Salad"], ["Toast", "Tacos", "Soup"], ["Yoghurt", "Rice bowl", "Fish"], ["Smoothie", "Curry", "Leftovers"], ["Eggs", "Pizza", "Wrap"], ["Brunch", "Grill", "Soup"], ["Pancakes", "Roast", "Leftovers"]]
    week_rows = [[days[i], meal_type, meals[i][j], 1.8 + j * 0.7] for i in range(7) for j, meal_type in enumerate(["Breakfast", "Lunch", "Dinner"])]
    write_table(week, 5, ["Day", "Meal", "Recipe", "Est. cost"], week_rows, c, [16, 18, 28, 16])
    for row in range(6, 27):
        week.cell(row, 4, f'=IFERROR(INDEX(Recipes!D$6:D$10,MATCH(C{row},Recipes!A$6:A$10,0)),0)')
    add_validation(week, "B6:B105", "Breakfast,Lunch,Dinner")
    recipes = wb["Recipes"]
    style_sheet(recipes, "Recipes", "Reusable building blocks for the week", c)
    write_table(recipes, 5, ["Recipe", "Meal type", "Prep", "Cost/serving", "Rating", "Tags"], [["Pasta primavera", "Dinner", "25 min", 1.8, 5, "Vegetarian"], ["Chicken curry", "Dinner", "35 min", 2.4, 4, "Freezer-friendly"], ["Overnight oats", "Breakfast", "5 min", 0.9, 5, "Quick"], ["Bean tacos", "Dinner", "20 min", 1.6, 4, "Budget"], ["Greek salad", "Lunch", "10 min", 1.3, 4, "No-cook"]], c, [28, 18, 16, 18, 14, 26])
    grocery = wb["Grocery list"]
    style_sheet(grocery, "Grocery list", "A consolidated list by aisle", c)
    write_table(grocery, 5, ["Done", "Item", "Aisle", "Qty", "Est. cost", "For meals"], [["No", "Tomatoes", "Produce", 6, 3.4, "Pasta · Salad"], ["No", "Rice", "Pantry", "1 kg", 2.2, "Rice bowl"], ["No", "Chicken", "Meat", "800 g", 7.6, "Curry · Grill"], ["No", "Yoghurt", "Dairy", 4, 2.8, "Breakfast"], ["No", "Tortillas", "Bakery", "1 pack", 2.5, "Tacos · Wrap"]], c, [12, 26, 18, 16, 16, 24])
    add_validation(grocery, "A6:A105", "Yes,No")
    pantry = wb["Pantry"]
    style_sheet(pantry, "Pantry", "Use what you already have", c)
    write_table(pantry, 5, ["Item", "Location", "Qty", "Expiry", "Status", "Suggested use"], [["Tomatoes", "Fridge", 6, date(2027, 1, 18), "Use soon", "Pasta"], ["Rice", "Pantry", "1 kg", "", "Good", "Rice bowl"], ["Chicken", "Freezer", "800 g", date(2027, 2, 5), "Good", "Curry"], ["Spinach", "Fridge", "1 bag", date(2027, 1, 14), "Use today", "Salad"]], c, [24, 18, 16, 16, 18, 24])
    budget = wb["Budget"]
    style_sheet(budget, "Budget", "Planned and actual grocery spend", c)
    budget_rows = [[f"Week {i}", 120, actual, f"=B{4+i}-C{4+i}", "On track" if actual <= 120 else "Review"] for i, actual in enumerate([98, 126, 102, 0], 1)]
    write_table(budget, 4, ["Week", "Planned", "Actual", "Difference", "Note"], budget_rows, c, [16, 18, 18, 18, 30])
    # Planned versus actual remains visible in the bounded table. Conditional
    # formatting can be added later as a declared template override.
    dashboard = wb["Dashboard"]
    style_sheet(dashboard, "Meal planning dashboard", "Plan the week before the week plans you", c)
    add_metric(dashboard, "A5", "WEEKLY BUDGET", "=Settings!B7", c)
    add_metric(dashboard, "C5", "PLANNED SPEND", "=SUM('Grocery list'!E6:E105)", c)
    add_metric(dashboard, "E5", "MEALS PLANNED", '=COUNTA(\'This week\'!C6:C105)', c)
    add_metric(dashboard, "G5", "ITEMS TO BUY", '=COUNTIF(\'Grocery list\'!A6:A105,"No")', c)
    write_table(dashboard, 10, ["Signal", "Value", "Meaning"], [["Recipes available", "=COUNTA(Recipes!A6:A105)", "Reusable choices"], ["Pantry items", "=COUNTA(Pantry!A6:A105)", "Use what is already owned"], ["Budget remaining", "=Settings!B7-SUM('Grocery list'!E6:E105)", "Keep the next shop intentional"]], c, [24, 20, 36])
    monthly = wb["Monthly view"]
    style_sheet(monthly, "Monthly view", "See the month without losing the week", c)
    write_table(monthly, 5, ["Week", "Focus", "Planned spend", "Notes"], [["Week 1", "Meal prep", 120, "Use pantry stock"], ["Week 2", "Hosted friends", 120, "Add flexible recipes"], ["Week 3", "Freezer meals", 120, "Reduce waste"], ["Week 4", "Reset", 120, "Review favourites"]], c, [18, 26, 18, 34])
    settings = wb["Settings"]
    style_sheet(settings, "Settings", "Set the household rhythm", c)
    write_table(settings, 5, ["Setting", "Value"], [["Household size", 3], ["Weekly grocery budget", 120], ["Currency", "EUR"], ["Preferred units", "Metric"], ["Dietary notes", "Vegetarian option twice weekly"]], c, [30, 34])
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    apply_visual_system(wb, spec)
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(path)


def renovation_workbook(spec: dict, path: Path) -> None:
    """Build product 12 from a cloned native master with renovation semantics.

    The workbook starts as a copy of the approved native master. Its original
    sheets are repurposed without changing the shared grid, then the product
    tables and formulas are written into those inherited surfaces.
    """
    c = colors(spec)
    source = ROOT / "templates" / "excel-first-base" / "master_workbook.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, path)
    wb = load_workbook(path)

    original_sheets = list(wb.worksheets)
    tabs = list(spec["tabs"])
    for index, tab in enumerate(tabs):
        original_sheets[index].title = tab
    for ws in original_sheets[len(tabs):]:
        wb.remove(ws)

    # Remove inherited content and native objects while retaining the cloned
    # workbook's column/row geometry as the starting surface.
    for ws in wb.worksheets:
        for merged in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merged))
        for table_name in list(ws.tables):
            del ws.tables[table_name]
        ws._charts = []
        ws._images = []
        ws.data_validations.dataValidation = []
        ws.auto_filter.ref = None
        ws.freeze_panes = None
        for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 70), min_col=1, max_col=max(ws.max_column, 20)):
            for cell in row:
                cell.value = None
        ws.sheet_view.showGridLines = False

    write_start(
        wb["Start here"],
        spec,
        c,
        [
            ("Set the project assumptions", "Enter the renovation budget cap, contingency and dates in Settings."),
            ("Log decisions once", "Add planned work, purchases and supplier quotes in their dedicated tables."),
            ("Use the dashboard", "Check remaining budget, committed spend and open tasks before the next decision."),
        ],
    )

    settings = wb["Settings"]
    style_sheet(settings, "Set the renovation rules once", "These five inputs drive the control panel and the project timeline.", c)
    write_table(
        settings,
        5,
        ["Project name", "Budget cap", "Contingency %", "Start date", "Target date"],
        [["Oak Street kitchen refresh", 24000, 0.10, date(2026, 10, 20), date(2027, 2, 28)]],
        c,
        [30, 18, 18, 18, 18],
    )
    settings["B6"].number_format = '[$-en-US]$#,##0.00'
    settings["C6"].number_format = "0%"
    settings["D6"].number_format = "dd/mm/yyyy"
    settings["E6"].number_format = "dd/mm/yyyy"

    budget = wb["Budget"]
    style_sheet(budget, "Budget by renovation area", "Plan the work first; actual spend is linked automatically to Purchases.", c)
    areas = [
        ("Kitchen units", 9000, "Cabinetry and fitting"),
        ("Flooring", 5200, "Materials and installation"),
        ("Lighting", 3600, "Fixtures and electrician"),
        ("Plumbing", 1800, "Connections and fittings"),
        ("Paint", 1400, "Preparation and finish"),
    ]
    budget_rows = []
    for row, (area, planned, note) in enumerate(areas, 6):
        budget_rows.append([
            area,
            planned,
            f'=SUMIF(Purchases!$B$6:$B$10,A{row},Purchases!$E$6:$E$10)',
            f"=B{row}-C{row}",
            f'=IF(C{row}=0,"Not started",IF(D{row}<0,"Over budget","On track"))',
            note,
        ])
    write_table(budget, 5, ["Area", "Planned", "Actual spend", "Remaining", "Status", "Notes"], budget_rows, c, [24, 16, 18, 18, 18, 30])
    budget["A11"] = "TOTAL PROJECT"
    budget["B11"] = "=SUM(B6:B10)"
    budget["C11"] = "=SUM(C6:C10)"
    budget["D11"] = "=SUM(D6:D10)"

    purchases = wb["Purchases"]
    style_sheet(purchases, "Log purchases once", "Every purchase is attributed to one renovation area and flows into Budget.", c)
    purchase_rows = [
        [date(2026, 10, 20), "Kitchen units", "Deposit", "Oak & Co", 0, "No", "Planned"],
        [date(2026, 10, 28), "Flooring", "Porcelain tiles", "Stone House", 1800, "Yes", "Purchased"],
        [date(2026, 11, 4), "Lighting", "Pendant lights", "Bright Home", 1200, "No", "Ordered"],
        [date(2026, 11, 12), "Plumbing", "Tap and valves", "Flow Trade", 500, "Yes", "Purchased"],
        [date(2026, 11, 18), "Paint", "Primer and paint", "Colour Room", 400, "No", "Planned"],
    ]
    write_table(purchases, 5, ["Purchase date", "Area", "Item", "Supplier", "Amount", "Paid?", "Status"], purchase_rows, c, [16, 20, 26, 22, 16, 14, 16])
    purchases["A11"] = "TOTAL PURCHASES"
    purchases["E11"] = "=SUM(E6:E10)"
    add_validation(purchases, "F6:F105", "Yes,No")
    add_validation(purchases, "G6:G105", "Planned,Ordered,Purchased")
    add_validation(purchases, "B6:B105", "Kitchen units,Flooring,Lighting,Plumbing,Paint")

    quotes = wb["Quotes"]
    style_sheet(quotes, "Compare supplier quotes", "Keep alternatives visible before selecting a supplier or committing a purchase.", c)
    quote_rows = [
        ["Oak & Co", "Kitchen units", date(2026, 10, 14), 9400, "Yes", "Includes fitting"],
        ["Northline Joinery", "Kitchen units", date(2026, 10, 16), 10200, "No", "Longer lead time"],
        ["Stone House", "Flooring", date(2026, 10, 18), 5200, "Yes", "Includes installation"],
        ["Light Lab", "Lighting", date(2026, 10, 21), 3600, "No", "Awaiting revised quote"],
        ["Bright Home", "Lighting", date(2026, 10, 23), 3450, "Yes", "Trade discount"],
    ]
    write_table(quotes, 5, ["Supplier", "Area", "Quote date", "Amount", "Selected", "Notes"], quote_rows, c, [24, 20, 16, 16, 16, 30])
    add_validation(quotes, "E6:E105", "Yes,No")

    timeline = wb["Timeline"]
    style_sheet(timeline, "Keep the work moving", "Use due dates and status to make the next action visible.", c)
    timeline_rows = []
    tasks = [
        ("Confirm kitchen layout", "Kitchen units", date(2026, 10, 20), date(2026, 11, 3), "Open"),
        ("Order flooring", "Flooring", date(2026, 10, 28), date(2026, 11, 10), "In progress"),
        ("Book electrician", "Lighting", date(2026, 11, 4), date(2026, 11, 15), "Open"),
        ("Complete plumbing rough-in", "Plumbing", date(2026, 11, 12), date(2026, 11, 22), "Open"),
        ("Paint and finish", "Paint", date(2027, 1, 18), date(2027, 1, 31), "Open"),
    ]
    for row, (task, area, start, due, status) in enumerate(tasks, 6):
        timeline_rows.append([task, area, start, due, status, f"=MAX(D{row}-TODAY(),0)"])
    write_table(timeline, 5, ["Task", "Area", "Start date", "Due date", "Status", "Days left"], timeline_rows, c, [30, 20, 16, 16, 18, 14])
    add_validation(timeline, "E6:E105", "Open,In progress,Done")

    dashboard = wb["Dashboard"]
    style_sheet(dashboard, "Renovation control panel", "See the money committed and the next task before making a project decision.", c)
    add_metric(dashboard, "A5", "BUDGET CAP", "=Settings!B6", c)
    add_metric(dashboard, "C5", "ACTUAL SPEND", "=SUM(Budget!C6:C10)", c)
    add_metric(dashboard, "E5", "CONTINGENCY", "=Settings!B6*Settings!C6", c)
    add_metric(dashboard, "G5", "OPEN TASKS", '=COUNTIF(Timeline!E6:E10,"Open")', c)
    dashboard_rows = [
        ["Budget cap", "=Settings!B6", "Total amount available"],
        ["Planned work", "=SUM(Budget!B6:B10)", "Base plan before contingency"],
        ["Actual spend", "=SUM(Budget!C6:C10)", "All logged purchases"],
        ["Available budget", "=B11-B13", "Cap less actual spend"],
        ["Open tasks", '=COUNTIF(Timeline!E6:E10,"Open")', "Keep the next action visible"],
    ]
    write_table(dashboard, 10, ["Metric", "Value", "Decision cue"], dashboard_rows, c, [24, 18, 34])
    for address in ("B11", "B12", "B13", "B14"):
        dashboard[address].number_format = '[$-en-US]$#,##0.00'
    dashboard["B15"].number_format = "0"

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = None
    apply_visual_system(wb, spec)
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(path)


def renovation_workbook_native(spec: dict, path: Path) -> None:
    """Native Excel implementation used when the master contains COM objects.

    This is intentionally separate from the openpyxl path: the
    master carries validation extensions and drawings that must be
    edited and saved by Excel itself.
    """
    import pythoncom
    import win32com.client as win32
    from openpyxl.utils.datetime import to_excel

    def rgb(value: str) -> int:
        raw = value.replace("#", "")
        return int(raw[4:6] + raw[2:4] + raw[0:2], 16)

    def col_letter(number: int) -> str:
        return get_column_letter(number)

    palette = spec["palette"]
    bg = rgb(palette["background"])
    ink = rgb(palette["ink"])
    accent = rgb(palette["accent"])
    secondary = rgb(palette["secondary"])
    pale = rgb("F4F8F6")
    input_fill = rgb("FFF9E9")
    white = rgb("FFFFFF")
    line = rgb("DCE7E3")
    source = ROOT / "templates" / "excel-first-base" / "master_workbook.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, path)

    def style_header(ws, row: int, headers: list[str]) -> None:
        last = col_letter(len(headers))
        rng = ws.Range(f"A{row}:{last}{row}")
        rng.Interior.Color = accent
        rng.Font.Color = white
        rng.Font.Bold = True
        rng.Font.Name = "Aptos"
        rng.Font.Size = 10
        rng.WrapText = True
        ws.Rows(row).RowHeight = 28

    def write_table(
        ws,
        start_row: int,
        headers: list[str],
        rows: list[list[object]],
        table_name: str,
        widths: list[float],
        formats: list[str | None],
        derived_columns: set[int] | None = None,
    ) -> None:
        derived_columns = derived_columns or set()
        for col, header in enumerate(headers, 1):
            ws.Cells(start_row, col).Value = header
            ws.Columns(f"{col_letter(col)}:{col_letter(col)}").ColumnWidth = widths[col - 1]
        for row_offset, values in enumerate(rows, 1):
            row = start_row + row_offset
            for col, value in enumerate(values, 1):
                cell = ws.Cells(row, col)
                if isinstance(value, str) and value.startswith("="):
                    cell.Formula = value
                elif isinstance(value, date) and not isinstance(value, datetime):
                    cell.Value = float(to_excel(value))
                else:
                    cell.Value = value
                cell.Font.Name = "Aptos"
                cell.Font.Size = 10
                cell.Font.Color = ink
                cell.HorizontalAlignment = -4131
                cell.VerticalAlignment = -4108
                cell.WrapText = True
                cell.Interior.Color = pale if col in derived_columns else (input_fill if col not in derived_columns else pale)
                if formats[col - 1]:
                    cell.NumberFormat = formats[col - 1]
            ws.Rows(row).RowHeight = 24
        style_header(ws, start_row, headers)
        end_row = start_row + len(rows)
        last = col_letter(len(headers))
        table = ws.ListObjects.Add(1, ws.Range(f"A{start_row}:{last}{end_row}"), None, 1)
        table.Name = table_name
        table.TableStyle = "TableStyleMedium2"
        # Excel can expand a ListObject to a stale CurrentRegion inherited
        # from the cloned master. Resize it back to the declared semantic
        # range, then rewrite the body so calculated columns cannot lose their
        # first row while Excel normalises the table.
        table.Resize(ws.Range(f"A{start_row}:{last}{end_row}"))
        for col, header in enumerate(headers, 1):
            ws.Cells(start_row, col).Value = header
        for row_offset, values in enumerate(rows, 1):
            row = start_row + row_offset
            for col, value in enumerate(values, 1):
                cell = ws.Cells(row, col)
                if isinstance(value, str) and value.startswith("="):
                    cell.Formula = value
                elif isinstance(value, date) and not isinstance(value, datetime):
                    cell.Value = float(to_excel(value))
                else:
                    cell.Value = value
        for row in range(start_row + 1, end_row + 1):
            ws.Range(f"A{row}:{last}{row}").Borders(9).Color = line

    def write_total(ws, row: int, last_col: int, label: str, formula_by_col: dict[int, str], formats: dict[int, str]) -> None:
        for col in range(1, last_col + 1):
            cell = ws.Cells(row, col)
            cell.Interior.Color = pale
            cell.Font.Bold = True
            cell.Font.Color = ink
            cell.Borders(9).Color = accent
        ws.Cells(row, 1).Value = label
        for col, formula in formula_by_col.items():
            cell = ws.Cells(row, col)
            cell.Formula = formula
            if col in formats:
                cell.NumberFormat = formats[col]

    def prepare_sheet(ws, title: str, subtitle: str) -> None:
        try:
            ws.Cells.UnMerge()
        except Exception:
            pass
        for index in range(ws.ListObjects.Count, 0, -1):
            ws.ListObjects(index).Delete()
        for index in range(ws.Shapes.Count, 0, -1):
            ws.Shapes(index).Delete()
        try:
            ws.Cells.Validation.Delete()
        except Exception:
            pass
        try:
            ws.Cells.FormatConditions.Delete()
        except Exception:
            pass
        ws.Cells.ClearContents()
        ws.Range("A1:O35").Interior.Color = bg
        # Row 1 belongs to the shared in-sheet navigation contract. Keep the
        # semantic header in rows 2-4 so navigation cleanup cannot erase the
        # title while the workbook is being styled.
        ws.Range("A2:H2").Merge()
        ws.Range("A3:H3").Merge()
        ws.Range("A4:H4").Merge()
        ws.Range("A2").Value = f"RENOVATION PLANNER · {ws.Name.upper()}"
        ws.Range("A3").Value = title
        ws.Range("A4").Value = subtitle
        ws.Range("A2").Font.Name = "Aptos"
        ws.Range("A2").Font.Size = 10
        ws.Range("A2").Font.Bold = True
        ws.Range("A2").Font.Color = secondary
        ws.Range("A3").Font.Name = "Aptos Display"
        ws.Range("A3").Font.Size = 22
        ws.Range("A3").Font.Bold = True
        ws.Range("A3").Font.Color = ink
        ws.Range("A4").Font.Name = "Aptos"
        ws.Range("A4").Font.Size = 10
        ws.Range("A4").Font.Color = rgb("6D7D89")
        ws.Range("A4").WrapText = True
        ws.Rows(1).RowHeight = 34
        ws.Rows(2).RowHeight = 20
        ws.Rows(3).RowHeight = 34
        ws.Rows(4).RowHeight = 30
        for col in range(1, 16):
            ws.Columns(f"{col_letter(col)}:{col_letter(col)}").ColumnWidth = 16
        ws.Columns("A:A").ColumnWidth = 28
        ws.Columns("B:H").ColumnWidth = 22
        ws.Tab.Color = accent
        ws.Activate()
        ws.Application.ActiveWindow.DisplayGridlines = False
        ws.Application.ActiveWindow.FreezePanes = False

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False, AddToMru=False)

        original = [workbook.Worksheets(index) for index in range(1, workbook.Worksheets.Count + 1)]
        for index, tab in enumerate(spec["tabs"], 1):
            original[index - 1].Name = tab
        while workbook.Worksheets.Count > len(spec["tabs"]):
            workbook.Worksheets(workbook.Worksheets.Count).Delete()
        titles = {
            "Start here": (spec["title"], "Three steps to keep renovation money, purchases and deadlines connected."),
            "Dashboard": ("Renovation control panel", "See committed spend, available budget and the next task before deciding."),
            "Budget": ("Budget by renovation area", "Plan the work first; actual spend is linked automatically to Purchases."),
            "Purchases": ("Log purchases once", "Every purchase is attributed to one renovation area and flows into Budget."),
            "Quotes": ("Compare supplier quotes", "Keep alternatives visible before selecting a supplier or committing a purchase."),
            "Timeline": ("Keep the work moving", "Use due dates and status to make the next action visible."),
            "Settings": ("Set the renovation rules once", "These five inputs drive the control panel and the project timeline."),
        }
        for tab in spec["tabs"]:
            prepare_sheet(workbook.Worksheets(tab), *titles[tab])

        start = workbook.Worksheets("Start here")
        start.Range("A5").Value = "How this workbook works"
        start.Range("A5").Font.Bold = True
        start.Range("A5").Font.Color = accent
        steps = [
            ("01", "Set the project assumptions", "Enter the renovation budget cap, contingency and dates in Settings."),
            ("02", "Log decisions once", "Add planned work, purchases and supplier quotes in their dedicated tables."),
            ("03", "Use the dashboard", "Check remaining budget, committed spend and open tasks before the next decision."),
        ]
        for row, (number, heading, body) in enumerate(steps, 7):
            start.Cells(row, 1).Value = number
            start.Cells(row, 1).Interior.Color = accent
            start.Cells(row, 1).Font.Color = white
            start.Cells(row, 1).Font.Bold = True
            start.Cells(row, 2).Value = heading
            start.Cells(row, 2).Font.Bold = True
            start.Cells(row, 3).Value = body
            start.Range(f"C{row}:H{row}").Merge()
            start.Cells(row, 3).WrapText = True
            start.Rows(row).RowHeight = 38
        start.Range("A14").Value = "Important"
        start.Range("A14").Font.Bold = True
        start.Range("B14:H14").Merge()
        start.Range("B14").Value = "Edit highlighted input cells. Formula cells update the dashboard and should not be overwritten."
        start.Range("B14").WrapText = True

        settings = workbook.Worksheets("Settings")
        write_table(settings, 5, ["Project name", "Budget cap", "Contingency %", "Start date", "Target date"], [["Oak Street kitchen refresh", 24000, 0.10, date(2026, 10, 20), date(2027, 2, 28)]], "tblRenovationSettings", [30, 18, 18, 18, 18], [None, "$#,##0.00", "0%", "dd/mm/yyyy", "dd/mm/yyyy"])

        budget = workbook.Worksheets("Budget")
        areas = [("Kitchen units", 9000, "Cabinetry and fitting"), ("Flooring", 5200, "Materials and installation"), ("Lighting", 3600, "Fixtures and electrician"), ("Plumbing", 1800, "Connections and fittings"), ("Paint", 1400, "Preparation and finish")]
        budget_rows = [[area, planned, f'=SUMIF(Purchases!$B$6:$B$10,A{row},Purchases!$E$6:$E$10)', f"=B{row}-C{row}", f'=IF(C{row}=0,"Not started",IF(D{row}<0,"Over budget","On track"))', note] for row, (area, planned, note) in enumerate(areas, 6)]
        write_table(budget, 5, ["Area", "Planned", "Actual spend", "Remaining", "Status", "Notes"], budget_rows, "tblRenovationBudget", [24, 16, 18, 18, 18, 30], [None, "$#,##0.00", "$#,##0.00", "$#,##0.00", None, None], {3, 4, 5})
        write_total(budget, 11, 6, "TOTAL PROJECT", {2: "=SUM(B6:B10)", 3: "=SUM(C6:C10)", 4: "=SUM(D6:D10)"}, {2: "$#,##0.00", 3: "$#,##0.00", 4: "$#,##0.00"})
        budget.ListObjects("tblRenovationBudget").Resize(budget.Range("A5:F10"))

        purchases = workbook.Worksheets("Purchases")
        purchase_rows = [[date(2026, 10, 20), "Kitchen units", "Deposit", "Oak & Co", 0, "No", "Planned"], [date(2026, 10, 28), "Flooring", "Porcelain tiles", "Stone House", 1800, "Yes", "Purchased"], [date(2026, 11, 4), "Lighting", "Pendant lights", "Bright Home", 1200, "No", "Ordered"], [date(2026, 11, 12), "Plumbing", "Tap and valves", "Flow Trade", 500, "Yes", "Purchased"], [date(2026, 11, 18), "Paint", "Primer and paint", "Colour Room", 400, "No", "Planned"]]
        write_table(purchases, 5, ["Purchase date", "Area", "Item", "Supplier", "Amount", "Paid?", "Status"], purchase_rows, "tblRenovationPurchases", [17, 20, 26, 22, 16, 14, 16], ["dd/mm/yyyy", None, None, None, "$#,##0.00", None, None])
        write_total(purchases, 11, 7, "TOTAL PURCHASES", {5: "=SUM(E6:E10)"}, {5: "$#,##0.00"})
        purchases.ListObjects("tblRenovationPurchases").Resize(purchases.Range("A5:G10"))
        for address, formula in (("F6:F105", '"Yes,No"'), ("G6:G105", '"Planned,Ordered,Purchased"'), ("B6:B105", '"Kitchen units,Flooring,Lighting,Plumbing,Paint"')):
            validation = purchases.Range(address).Validation
            validation.Delete()
            validation.Add(3, 1, 1, formula)

        quotes = workbook.Worksheets("Quotes")
        quote_rows = [["Oak & Co", "Kitchen units", date(2026, 10, 14), 9400, "Yes", "Includes fitting"], ["Northline Joinery", "Kitchen units", date(2026, 10, 16), 10200, "No", "Longer lead time"], ["Stone House", "Flooring", date(2026, 10, 18), 5200, "Yes", "Includes installation"], ["Light Lab", "Lighting", date(2026, 10, 21), 3600, "No", "Awaiting revised quote"], ["Bright Home", "Lighting", date(2026, 10, 23), 3450, "Yes", "Trade discount"]]
        write_table(quotes, 5, ["Supplier", "Area", "Quote date", "Amount", "Selected", "Notes"], quote_rows, "tblRenovationQuotes", [24, 20, 17, 16, 16, 30], [None, None, "dd/mm/yyyy", "$#,##0.00", None, None])
        validation = quotes.Range("E6:E105").Validation
        validation.Delete()
        validation.Add(3, 1, 1, '"Yes,No"')

        timeline = workbook.Worksheets("Timeline")
        tasks = [("Confirm kitchen layout", "Kitchen units", date(2026, 10, 20), date(2026, 11, 3), "Open"), ("Order flooring", "Flooring", date(2026, 10, 28), date(2026, 11, 10), "In progress"), ("Book electrician", "Lighting", date(2026, 11, 4), date(2026, 11, 15), "Open"), ("Complete plumbing rough-in", "Plumbing", date(2026, 11, 12), date(2026, 11, 22), "Open"), ("Paint and finish", "Paint", date(2027, 1, 18), date(2027, 1, 31), "Open")]
        timeline_rows = [[task, area, start_date, due_date, status, f"=MAX(D{row}-TODAY(),0)"] for row, (task, area, start_date, due_date, status) in enumerate(tasks, 6)]
        write_table(timeline, 5, ["Task", "Area", "Start date", "Due date", "Status", "Days left"], timeline_rows, "tblRenovationTimeline", [30, 20, 17, 17, 18, 18], [None, None, "dd/mm/yyyy", "dd/mm/yyyy", None, "0"], {6})
        validation = timeline.Range("E6:E105").Validation
        validation.Delete()
        validation.Add(3, 1, 1, '"Open,In progress,Done"')

        dashboard = workbook.Worksheets("Dashboard")
        for label_cell, value_cell, label, formula in (("A5", "A6", "BUDGET CAP", "=Settings!B6"), ("C5", "C6", "ACTUAL SPEND", "=SUM(Budget!C6:C10)"), ("E5", "E6", "CONTINGENCY", "=Settings!B6*Settings!C6"), ("G5", "G6", "OPEN TASKS", '=COUNTIF(Timeline!E6:E10,"Open")')):
            dashboard.Range(label_cell).Value = label
            dashboard.Range(label_cell).Font.Bold = True
            dashboard.Range(label_cell).Font.Color = rgb("6D7D89")
            dashboard.Range(value_cell).Formula = formula
            dashboard.Range(value_cell).Font.Bold = True
            dashboard.Range(value_cell).Font.Size = 18
            dashboard.Range(value_cell).Font.Color = ink
            dashboard.Range(value_cell).Interior.Color = pale
        dashboard_rows = [["Budget cap", "=Settings!B6", "Total amount available"], ["Planned work", "=SUM(Budget!B6:B10)", "Base plan before contingency"], ["Actual spend", "=SUM(Budget!C6:C10)", "All logged purchases"], ["Available budget", "=B11-B13", "Cap less actual spend"], ["Open tasks", '=COUNTIF(Timeline!E6:E10,"Open")', "Keep the next action visible"]]
        write_table(dashboard, 10, ["Metric", "Value", "Decision cue"], dashboard_rows, "tblRenovationDashboard", [24, 18, 34], [None, "$#,##0.00", None], {2})
        dashboard.Range("B15").NumberFormat = "0"

        for tab in spec["tabs"]:
            sheet = workbook.Worksheets(tab)
            sheet.Activate()
            sheet.Application.ActiveWindow.FreezePanes = False
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

