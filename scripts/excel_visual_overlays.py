from __future__ import annotations

import json
import re
from pathlib import Path

import pythoncom
import win32com.client as win32


SHAPE_PREFIX = "FactoryVisual_"
MSO_ROUNDED_RECTANGLE = 5
MSO_RECTANGLE = 1


PRODUCT8_TABLE_WIDTHS = {
    # Proportions inherited from the master workbook pattern. The totals stay
    # close to the shared grid, while semantic label columns receive the room
    # they need instead of making every column equally wide.
    # Values are Excel character-width units from the native workbook grid.
    "Budget": {"C": 70, "D": 41, "E": 31, "F": 35, "G": 47},
    "Guests": {"C": 43.8, "D": 29.6, "E": 36.3, "F": 37.4, "G": 34.2, "H": 40},
    "Vendors": {"C": 46.7, "D": 42, "E": 33.7, "F": 28.2, "G": 32.4, "H": 38},
    "Payments": {"C": 26, "D": 42.2, "E": 44.7, "F": 29.5, "G": 34.4, "H": 44},
    "Checklist": {"C": 74.2, "D": 34.7, "E": 49.3, "F": 29.2, "G": 34},
    "Settings": {"C": 70.6, "D": 66.9, "E": 86},
}


PRODUCT8_TABLE_LAST_COLUMN = {
    "Budget": "G",
    "Guests": "H",
    "Vendors": "H",
    "Payments": "H",
    "Checklist": "G",
    "Settings": "E",
}


def excel_rgb(hex_value: str) -> int:
    value = hex_value.lstrip("#")
    red, green, blue = (int(value[index : index + 2], 16) for index in (0, 2, 4))
    return red + (green << 8) + (blue << 16)


def style_shape(shape, line_rgb: int, transparency: float = 1.0, weight: float = 1.5) -> None:
    shape.Fill.Visible = 0
    shape.Line.Visible = -1
    shape.Line.ForeColor.RGB = line_rgb
    shape.Line.Weight = weight
    shape.Fill.Transparency = transparency


def rounded(
    sheet,
    name: str,
    cell_range,
    line_rgb: int,
    inset: float = 2.0,
    pad_left: float = 0.0,
    pad_top: float = 0.0,
    pad_right: float = 0.0,
    pad_bottom: float = 0.0,
):
    left = float(cell_range.Left) - pad_left + inset
    top = float(cell_range.Top) - pad_top + inset
    width = max(30.0, float(cell_range.Width) + pad_left + pad_right - inset * 2)
    height = max(24.0, float(cell_range.Height) + pad_top + pad_bottom - inset * 2)
    shape = sheet.Shapes.AddShape(MSO_ROUNDED_RECTANGLE, left, top, width, height)
    shape.Name = f"{SHAPE_PREFIX}{name}"
    style_shape(shape, line_rgb)
    # Bring filled accents above cell fills. A send-backward order would put
    # the rectangle behind the worksheet's opaque cell background.
    shape.ZOrder(0)
    return shape


def badge(sheet, text: str, cell_range, line_rgb: int, ink_rgb: int) -> None:
    shape = rounded(sheet, "IdentityBadge", cell_range, line_rgb, 1.5)
    shape.Fill.Visible = -1
    shape.Fill.ForeColor.RGB = 0xFFFFFF
    shape.Fill.Transparency = 0.0
    shape.TextFrame2.TextRange.Text = text
    shape.TextFrame2.TextRange.Font.Size = 10
    shape.TextFrame2.TextRange.Font.Bold = -1
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ink_rgb
    shape.TextFrame2.VerticalAnchor = 3
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2


def nav_button(sheet, name: str, left: float, top: float, width: float, height: float, text: str, active: bool, accent: int, ink: int) -> None:
    """Draw a real pill-shaped navigation button across the worksheet surface."""
    shape = sheet.Shapes.AddShape(MSO_ROUNDED_RECTANGLE, left, top, max(42.0, width), max(22.0, height))
    shape.Name = f"{SHAPE_PREFIX}{name}"
    shape.Fill.Visible = -1
    shape.Fill.ForeColor.RGB = accent if active else excel_rgb("EEF3F0")
    shape.Fill.Transparency = 0.0
    shape.Line.Visible = -1
    shape.Line.ForeColor.RGB = accent if active else excel_rgb("D8E4E0")
    shape.Line.Weight = 0.8
    shape.TextFrame2.TextRange.Text = text
    shape.TextFrame2.TextRange.Font.Size = 10.5
    shape.TextFrame2.TextRange.Font.Bold = -1 if active else 0
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF if active else ink
    shape.TextFrame2.VerticalAnchor = 3
    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    shape.ZOrder(1)


def solid_rectangle(sheet, name: str, left: float, top: float, width: float, height: float, fill_rgb: int) -> None:
    """Add a small native accent block without changing the cell grid."""
    shape = sheet.Shapes.AddShape(MSO_RECTANGLE, left, top, max(2.0, width), max(2.0, height))
    shape.Name = f"{SHAPE_PREFIX}{name}"
    shape.Fill.Visible = -1
    shape.Fill.ForeColor.RGB = fill_rgb
    shape.Fill.Transparency = 0.0
    shape.Line.Visible = 0
    # Bring filled accents above cell fills. A send-backward order would put
    # the rectangle behind the worksheet's opaque cell background.
    shape.ZOrder(0)


def identity_frame(sheet, name: str, left: float, top: float, width: float, height: float, line_rgb: int) -> None:
    """Create the native frame around the immutable SVG artwork."""
    shape = sheet.Shapes.AddShape(MSO_ROUNDED_RECTANGLE, left, top, width, height)
    shape.Name = f"{SHAPE_PREFIX}{name}"
    shape.Fill.Visible = -1
    shape.Fill.ForeColor.RGB = 0xFFFFFF
    shape.Fill.Transparency = 0.0
    shape.Line.Visible = -1
    shape.Line.ForeColor.RGB = line_rgb
    shape.Line.Weight = 1.5
    # The outer frame must remain visible above the worksheet surface; the SVG
    # is inserted immediately afterwards and keeps the top drawing order.
    shape.ZOrder(0)


def svg_badge(sheet, svg_path: Path, cell_range) -> bool:
    """Place the shared product SVG into the workbook when Excel supports it."""
    try:
        left = float(cell_range.Left)
        top = float(cell_range.Top)
        width = float(cell_range.Width)
        height = float(cell_range.Height)
        # Preserve the source viewBox ratio. Excel's AddPicture accepts a
        # target rectangle and otherwise stretches an SVG to that rectangle;
        # the factory must resize the container, never the identity artwork.
        source = svg_path.read_text(encoding="utf-8")
        match = re.search(r"viewBox=['\"]\s*[-+]?\d+(?:\.\d+)?\s+[-+]?\d+(?:\.\d+)?\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*['\"]", source)
        view_width, view_height = (float(match.group(1)), float(match.group(2))) if match else (160.0, 150.0)
        source_ratio = view_width / max(0.01, view_height)
        box_ratio = width / max(0.01, height)
        if box_ratio > source_ratio:
            final_height = height
            final_width = height * source_ratio
            left += (width - final_width) / 2
        else:
            final_width = width
            final_height = width / source_ratio
            top += (height - final_height) / 2
        shape = sheet.Shapes.AddPicture(str(svg_path), -1, -1, left, top, final_width, final_height)
        shape.Name = f"{SHAPE_PREFIX}IdentitySvg"
        shape.LockAspectRatio = -1
        return True
    except Exception:
        return False


def svg_badge_at(sheet, svg_path: Path, left: float, top: float, width: float, height: float) -> bool:
    """Place the product identity at the fixed responsive header coordinates.

    A cell-anchored logo moves with each table's column proportions. The
    approved screen design has one shared header frame, so product-8 uses a
    fixed point frame in the worksheet viewport instead. The SVG itself is
    still contained and its source viewBox/aspect ratio is never changed.
    """
    try:
        source = svg_path.read_text(encoding="utf-8")
        match = re.search(
            r"viewBox=['\"]\s*[-+]?\d+(?:\.\d+)?\s+[-+]?\d+(?:\.\d+)?\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*['\"]",
            source,
        )
        view_width, view_height = (float(match.group(1)), float(match.group(2))) if match else (160.0, 150.0)
        source_ratio = view_width / max(0.01, view_height)
        box_ratio = width / max(0.01, height)
        if box_ratio > source_ratio:
            final_height = height
            final_width = height * source_ratio
            final_left = left + (width - final_width) / 2
            final_top = top
        else:
            final_width = width
            final_height = width / source_ratio
            final_left = left
            final_top = top + (height - final_height) / 2
        shape = sheet.Shapes.AddPicture(str(svg_path), -1, -1, final_left, final_top, final_width, final_height)
        shape.Name = f"{SHAPE_PREFIX}IdentitySvg"
        shape.LockAspectRatio = -1
        return True
    except Exception:
        return False


def apply_simple_native_overlays(book_path: Path, spec: dict) -> None:
    """Add only the native objects required by the global Excel contract.

    The table-first product does not use decorative cards, charts or external
    identity artwork. Its visual identity is carried by the master geometry,
    the declared palette and the workbook's native cell surfaces.
    """
    if spec.get("preserve_native_master_objects"):
        # The master already contains the native navigation and visual frame.
        # Do not create, import or replace an SVG here: future products must
        # derive identity from their five-colour palette and coherent data.
        return
    if (spec.get("navigation_contract") or {}).get("component") == "native_excel_sheet_tabs":
        return
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(str(book_path.resolve()), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        palette = spec["palette"]
        accent = excel_rgb(palette["accent"])
        ink = excel_rgb(palette["ink"])
        background = excel_rgb(palette["background"])
        native_navigation = spec.get("native_excel_design_contract", {}).get("navigation", {})
        nav_row = int(native_navigation.get("row", 32))
        nav_gap = float(native_navigation.get("gap_points", 8.0))
        nav_height = float(native_navigation.get("height_points", 34.0))
        for sheet in workbook.Worksheets:
            for index in range(sheet.Shapes.Count, 0, -1):
                shape = sheet.Shapes(index)
                if str(shape.Name).startswith(SHAPE_PREFIX):
                    shape.Delete()
            for col in range(1, 16):
                cell = sheet.Cells(nav_row, col)
                cell.Value = ""
                cell.Interior.Color = background
                cell.Borders.LineStyle = 0
            sheet.Rows(nav_row).RowHeight = nav_height
            nav_top = float(sheet.Range(f"A{nav_row}:O{nav_row}").Top) + 1.0
            # A navigation row is anchored to the declared content frame. This
            # keeps the buttons aligned with the design tables on every sheet;
            # fixed point widths are only a fallback for legacy contracts.
            range_by_sheet = native_navigation.get("anchor_range_by_sheet", {})
            configured_range = range_by_sheet.get(str(sheet.Name), native_navigation.get("anchor_range", "")) if isinstance(range_by_sheet, dict) else native_navigation.get("anchor_range", "")
            anchor_range = str(configured_range).strip()
            if anchor_range and ":" in anchor_range:
                start_col, end_col = [part.strip().upper() for part in anchor_range.split(":", 1)]
                frame = sheet.Range(f"{start_col}{nav_row}:{end_col}{nav_row}")
                nav_left = float(frame.Left)
                nav_width = float(frame.Width)
            else:
                nav_left = float(native_navigation.get("left_points", 77.0))
                nav_width = float(native_navigation.get("width_points", 1251.0))
            button_width = (nav_width - nav_gap * (len(spec["tabs"]) - 1)) / max(1, len(spec["tabs"]))
            for index, tab in enumerate(spec["tabs"]):
                nav_button(
                    sheet,
                    f"Nav{index + 1:02d}",
                    nav_left + index * (button_width + nav_gap),
                    nav_top,
                    button_width,
                    nav_height,
                    str(tab),
                    str(tab) == str(sheet.Name),
                    accent,
                    ink,
                )
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


def apply_visual_overlays(book_path: Path, spec: dict) -> None:
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(str(book_path.resolve()), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        try:
            workbook.ForceFullCalculation = True
            excel.CalculateFullRebuild()
        except Exception:
            pass
        palette = spec["palette"]
        accent = excel_rgb(palette["accent"])
        ink = excel_rgb(palette["ink"])
        for sheet in workbook.Worksheets:
            if spec.get("slug") == "wedding-planner":
                # Excel may normalise explicit OpenXML widths back to its
                # default 13-character width while saving shape overlays.
                # Restore the same native grid before measuring/placing the
                # rounded cards; otherwise the visual implementation drifts
                # between generation and native object placement.
                if str(sheet.Name) == "Dashboard":
                    widths = {"A": 20, **{column: 26 for column in "BCEFHIKL"}, **{column: 4 for column in "DGJ"}, "M": 18, "N": 4, "O": 4}
                elif str(sheet.Name) == "Start here":
                    widths = {"A": 20, "B": 14, "C": 34, **{column: 30 for column in "DEFGH"}, "I": 22.5, "J": 4, "K": 4, "L": 4, "M": 4}
                else:
                    # Inherit the calibrated table proportions instead of
                    # giving every column the same width. The master pattern
                    # reserves more space for labels/notes and less for
                    # compact numeric/status fields.
                    table_widths = PRODUCT8_TABLE_WIDTHS.get(
                        str(sheet.Name),
                        {column: 44 for column in "CDEFGH"},
                    )
                    widths = {
                        # A/B keep the C table anchor fixed while moving the
                        # header text to the same 5.5vw left safe area as the
                        # responsive reference.
                        "A": 14,
                        "B": 6,
                        **table_widths,
                    "I": 4,
                        "J": 4,
                        "K": 4,
                        "L": 4,
                        "M": 4,
                    }
                for column, width in widths.items():
                    sheet.Columns(f"{column}:{column}").ColumnWidth = width
                # Freeze panes are useful for raw data tables but create a
                # visible horizontal split line through this presentation
                # canvas. The inherited screen pattern deliberately has one
                # clean surface, so the workbook remains scrollable without
                # a visible split artefact.
                sheet.Activate()
                sheet.Application.ActiveWindow.FreezePanes = False
            for index in range(sheet.Shapes.Count, 0, -1):
                shape = sheet.Shapes(index)
                if str(shape.Name).startswith(SHAPE_PREFIX):
                    shape.Delete()
            identity_range = sheet.Range("J1:K3") if str(sheet.Name) in {"Dashboard", "Start here"} else sheet.Range("H1:I3")
            # Identity is deliberately native and palette-led. Do not import
            # an external SVG: the product name/motif and the five declared
            # colours are the reusable identity contract.
            if spec.get("slug") == "wedding-planner":
                identity_frame(sheet, "IdentityFrame", 1196.0, 25.5, 120.0, 120.0, accent)
            badge(sheet, str(spec["visual_identity"]["name"]), identity_range, accent, ink)
            if str(sheet.Name) == "Start here":
                rounded(sheet, "SetupSurface", sheet.Range("B5:I16"), accent, 2.0, pad_left=32.0, pad_top=10.5, pad_right=34.0)
                for row in range(7, 10):
                    rounded(sheet, f"SetupCard{row}", sheet.Range(f"B{row}:I{row}"), accent, 4.5)
            elif str(sheet.Name) == "Dashboard":
                rounded(sheet, "DashboardSurface", sheet.Range("B5:L20"), excel_rgb("E6D8D5"), 2.0, pad_left=32.0, pad_top=18.0, pad_right=24.0, pad_bottom=18.0)
                metric_line = excel_rgb("E8EFEC")
                metric_rails = (accent, excel_rgb(palette.get("secondary", "7A9B89")), excel_rgb(palette.get("highlight", "D3A94C")), excel_rgb("8A83B1"))
                for number, address in enumerate(("B5:C8", "E5:F8", "H5:I8", "K5:L8"), 1):
                    metric_range = sheet.Range(address)
                    rounded(sheet, f"Metric{number}", metric_range, metric_line, 2.0)
                    solid_rectangle(
                        sheet,
                        f"MetricRail{number}",
                        float(metric_range.Left) + 4.0,
                        float(metric_range.Top) + 2.0,
                        10.0,
                        max(20.0, float(metric_range.Height) - 4.0),
                        metric_rails[number - 1],
                    )
                rounded(sheet, "Summary", sheet.Range("B11:F20"), excel_rgb(palette.get("line", "E6D8D5")), 2.0)
                rounded(sheet, "Actions", sheet.Range("H11:L20"), ink, 2.0)
            else:
                # Do not use one universal A5:H25 card: it makes short tables
                # look lost inside a giant rounded rectangle and places the
                # border far away from the semantic content. These bounds are
                # the normalized card extents of the approved screens.
                card_end = spec.get("excel_card_end_rows", {}).get(str(sheet.Name), 15)
                last_column = PRODUCT8_TABLE_LAST_COLUMN.get(str(sheet.Name), "H")
                table_range = sheet.Range(f"C5:{last_column}{card_end}")
                rounded(
                    sheet,
                    "PrimaryTable",
                    table_range,
                    excel_rgb("E6D8D5") if spec.get("slug") == "wedding-planner" else accent,
                    2.0,
                    pad_left=34.0,
                    pad_top=20.0,
                    pad_right=34.0,
                    pad_bottom=20.5,
                )
            native_navigation = spec.get("native_excel_design_contract", {}).get("navigation", {})
            screen_rows = native_navigation.get("screen_rows", {})
            nav_row = int(screen_rows.get(str(sheet.Name), native_navigation.get("row", 24 if spec.get("slug") == "wedding-planner" else 32)))
            nav_range = sheet.Range(f"A{nav_row}:O{nav_row}")
            # Keep the cell row as a quiet semantic base and render the visible
            # navigation as separated pill buttons. This avoids the default
            # Excel-tab appearance of adjacent coloured cells.
            for col in range(1, 16):
                cell = sheet.Cells(nav_row, col)
                cell.Value = ""
                cell.Interior.Color = excel_rgb(spec["palette"]["background"])
                cell.Borders.LineStyle = 0
            sheet.Rows(nav_row).RowHeight = 34 if spec.get("slug") == "wedding-planner" else 30
            if spec.get("slug") == "wedding-planner":
                # These are native point coordinates from the Excel viewport
                # contract, not browser pixels and not a measurement of the
                # current UsedRange. This is the Excel equivalent of the
                # shared navigation contract's left/right anchoring.
                nav_left = float(native_navigation.get("left_points", 77.0))
                nav_top = float(nav_range.Top) + float(native_navigation.get("top_offset_points", {}).get(str(sheet.Name), 1.0))
                nav_width = float(native_navigation.get("width_points", 1251.0))
                gap = float(native_navigation.get("gap_points", 8.0))
            else:
                nav_left = float(nav_range.Left) + 72.0
                nav_top = float(nav_range.Top) + 2.0
                nav_width = max(320.0, float(nav_range.Width) - 144.0)
                gap = 8.0
            button_width = (nav_width - gap * (len(spec["tabs"]) - 1)) / max(1, len(spec["tabs"]))
            button_height = max(24.0, float(native_navigation.get("height_points", nav_range.Height - 4.0)) if spec.get("slug") == "wedding-planner" else float(nav_range.Height) - 4.0)
            for index, tab in enumerate(spec["tabs"]):
                nav_button(
                    sheet,
                    f"Nav{index + 1:02d}",
                    nav_left + index * (button_width + gap),
                    nav_top,
                    button_width,
                    button_height,
                    str(tab),
                    str(tab) == str(sheet.Name),
                    accent,
                    ink,
                )
        try:
            workbook.ForceFullCalculation = True
            excel.CalculateFullRebuild()
        except Exception:
            pass
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


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--book", type=Path, required=True)
    parser.add_argument("--spec", type=Path, required=True)
    args = parser.parse_args()
    apply_visual_overlays(args.book, json.loads(args.spec.read_text(encoding="utf-8")))
