from __future__ import annotations

"""Embed the family-budget SVG mark using Excel's native drawing engine."""

import argparse
import shutil
from pathlib import Path

import win32com.client as win32


ROOT = Path(__file__).resolve().parent.parent
DEFAULT_INPUT = ROOT / "outputs" / "listings" / "family-budget-sources-base" / "app" / "presupuesto-familiar-sources-base-v2.xlsx"
DEFAULT_OUTPUT = DEFAULT_INPUT
DEFAULT_ICON = ROOT / "assets" / "family-budget-icon.svg"

DEFAULT_LABELS = {
    "Inicio": "Inicio",
    "Resumen": "Resumen",
    "Presupuesto": "Presupuesto",
    "Gastos reales": "Reales",
    "Movimientos": "Movimientos",
    "Flujo de caja": "Flujo",
    "Metas y fondo": "Metas",
    "Deudas": "Deudas",
    "Listas": "Listas",
}
NAVY = "164E63"
TEAL = "0F766E"
MINT = "DDF3EF"
WHITE = "FFFFFF"
BUTTON_WIDTH = 68
BUTTON_HEIGHT = 21
BUTTON_GAP = 4


def excel_rgb(hex_color: str) -> int:
    value = hex_color.lstrip("#")
    red, green, blue = (int(value[index:index + 2], 16) for index in (0, 2, 4))
    return red + (green << 8) + (blue << 16)


def excel_column_letter(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def navigation_items(workbook) -> list[tuple[str, str, str]]:
    items = []
    for sheet in workbook.Worksheets:
        used = sheet.UsedRange
        last_column = used.Column + used.Columns.Count - 1
        label = DEFAULT_LABELS.get(sheet.Name, sheet.Name)
        if len(label) > 13:
            label = label[:12] + "…"
        items.append((sheet.Name, label, excel_column_letter(last_column)))
    return items


def clear_generated_shapes(sheet) -> None:
    for index in range(sheet.Shapes.Count, 0, -1):
        shape = sheet.Shapes.Item(index)
        if shape.Name == "FamilyBudgetIcon" or shape.Name.startswith("Nav_"):
            shape.Delete()


def add_navigation(sheet, items: list[tuple[str, str, str]]) -> None:
    sheet.Rows(1).RowHeight = 26
    start = sheet.Range("B1").Left
    top = sheet.Range("B1").Top + 2
    for index, (target, label, _) in enumerate(items):
        left = start + index * (BUTTON_WIDTH + BUTTON_GAP)
        button = sheet.Shapes.AddShape(5, left, top, BUTTON_WIDTH, BUTTON_HEIGHT)
        button.Name = f"Nav_{target.replace(' ', '_')}"
        button.Placement = 1
        button.Fill.ForeColor.RGB = excel_rgb(NAVY if sheet.Name == target else MINT)
        button.Line.Visible = 0
        text = button.TextFrame2
        text.MarginLeft = 2
        text.MarginRight = 2
        text.MarginTop = 0
        text.MarginBottom = 0
        text.VerticalAnchor = 3
        text.TextRange.Text = label
        text.TextRange.ParagraphFormat.Alignment = 2
        text.TextRange.Font.Name = "Aptos"
        text.TextRange.Font.Size = 8
        text.TextRange.Font.Bold = -1
        text.TextRange.Font.Fill.ForeColor.RGB = excel_rgb(WHITE if sheet.Name == target else NAVY)
        sheet.Hyperlinks.Add(button, "", f"'{target}'!A1")


def add_cover_icon(sheet, last_column: str, icon: Path) -> None:
    anchor = sheet.Range(f"{last_column}2")
    left = anchor.Left + max(anchor.Width - 28, 2)
    picture = sheet.Shapes.AddPicture(str(icon), 0, -1, left, anchor.Top + 1, 26, 26)
    picture.Name = "FamilyBudgetIcon"
    picture.Placement = 1


def embed_icon(source: Path, output: Path, icon: Path) -> Path:
    source = source.resolve()
    output = output.resolve()
    icon = icon.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    if source != output:
        shutil.copy2(source, output)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        items = navigation_items(workbook)
        for sheet_name, _, last_column in items:
            sheet = workbook.Worksheets(sheet_name)
            clear_generated_shapes(sheet)
            add_navigation(sheet, items)
            add_cover_icon(sheet, last_column, icon)
        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description="Incrusta el icono SVG de portada en un XLSX nativo")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--icon", type=Path, default=DEFAULT_ICON)
    args = parser.parse_args()
    source = args.input if args.input.is_absolute() else ROOT / args.input
    output = args.output if args.output.is_absolute() else ROOT / args.output
    icon = args.icon if args.icon.is_absolute() else ROOT / args.icon
    print(f"WORKBOOK={embed_icon(source, output, icon)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
