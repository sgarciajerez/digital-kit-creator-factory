from __future__ import annotations

"""Build a standalone family-budget workbook directly from the source patterns.

This builder deliberately does not use the HTML/design-spec template pipeline.
It creates a native Excel workbook with the recurring structures found in
templates/sources: monthly matrices, plan-vs-real tables, transaction logs,
cash-flow summaries, emergency-fund goals, debt inputs and editable lists.
"""

import argparse
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from model_quality import validate_model_file
from model_contracts import family_base_contract, write_model_behavior_markdown, write_model_contract_json


ROOT = Path(__file__).resolve().parent.parent
SOURCE_DIR = ROOT / "templates" / "sources"
DEFAULT_OUTPUT = ROOT / "outputs" / "listings" / "family-budget-sources-base" / "app" / "presupuesto-familiar-sources-base-final.xlsx"
ICON_PATH = ROOT / "assets" / "family-budget-icon.svg"

MONTHS = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
CATEGORIES = [
    "Vivienda", "Alimentación", "Transporte", "Salud", "Educación", "Ocio", "Otros",
]
TYPES = ["Ingreso", "Gasto", "Ahorro"]
STATES = ["Conciliado", "Pendiente"]
MEMBERS = ["Ana", "Carlos", "Familia"]

NAVY = "164E63"
TEAL = "0F766E"
MINT = "DDF3EF"
CORAL = "477D8E"
AMBER = "5C8D94"
CREAM = "F4F8F7"
INK = "20343D"
MUTED = "60767D"
LINE = "C6D4D8"
INPUT = "E7F2EF"
FORMULA = "E8F0F5"
WHITE = "FFFFFF"
PALE_CORAL = "EDF4F5"

THIN = Side(style="thin", color=LINE)
BORDER = Border(bottom=THIN)
CURRENCY = "#,##0.00"
PERCENT = "0.0%"
DATE_FORMAT = "dd/mm/yyyy"


def solid(color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=color)


def style_cell(cell, *, fill: str | None = None, color: str = INK, bold: bool = False,
               size: int = 10, align: str = "left", wrap: bool = False,
               number_format: str | None = None) -> None:
    if fill:
        cell.fill = solid(fill)
    cell.font = Font(name="Aptos", size=size, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    cell.border = BORDER
    if number_format:
        cell.number_format = number_format


def set_frame(ws, title: str, subtitle: str, last_col: str, widths: dict[str, float], tab_color: str) -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 95
    ws.sheet_properties.tabColor = tab_color
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    for col in range(1, ord(last_col.upper()) - ord("A") + 2):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = widths.get(letter, 12)
    for row in range(1, 40):
        ws.row_dimensions[row].height = 21
    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[4].height = 22
    for row, value, size, color in (
        (2, f"PRESUPUESTO FAMILIAR · {ws.title.upper()}", 10, TEAL),
        (3, title, 23, NAVY),
        (4, subtitle, 10, MUTED),
    ):
        ws.merge_cells(f"B{row}:{last_col}{row}")
        cell = ws[f"B{row}"]
        cell.value = value
        style_cell(cell, color=color, bold=row == 3, size=size, wrap=True)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.print_area = f"B2:{last_col}32"


def section(ws, row: int, start: str, end: str, label: str, fill: str = NAVY) -> None:
    ws.merge_cells(f"{start}{row}:{end}{row}")
    cell = ws[f"{start}{row}"]
    cell.value = label
    style_cell(cell, fill=fill, color=WHITE, bold=True, size=10)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 24


def merge_text(ws, cell_range: str, value: str, *, fill: str, color: str = INK,
               bold: bool = False, size: int = 10, align: str = "left") -> None:
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    cell.value = value
    style_cell(cell, fill=fill, color=color, bold=bold, size=size, align=align, wrap=True)


def header_row(ws, row: int, start_col: int, headers: list[str]) -> None:
    for index, value in enumerate(headers, start_col):
        cell = ws.cell(row, index, value)
        style_cell(cell, fill=NAVY, color=WHITE, bold=True, size=9, wrap=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32


def data_row(ws, row: int, start_col: int, values: list[object], *, inputs: set[int] | None = None,
             formulas: set[int] | None = None, number_formats: dict[int, str] | None = None,
             row_fill: str | None = None) -> None:
    inputs = inputs or set()
    formulas = formulas or set()
    number_formats = number_formats or {}
    for offset, value in enumerate(values):
        col = start_col + offset
        cell = ws.cell(row, col, value)
        is_formula = offset in formulas or (isinstance(value, str) and value.startswith("="))
        fill = FORMULA if is_formula else (INPUT if offset in inputs else (row_fill or WHITE))
        style_cell(
            cell,
            fill=fill,
            color=INK,
            size=10,
            wrap=True,
            number_format=number_formats.get(offset),
        )
        cell.alignment = Alignment(horizontal="left" if offset == 0 else "right", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 22


def add_table(ws, name: str, ref: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)


def add_list_validation(ws, formula: str, cell_range: str) -> None:
    validation = DataValidation(type="list", formula1=formula, allow_blank=True)
    validation.error = "Selecciona un valor de la lista."
    validation.errorTitle = "Valor no válido"
    validation.prompt = "Elige una opción para mantener los resúmenes consistentes."
    validation.promptTitle = "Lista"
    ws.add_data_validation(validation)
    validation.add(cell_range)


def clear_bottom_border(ws, row: int) -> None:
    """Keep the row-7 layout clean; table styles provide their own header edge."""
    for col in range(1, ws.max_column + 1):
        ws.cell(row, col).border = Border()


def card(ws, start_col: int, end_col: int, top: int, bottom: int, label: str, formula: str, fill: str) -> None:
    start = get_column_letter(start_col)
    end = get_column_letter(end_col)
    ws.merge_cells(f"{start}{top}:{end}{top}")
    ws.merge_cells(f"{start}{top + 1}:{end}{bottom}")
    label_cell = ws.cell(top, start_col, label.upper())
    value_cell = ws.cell(top + 1, start_col, formula)
    style_cell(label_cell, fill=fill, color=MUTED, bold=True, size=9)
    style_cell(value_cell, fill=fill, color=NAVY, bold=True, size=19, number_format=CURRENCY)
    label_cell.alignment = Alignment(horizontal="left", vertical="center")
    value_cell.alignment = Alignment(horizontal="left", vertical="center")
    for row in range(top, bottom + 1):
        for col in range(start_col, end_col + 1):
            ws.cell(row, col).border = BORDER
    ws.row_dimensions[top].height = 20
    ws.row_dimensions[top + 1].height = 28
    ws.row_dimensions[bottom].height = 20
    for row in range(top, bottom + 1):
        for col in range(start_col, end_col + 1):
            ws.cell(row, col).border = Border()


def build_inicio(ws, source_count: int) -> None:
    set_frame(
        ws,
        "Una base clara para organizar el dinero del hogar",
        f"Libro nativo construido desde {source_count} fuentes de templates/sources · las celdas amarillas son editables.",
        "J",
        {"A": 3, "B": 18, "C": 20, "D": 20, "E": 20, "F": 3, "G": 21, "H": 16, "I": 16, "J": 16},
        TEAL,
    )
    section(ws, 6, "B", "E", "CÓMO USARLO", TEAL)
    steps = [
        ("01", "Configura el hogar", "Cambia mes, ingresos, ahorro y miembros en Listas."),
        ("02", "Planifica los meses", "Completa la matriz Presupuesto con tus importes."),
        ("03", "Registra cada movimiento", "Movimientos alimenta Gastos reales y Flujo de caja."),
        ("04", "Revisa y decide", "Resumen, Metas y Deudas convierten los datos en acciones."),
    ]
    for row, (number, heading, text) in enumerate(steps, 7):
        ws[f"B{row}"] = number
        style_cell(ws[f"B{row}"], fill=MINT, color=TEAL, bold=True, size=12, align="center")
        merge_text(ws, f"C{row}:D{row}", heading, fill=WHITE, color=NAVY, bold=True)
        merge_text(ws, f"E{row}:E{row}", text, fill=WHITE, color=MUTED, size=9)
        ws.row_dimensions[row].height = 34
    section(ws, 6, "G", "J", "PARÁMETROS RÁPIDOS", CORAL)
    quick = [
        ("Mes seleccionado", "='Listas'!$C$8", CURRENCY),
        ("Ingresos netos", "='Listas'!$C$9", CURRENCY),
        ("Ahorro objetivo", "='Listas'!$C$10", CURRENCY),
        ("Miembro A", "='Listas'!$C$11", None),
        ("Miembro B", "='Listas'!$C$12", None),
    ]
    for row, (label, formula, number_format) in enumerate(quick, 7):
        ws[f"G{row}"] = label
        style_cell(ws[f"G{row}"], fill=CREAM, color=MUTED, bold=True, size=9, wrap=True)
        ws.merge_cells(f"H{row}:J{row}")
        ws[f"H{row}"] = formula
        style_cell(ws[f"H{row}"], fill=FORMULA, color=NAVY, bold=True, size=11, number_format=number_format)
    section(ws, 14, "B", "J", "MAPA DEL LIBRO", NAVY)
    map_rows = [
        ("Resumen", "Panel mensual con indicadores, desviación por categoría y gráfico."),
        ("Presupuesto", "Matriz anual planificada, inspirada en los libros mensuales y anuales."),
        ("Gastos reales", "Matriz calculada a partir del registro de movimientos."),
        ("Movimientos", "Diario de ingresos, gastos y ahorros con listas desplegables."),
        ("Flujo de caja", "Lectura mensual de ingresos, gastos, ahorro y disponible."),
        ("Metas y fondo", "Fondo de emergencia y objetivos con progreso y fecha."),
        ("Deudas", "Pagos comprometidos, interés estimado y meses restantes."),
        ("Listas", "Parámetros editables y catálogos que alimentan las validaciones."),
    ]
    for row, (name, text) in enumerate(map_rows, 15):
        ws[f"B{row}"] = name
        style_cell(ws[f"B{row}"], fill=MINT if row % 2 else CREAM, color=NAVY, bold=True, size=9)
        merge_text(ws, f"C{row}:J{row}", text, fill=MINT if row % 2 else CREAM, color=INK, size=9)
        ws.row_dimensions[row].height = 23
    merge_text(ws, "B25:J27", "Consejo: duplica el archivo antes de iniciar un nuevo año. El encuadre y las fórmulas quedan protegidos por una estructura simple, visible y fácil de adaptar.", fill=INPUT, color=INK, bold=True, size=10)


def build_listas(ws) -> None:
    set_frame(
        ws,
        "Parámetros y catálogos editables",
        "Cambia aquí los valores amarillos; las listas alimentan las validaciones del resto del libro.",
        "M",
        {"A": 3, "B": 25, "C": 19, "D": 3, "E": 16, "F": 3, "G": 14, "H": 3, "I": 21, "J": 3, "K": 14, "L": 3, "M": 15},
        AMBER,
    )
    section(ws, 6, "B", "C", "PARÁMETROS DEL HOGAR", AMBER)
    header_row(ws, 7, 2, ["Parámetro", "Valor"])
    parameters = [
        ("Mes seleccionado", "Enero"),
        ("Ingresos netos mensuales", 3200),
        ("Ahorro mensual objetivo", 300),
        ("Miembro A", "Ana"),
        ("Miembro B", "Carlos"),
        ("Moneda", "EUR"),
    ]
    for row, (label, value) in enumerate(parameters, 8):
        data_row(ws, row, 2, [label, value], inputs={1}, number_formats={1: CURRENCY} if isinstance(value, (int, float)) else {})
    add_table(ws, "tblParametros", "B7:C13")

    section(ws, 6, "E", "E", "MESES", TEAL)
    for row, month in enumerate(MONTHS, 7):
        ws[f"E{row}"] = month
        style_cell(ws[f"E{row}"], fill=MINT, color=INK, size=10)
    section(ws, 6, "G", "G", "TIPOS", TEAL)
    for row, value in enumerate(TYPES, 7):
        ws[f"G{row}"] = value
        style_cell(ws[f"G{row}"], fill=MINT, color=INK, size=10)
    section(ws, 6, "I", "I", "CATEGORÍAS", TEAL)
    for row, value in enumerate(CATEGORIES, 7):
        ws[f"I{row}"] = value
        style_cell(ws[f"I{row}"], fill=MINT, color=INK, size=10)
    section(ws, 6, "K", "K", "MIEMBROS", TEAL)
    for row, value in enumerate(MEMBERS, 7):
        ws[f"K{row}"] = value
        style_cell(ws[f"K{row}"], fill=MINT, color=INK, size=10)
    section(ws, 6, "M", "M", "ESTADOS", TEAL)
    for row, value in enumerate(STATES, 7):
        ws[f"M{row}"] = value
        style_cell(ws[f"M{row}"], fill=MINT, color=INK, size=10)
    merge_text(ws, "E21:M23", "Las listas son visibles para que las futuras plantillas puedan sustituir categorías sin esconder la lógica del libro.", fill=INPUT, color=INK, size=10)
    add_list_validation(ws, "=Listas!$E$7:$E$18", "C8")
    add_list_validation(ws, "=Listas!$K$7:$K$9", "C11:C12")
    add_list_validation(ws, "=Listas!$M$7:$M$8", "C13")


def build_presupuesto(ws) -> None:
    set_frame(
        ws,
        "Presupuesto planificado por mes",
        "Matriz anual con la distribución temporal que se repite en los ejemplos de presupuesto personal, ventas y empresas.",
        "O",
        {"A": 3, "B": 23, **{get_column_letter(col): 11 for col in range(3, 15)}, "O": 15},
        NAVY,
    )
    section(ws, 6, "B", "O", "PLAN ANUAL · CELDAS AMARILLAS = ENTRADA", NAVY)
    headers = ["Categoría", *MONTHS, "Total anual"]
    header_row(ws, 7, 2, headers)
    base = {
        "Vivienda": [1100, 1100, 1100, 1100, 1100, 1100, 1100, 1100, 1100, 1100, 1100, 1100],
        "Alimentación": [420, 420, 430, 430, 440, 440, 460, 460, 440, 440, 460, 500],
        "Transporte": [250, 250, 250, 250, 250, 250, 250, 250, 250, 250, 250, 250],
        "Salud": [350, 350, 350, 350, 350, 350, 350, 350, 350, 350, 350, 500],
        "Educación": [140, 140, 140, 140, 140, 140, 180, 180, 180, 180, 180, 180],
        "Ocio": [280, 280, 300, 300, 320, 320, 420, 420, 300, 300, 340, 450],
        "Otros": [160, 160, 170, 170, 180, 180, 190, 190, 180, 180, 200, 220],
    }
    for row, category in enumerate(CATEGORIES, 8):
        values = [category, *base[category], f"=SUM(C{row}:N{row})"]
        data_row(ws, row, 2, values, inputs=set(range(1, 13)), formulas={13}, number_formats={i: CURRENCY for i in range(1, 14)})
    total_row = 15
    values = ["TOTAL DEL HOGAR"] + [f"=SUM({get_column_letter(col)}8:{get_column_letter(col)}14)" for col in range(3, 16)]
    data_row(ws, total_row, 2, values, formulas=set(range(1, 14)), number_formats={i: CURRENCY for i in range(1, 14)}, row_fill=MINT)
    for col in range(2, 16):
        style_cell(ws.cell(total_row, col), fill=TEAL if col == 2 else MINT, color=WHITE if col == 2 else NAVY, bold=True, size=10, number_format=CURRENCY if col > 2 else None)
    add_table(ws, "tblPresupuesto", "B7:O14")
    merge_text(ws, "B18:O19", "La matriz conserva el patrón de los presupuestos mensual/anual: categorías en filas, meses en columnas y total anual calculado a la derecha.", fill=INPUT, color=INK, size=10)


def build_gastos_reales(ws) -> None:
    set_frame(
        ws,
        "Gasto real calculado desde los movimientos",
        "Cada importe se obtiene del diario de movimientos; no hace falta duplicar la entrada en dos hojas.",
        "O",
        {"A": 3, "B": 23, **{get_column_letter(col): 11 for col in range(3, 15)}, "O": 15},
        CORAL,
    )
    section(ws, 6, "B", "O", "REAL ACUMULADO · FÓRMULAS AZULES", CORAL)
    header_row(ws, 7, 2, ["Categoría", *MONTHS, "Total anual"])
    for row, category_row in enumerate(range(8, 15), 8):
        values: list[object] = [f"='Presupuesto'!B{category_row}"]
        for col in range(3, 15):
            letter = get_column_letter(col)
            values.append(
                f'=SUMIFS(\'Movimientos\'!$G$8:$G$200,\'Movimientos\'!$D$8:$D$200,"Gasto",\'Movimientos\'!$E$8:$E$200,$B{row},\'Movimientos\'!$C$8:$C$200,{letter}$7)'
            )
        values.append(f"=SUM(C{row}:N{row})")
        data_row(ws, row, 2, values, formulas=set(range(0, 14)), number_formats={i: CURRENCY for i in range(1, 14)})
    total_row = 15
    values = ["TOTAL DEL HOGAR"] + [f"=SUM({get_column_letter(col)}8:{get_column_letter(col)}14)" for col in range(3, 16)]
    data_row(ws, total_row, 2, values, formulas=set(range(1, 14)), number_formats={i: CURRENCY for i in range(1, 14)}, row_fill=MINT)
    for col in range(2, 16):
        style_cell(ws.cell(total_row, col), fill=CORAL if col == 2 else MINT, color=WHITE if col == 2 else NAVY, bold=True, size=10, number_format=CURRENCY if col > 2 else None)
    add_table(ws, "tblGastosReales", "B7:O14")
    merge_text(ws, "B18:O19", "La hoja sigue el patrón de los ejemplos de presupuesto versus gastado, pero centraliza el dato real en Movimientos para evitar descuadres.", fill=PALE_CORAL, color=INK, size=10)


def build_movimientos(ws) -> None:
    set_frame(
        ws,
        "Diario de ingresos, gastos y ahorros",
        "Registra cada operación una sola vez; las hojas de análisis la agrupan por mes y categoría.",
        "M",
        {"A": 3, "B": 14, "C": 14, "D": 12, "E": 21, "F": 24, "G": 14, "H": 15, "I": 15, "J": 3, "K": 18, "L": 18, "M": 18},
        TEAL,
    )
    section(ws, 6, "B", "I", "REGISTRO OPERATIVO · ENTRADA MANUAL", TEAL)
    headers = ["Fecha", "Mes", "Tipo", "Categoría", "Concepto", "Importe", "Pagado por", "Estado"]
    header_row(ws, 7, 2, headers)
    samples = [
        [date(2027, 1, 5), "Enero", "Gasto", "Alimentación", "Compra semanal", 86, "Familia", "Conciliado"],
        [date(2027, 1, 8), "Enero", "Gasto", "Vivienda", "Electricidad", 62, "Familia", "Conciliado"],
        [date(2027, 1, 12), "Enero", "Gasto", "Transporte", "Abono mensual", 48, "Ana", "Pendiente"],
        [date(2027, 1, 19), "Enero", "Gasto", "Salud", "Farmacia", 29, "Carlos", "Conciliado"],
        [date(2027, 1, 25), "Enero", "Ingreso", "Otros", "Nómina", 3200, "Familia", "Conciliado"],
        [date(2027, 1, 28), "Enero", "Ahorro", "Otros", "Fondo de emergencia", 250, "Familia", "Pendiente"],
    ]
    for row in range(8, 28):
        values = samples[row - 8] if row - 8 < len(samples) else [None] * 8
        data_row(ws, row, 2, values, inputs=set(range(8)), number_formats={0: DATE_FORMAT, 5: CURRENCY})
    add_table(ws, "tblMovimientos", "B7:I27")
    add_list_validation(ws, "=Listas!$E$7:$E$18", "C8:C200")
    add_list_validation(ws, "=Listas!$G$7:$G$9", "D8:D200")
    add_list_validation(ws, "=Listas!$I$7:$I$13", "E8:E200")
    add_list_validation(ws, "=Listas!$K$7:$K$9", "H8:H200")
    add_list_validation(ws, "=Listas!$M$7:$M$8", "I8:I200")
    section(ws, 6, "K", "M", "REGLAS DEL DIARIO", CORAL)
    merge_text(ws, "K7:M10", "Una fila = una operación. Usa Tipo para separar ingresos, gastos y ahorros; no introduzcas un gasto manual también en Gastos reales.", fill=CREAM, color=INK, size=10)
    merge_text(ws, "K12:M15", "Las listas desplegables mantienen consistentes los nombres que usan los resúmenes y gráficos.", fill=MINT, color=INK, size=10)


def build_flujo_caja(ws) -> None:
    set_frame(
        ws,
        "Flujo de caja mensual",
        "La lectura anual combina ingresos, gastos planificados, movimientos reales, ahorro y saldo disponible.",
        "Q",
        {"A": 3, "B": 14, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 15, "I": 15, "J": 25, "K": 3, **{get_column_letter(col): 12 for col in range(12, 18)}},
        NAVY,
    )
    section(ws, 6, "B", "J", "LECTURA MENSUAL", NAVY)
    headers = ["Mes", "Ingresos plan", "Ingresos real", "Gastos plan", "Gastos real", "Ahorro plan", "Ahorro real", "Disponible", "Comentarios"]
    header_row(ws, 7, 2, headers)
    for row, month in enumerate(MONTHS, 8):
        month_index = row - 7
        values = [
            month,
            "='Listas'!$C$9",
            f'=SUMIFS(\'Movimientos\'!$G$8:$G$200,\'Movimientos\'!$D$8:$D$200,"Ingreso",\'Movimientos\'!$C$8:$C$200,$B{row})',
            f"=SUM(INDEX('Presupuesto'!$C$8:$N$14,0,ROW()-7))",
            f"=SUM(INDEX('Gastos reales'!$C$8:$N$14,0,ROW()-7))",
            "='Listas'!$C$10",
            f'=SUMIFS(\'Movimientos\'!$G$8:$G$200,\'Movimientos\'!$D$8:$D$200,"Ahorro",\'Movimientos\'!$C$8:$C$200,$B{row})',
            f"=D{row}-F{row}-H{row}",
            "",
        ]
        data_row(ws, row, 2, values, inputs={8}, formulas=set(range(1, 8)), number_formats={i: CURRENCY for i in range(1, 8)})
    total_row = 20
    values = ["TOTAL ANUAL"] + [f"=SUM({get_column_letter(col)}8:{get_column_letter(col)}19)" for col in range(3, 10)] + [""]
    data_row(ws, total_row, 2, values, formulas=set(range(1, 8)), number_formats={i: CURRENCY for i in range(1, 8)}, row_fill=MINT)
    for col in range(2, 11):
        style_cell(ws.cell(total_row, col), fill=NAVY if col == 2 else MINT, color=WHITE if col == 2 else NAVY, bold=True, size=10, number_format=CURRENCY if col > 2 and col < 10 else None)
    add_table(ws, "tblFlujoCaja", "B7:J19")
    chart = LineChart()
    chart.title = "Ingresos planificados y reales"
    chart.y_axis.title = "Importe"
    chart.x_axis.title = "Mes"
    chart.style = 13
    chart.height = 7
    chart.width = 13
    chart.add_data(Reference(ws, min_col=3, max_col=4, min_row=7, max_row=19), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, min_row=8, max_row=19))
    chart.series[0].graphicalProperties.solidFill = TEAL
    chart.series[1].graphicalProperties.solidFill = CORAL
    ws.add_chart(chart, "L7")
    merge_text(ws, "B23:J24", "El saldo disponible se calcula como ingreso real menos gasto real y ahorro real; sirve como señal de revisión, no como asesoramiento financiero.", fill=INPUT, color=INK, size=10)


def build_resumen(ws) -> None:
    set_frame(
        ws,
        "El mes, en una sola pantalla",
        "Panel de decisión inspirado en los balances, reportes mensuales y comparaciones presupuesto/real de las fuentes.",
        "N",
        {"A": 3, "B": 22, "C": 16, "D": 16, "E": 16, "F": 14, "G": 3, **{get_column_letter(col): 12 for col in range(8, 15)}},
        TEAL,
    )
    ws["B5"] = "Mes seleccionado"
    style_cell(ws["B5"], fill=CREAM, color=MUTED, bold=True, size=9)
    ws["C5"] = "='Listas'!$C$8"
    style_cell(ws["C5"], fill=INPUT, color=NAVY, bold=True, size=10)
    add_list_validation(ws, "=Listas!$E$7:$E$18", "C5")
    flow_ref = lambda column: f"=INDEX('Flujo de caja'!${column}$8:${column}$19,MATCH($C$5,'Flujo de caja'!$B$8:$B$19,0))"
    card(ws, 2, 4, 6, 8, "Ingresos reales", flow_ref("D"), MINT)
    card(ws, 5, 7, 6, 8, "Gastos reales", flow_ref("F"), PALE_CORAL)
    card(ws, 8, 10, 6, 8, "Ahorro real", flow_ref("H"), MINT)
    card(ws, 11, 13, 6, 8, "Disponible", flow_ref("I"), INPUT)
    section(ws, 10, "B", "F", "GASTO POR CATEGORÍA", NAVY)
    header_row(ws, 11, 2, ["Categoría", "Plan", "Real", "Diferencia", "% gastado"])
    for row, category_row in enumerate(range(8, 15), 12):
        plan = f"=INDEX('Presupuesto'!$C$8:$N$14,MATCH($B{row},'Presupuesto'!$B$8:$B$14,0),MATCH($C$5,'Presupuesto'!$C$7:$N$7,0))"
        real = f"=INDEX('Gastos reales'!$C$8:$N$14,MATCH($B{row},'Gastos reales'!$B$8:$B$14,0),MATCH($C$5,'Gastos reales'!$C$7:$N$7,0))"
        values = [f"='Presupuesto'!B{category_row}", plan, real, f"=C{row}-D{row}", f"=IFERROR(D{row}/C{row},0)"]
        data_row(ws, row, 2, values, formulas={0, 1, 2, 3, 4}, number_formats={1: CURRENCY, 2: CURRENCY, 3: CURRENCY, 4: PERCENT})
    total_row = 19
    values = ["TOTAL", "=SUM(C12:C18)", "=SUM(D12:D18)", "=SUM(E12:E18)", "=IFERROR(D19/C19,0)"]
    data_row(ws, total_row, 2, values, formulas={1, 2, 3, 4}, number_formats={1: CURRENCY, 2: CURRENCY, 3: CURRENCY, 4: PERCENT}, row_fill=MINT)
    for col in range(2, 7):
        style_cell(ws.cell(total_row, col), fill=TEAL if col == 2 else MINT, color=WHITE if col == 2 else NAVY, bold=True, size=10, number_format=PERCENT if col == 6 else (CURRENCY if col > 2 else None))
    add_table(ws, "tblResumenCategorias", "B11:F18")
    ws.conditional_formatting.add("E12:E18", CellIsRule(operator="lessThan", formula=["0"], fill=solid(PALE_CORAL)))
    ws.conditional_formatting.add("F12:F18", DataBarRule(start_type="num", start_value=0, end_type="max", color=TEAL, showValue=True))
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Plan frente a real"
    chart.y_axis.title = "Categoría"
    chart.x_axis.title = "Importe"
    chart.height = 8
    chart.width = 13
    chart.add_data(Reference(ws, min_col=3, max_col=4, min_row=11, max_row=18), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, min_row=12, max_row=18))
    chart.series[0].graphicalProperties.solidFill = TEAL
    chart.series[1].graphicalProperties.solidFill = CORAL
    ws.add_chart(chart, "H11")
    merge_text(ws, "B22:F24", "La diferencia es Plan - Real: positiva = queda margen; negativa = la categoría ha superado lo previsto.", fill=INPUT, color=INK, size=10)


def build_metas(ws) -> None:
    set_frame(
        ws,
        "Metas de ahorro y fondo de emergencia",
        "El bloque de metas recoge el patrón de fondo necesario, ahorro acumulado, pendiente y fecha de logro.",
        "J",
        {"A": 3, "B": 23, "C": 14, "D": 11, "E": 15, "F": 14, "G": 14, "H": 12, "I": 15, "J": 14},
        AMBER,
    )
    section(ws, 6, "B", "J", "PARÁMETROS DE LA META", AMBER)
    labels = [("B7", "Mes actual"), ("E7", "Inflación mensual"), ("H7", "Ahorro mensual")]
    for coord, value in labels:
        ws[coord] = value
        style_cell(ws[coord], fill=CREAM, color=MUTED, bold=True, size=9)
    ws["C7"] = "='Listas'!$C$8"
    ws["F7"] = 0.02
    ws["I7"] = "='Listas'!$C$10"
    style_cell(ws["C7"], fill=FORMULA, color=NAVY, bold=True, size=10)
    style_cell(ws["F7"], fill=INPUT, color=INK, bold=True, size=10, number_format=PERCENT)
    style_cell(ws["I7"], fill=FORMULA, color=NAVY, bold=True, size=10, number_format=CURRENCY)
    header_row(ws, 9, 2, ["Meta", "Gasto mensual", "Meses", "Meta necesaria", "Ahorrado", "Pendiente", "Progreso", "Fecha objetivo", "Estado"])
    goals = [
        ["Fondo de emergencia", 2400, 6, "=C10*D10", 2500, "=MAX(E10-F10,0)", "=IFERROR(F10/E10,0)", date(2027, 12, 31), '=IF(H10>=1,"Completa","En curso")'],
        ["Vacaciones", 450, 4, "=C11*D11", 650, "=MAX(E11-F11,0)", "=IFERROR(F11/E11,0)", date(2027, 6, 30), '=IF(H11>=1,"Completa","En curso")'],
        ["Reparaciones del hogar", 300, 4, "=C12*D12", 300, "=MAX(E12-F12,0)", "=IFERROR(F12/E12,0)", date(2027, 10, 31), '=IF(H12>=1,"Completa","En curso")'],
    ]
    for row, values in enumerate(goals, 10):
        data_row(ws, row, 2, values, inputs={0, 1, 2, 4, 7}, formulas={3, 5, 6, 8}, number_formats={1: CURRENCY, 2: "0", 3: CURRENCY, 4: CURRENCY, 5: CURRENCY, 6: PERCENT, 7: DATE_FORMAT})
    total_row = 13
    values = ["TOTAL", "", "", "=SUM(E10:E12)", "=SUM(F10:F12)", "=SUM(G10:G12)", "=IFERROR(F13/E13,0)", "", ""]
    data_row(ws, total_row, 2, values, formulas={3, 4, 5, 6}, number_formats={3: CURRENCY, 4: CURRENCY, 5: CURRENCY, 6: PERCENT}, row_fill=MINT)
    add_table(ws, "tblMetas", "B9:J12")
    add_list_validation(ws, "=Listas!$I$7:$I$13", "B10:B100")
    ws.conditional_formatting.add("H10:H12", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=TEAL, showValue=True))
    merge_text(ws, "B16:J18", "El fondo de emergencia se expresa como meses de gasto, igual que en las fuentes de fondo de emergencia; las demás metas pueden usar un objetivo directo cambiando Meses a 1.", fill=INPUT, color=INK, size=10)


def build_deudas(ws) -> None:
    set_frame(
        ws,
        "Deudas y pagos comprometidos",
        "Bloque compacto inspirado en las hojas de comparación de hipotecas y amortización del préstamo.",
        "H",
        {"A": 3, "B": 22, "C": 15, "D": 14, "E": 15, "F": 15, "G": 18, "H": 15},
        CORAL,
    )
    section(ws, 6, "B", "H", "OBLIGACIONES ABIERTAS", CORAL)
    headers = ["Deuda", "Saldo actual", "Interés anual", "Pago mensual", "Meses restantes", "Interés estimado", "Estado"]
    header_row(ws, 7, 2, headers)
    rows = [
        ["Hipoteca", 180000, 0.035, 850, 240, "=C8*D8*E8/12", '=IF(E8>0,"Activa","Pagada")'],
        ["Préstamo coche", 9200, 0.07, 280, 36, "=C9*D9*E9/12", '=IF(E9>0,"Activa","Pagada")'],
        ["Tarjeta", 2400, 0.19, 150, 18, "=C10*D10*E10/12", '=IF(E10>0,"Activa","Pagada")'],
    ]
    for row, values in enumerate(rows, 8):
        data_row(ws, row, 2, values, inputs={0, 1, 2, 3, 4}, formulas={5, 6}, number_formats={1: CURRENCY, 2: PERCENT, 3: CURRENCY, 4: "0", 5: CURRENCY})
    total_row = 11
    values = ["TOTAL", "=SUM(C8:C10)", "", "=SUM(E8:E10)", "", "=SUM(G8:G10)", ""]
    data_row(ws, total_row, 2, values, formulas={1, 3, 5}, number_formats={1: CURRENCY, 3: CURRENCY, 5: CURRENCY}, row_fill=MINT)
    for col in range(2, 9):
        style_cell(ws.cell(total_row, col), fill=CORAL if col == 2 else MINT, color=WHITE if col == 2 else NAVY, bold=True, size=10, number_format=CURRENCY if col in {3, 5, 7} else None)
    add_table(ws, "tblDeudas", "B7:H10")
    ws.conditional_formatting.add("C8:C10", DataBarRule(start_type="min", end_type="max", color=CORAL, showValue=True))
    merge_text(ws, "B14:H17", "El interés estimado es una señal orientativa basada en saldo × tasa × meses / 12; para una amortización contractual completa, sustituye este bloque por un calendario específico.", fill=PALE_CORAL, color=INK, size=10)


def build_workbook(output: Path) -> Path:
    source_count = len([path for path in SOURCE_DIR.iterdir() if path.is_file()])
    workbook = Workbook()
    workbook.remove(workbook.active)
    sheets = {
        "Inicio": workbook.create_sheet("Inicio"),
        "Resumen": workbook.create_sheet("Resumen"),
        "Presupuesto": workbook.create_sheet("Presupuesto"),
        "Gastos reales": workbook.create_sheet("Gastos reales"),
        "Movimientos": workbook.create_sheet("Movimientos"),
        "Flujo de caja": workbook.create_sheet("Flujo de caja"),
        "Metas y fondo": workbook.create_sheet("Metas y fondo"),
        "Deudas": workbook.create_sheet("Deudas"),
        "Listas": workbook.create_sheet("Listas"),
    }
    build_inicio(sheets["Inicio"], source_count)
    build_resumen(sheets["Resumen"])
    build_presupuesto(sheets["Presupuesto"])
    build_gastos_reales(sheets["Gastos reales"])
    build_movimientos(sheets["Movimientos"])
    build_flujo_caja(sheets["Flujo de caja"])
    build_metas(sheets["Metas y fondo"])
    build_deudas(sheets["Deudas"])
    build_listas(sheets["Listas"])
    for sheet in sheets.values():
        clear_bottom_border(sheet, 7)
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.properties.title = "Presupuesto familiar · base construida desde sources"
    workbook.properties.subject = "Planificación mensual, movimientos, flujo de caja, metas y deudas"
    workbook.properties.creator = "Creador de Excels"
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    workbook.close()
    # Add the SVG as a native drawing after openpyxl closes the workbook. This
    # keeps the direct builder intact while preserving tables and validations.
    from embed_svg_icon import embed_icon
    icon_output = output.with_name(f"{output.stem}.icon.xlsx")
    embed_icon(output, icon_output, ICON_PATH)
    icon_output.replace(output)
    contract = family_base_contract()
    artifact_dir = output.parent.parent
    behavior_path = artifact_dir / "model_behavior.md"
    write_model_behavior_markdown(behavior_path, contract)
    write_model_contract_json(artifact_dir / "model_contract.json", contract)
    quality = validate_model_file(
        output,
        expected_sheets=list(sheets),
        model_contract=contract,
        model_contract_path=behavior_path,
    )
    if not quality["passed"]:
        raise RuntimeError(f"El modelo no supera la validación global: {quality['errors']}")
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description="Crea un libro Excel familiar nativo desde cero a partir de templates/sources")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    output = args.output if args.output.is_absolute() else ROOT / args.output
    print(f"WORKBOOK={build_workbook(output)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
