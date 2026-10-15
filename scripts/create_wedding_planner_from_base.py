from __future__ import annotations

"""Create a wedding planner by inheriting the approved native Excel base."""

import argparse
import warnings
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.chart import BarChart
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from embed_svg_icon import embed_icon


ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "templates" / "excel-first-base" / "master_workbook.xlsx"
ICON = ROOT / "assets" / "wedding-planner-icon.svg"
DEFAULT_OUTPUT = ROOT / "outputs" / "listings" / "wedding-planner-base" / "app" / "planificador-de-boda-base.xlsx"

WEDDING_PALETTE = {
    "NAVY": "5A2E46",
    "TEAL": "A85D73",
    "MINT": "F5E3E8",
    "CORAL": "9A6178",
    "AMBER": "B9798E",
    "CREAM": "FBF7F4",
    "INK": "3B2B32",
    "MUTED": "7E6871",
    "LINE": "E1D2D7",
    "INPUT": "F6E8DF",
    "FORMULA": "F3EAF0",
    "PALE": "F8EFF2",
    "WHITE": "FFFFFF",
}

OLD_TO_NEW = {
    "164E63": WEDDING_PALETTE["NAVY"],
    "0F766E": WEDDING_PALETTE["TEAL"],
    "DDF3EF": WEDDING_PALETTE["MINT"],
    "477D8E": WEDDING_PALETTE["CORAL"],
    "5C8D94": WEDDING_PALETTE["AMBER"],
    "F4F8F7": WEDDING_PALETTE["CREAM"],
    "20343D": WEDDING_PALETTE["INK"],
    "60767D": WEDDING_PALETTE["MUTED"],
    "C6D4D8": WEDDING_PALETTE["LINE"],
    "E7F2EF": WEDDING_PALETTE["INPUT"],
    "E8F0F5": WEDDING_PALETTE["FORMULA"],
    "EDF4F5": WEDDING_PALETTE["PALE"],
    "FFFFFF": WEDDING_PALETTE["WHITE"],
}

SHEET_RENAMES = {
    "Gastos reales": "Pagos reales",
    "Movimientos": "Proveedores",
    "Flujo de caja": "Flujo de pagos",
    "Metas y fondo": "Checklist",
    "Deudas": "Invitados",
}

CATEGORIES = [
    "Lugar y ceremonia",
    "Banquete",
    "Fotografía y vídeo",
    "Música",
    "Decoración y flores",
    "Vestuario y belleza",
    "Papelería y detalles",
]
MONTHS = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
TYPES = ["Pago", "Ingreso", "Ajuste"]
PAYERS = ["Pareja", "Familia", "Proveedor"]
STATES = ["Pendiente", "Pagado", "Confirmado"]


def mapped_color(color) -> str | None:
    if not color or color.type != "rgb" or not color.rgb:
        return None
    return OLD_TO_NEW.get(color.rgb[-6:].upper(), color.rgb[-6:].upper())


def retheme_workbook(workbook) -> None:
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.fill.fill_type == "solid":
                    color = mapped_color(cell.fill.fgColor)
                    if color:
                        cell.fill = PatternFill(fill_type="solid", fgColor=color)
                font_color = mapped_color(cell.font.color)
                if font_color:
                    cell.font = cell.font.copy(color=font_color)
                cell.border = Border(
                    left=Side(style=cell.border.left.style, color=mapped_color(cell.border.left.color)),
                    right=Side(style=cell.border.right.style, color=mapped_color(cell.border.right.color)),
                    top=Side(style=cell.border.top.style, color=mapped_color(cell.border.top.color)),
                    bottom=Side(style=cell.border.bottom.style, color=mapped_color(cell.border.bottom.color)),
                )
        tab_color = WEDDING_PALETTE["TEAL"] if sheet.title in {"Inicio", "Resumen", "Proveedores", "Listas"} else WEDDING_PALETTE["NAVY"]
        sheet.sheet_properties.tabColor = tab_color


def set_frame_text(sheet, title: str, subtitle: str) -> None:
    sheet["B2"] = f"PLANIFICADOR DE BODA · {sheet.title.upper()}"
    sheet["B3"] = title
    sheet["B4"] = subtitle


def clear_validations(sheet) -> None:
    sheet.data_validations.dataValidation.clear()


def add_validation(sheet, formula: str, cell_range: str) -> None:
    validation = DataValidation(type="list", formula1=formula, allow_blank=True)
    validation.error = "Selecciona un valor de la lista."
    validation.errorTitle = "Valor no válido"
    validation.prompt = "Usa una opción de las listas del libro."
    validation.promptTitle = "Lista"
    sheet.add_data_validation(validation)
    validation.add(cell_range)


def replace_sheet_references(workbook) -> None:
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.startswith("="):
                    continue
                value = cell.value
                for old, new in SHEET_RENAMES.items():
                    value = value.replace(f"'{old}'", f"'{new}'")
                value = value.replace('"Gasto"', '"Pago"').replace('"Ahorro"', '"Ajuste"')
                cell.value = value


def replace_visible_text(workbook) -> None:
    replacements = {
        "PRESUPUESTO FAMILIAR": "PLANIFICADOR DE BODA",
        "Parámetros del hogar": "Parámetros de la boda",
        "PARÁMETROS DEL HOGAR": "PARÁMETROS DE LA BODA",
        "GASTO POR CATEGORÍA": "PRESUPUESTO POR CATEGORÍA",
        "Gasto real calculado desde los movimientos": "Pagos reales calculados desde proveedores",
        "Diario de ingresos, gastos y ahorros": "Registro de proveedores y pagos",
        "Flujo de caja mensual": "Flujo de pagos de la boda",
        "Metas de ahorro y fondo de emergencia": "Checklist de la boda",
        "Deudas y pagos comprometidos": "Invitados y mesas",
    }
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and not cell.value.startswith("="):
                    for old, new in replacements.items():
                        cell.value = cell.value.replace(old, new)


def set_values(sheet, rows: dict[str, object]) -> None:
    for coordinate, value in rows.items():
        sheet[coordinate] = value


def build_inicio(workbook) -> None:
    sheet = workbook["Inicio"]
    set_frame_text(sheet, "Organiza tu boda con una visión clara", "Presupuesto, proveedores, pagos, invitados y tareas en un solo libro.")
    set_values(sheet, {
        "B6": "CÓMO USARLO", "G6": "PARÁMETROS RÁPIDOS", "B14": "MAPA DEL LIBRO",
        "C7": "Define la celebración", "E7": "Fecha, presupuesto e invitados en Listas.",
        "C8": "Planifica el presupuesto", "E8": "Distribuye los importes por categoría y momento.",
        "C9": "Registra proveedores", "E9": "Anota contratos, pagos pendientes y responsables.",
        "C10": "Cierra cada decisión", "E10": "Usa Checklist, Invitados y Resumen para avanzar.",
        "G7": "Fecha de boda", "G8": "Presupuesto total", "G9": "Invitados previstos", "G10": "Pareja A", "G11": "Pareja B",
        "H7": "='Listas'!$C$8", "H8": "='Listas'!$C$9", "H9": "='Listas'!$C$10", "H10": "='Listas'!$C$11", "H11": "='Listas'!$C$12",
        "B15": "Resumen", "C15": "Visión general del presupuesto, pagos, pendientes e invitados confirmados.",
        "B16": "Presupuesto", "C16": "Plan anual por categorías y calendario de desembolsos.",
        "B17": "Pagos reales", "C17": "Importes pagados y comparativa frente al plan.",
        "B18": "Proveedores", "C18": "Registro de contratos, pagos, responsables y estados.",
        "B19": "Flujo de pagos", "C19": "Lectura mensual de presupuesto, pagos y saldo previsto.",
        "B20": "Checklist", "C20": "Tareas clave con responsables, fechas y progreso.",
        "B21": "Invitados", "C21": "Conteo por grupo, confirmación, menú y mesa.",
        "B22": "Listas", "C22": "Parámetros y catálogos editables del libro.",
        "B25": "Consejo: duplica el archivo para cada boda y cambia primero las celdas de Listas. La estructura, la navegación y las fórmulas ya están preparadas para crecer sin perder el encuadre.",
    })
    sheet["H7"].number_format = "dd/mm/yyyy"
    sheet["H8"].number_format = "#,##0.00"
    sheet["H9"].number_format = "0"


def build_listas(workbook) -> None:
    sheet = workbook["Listas"]
    set_frame_text(sheet, "Parámetros y catálogos de la boda", "Cambia aquí los valores de entrada; las listas alimentan las validaciones del resto del libro.")
    set_values(sheet, {
        "B6": "PARÁMETROS DE LA BODA", "E6": "MESES", "G6": "TIPOS", "I6": "CATEGORÍAS", "K6": "RESPONSABLES", "M6": "ESTADOS",
        "B7": "Parámetro", "C7": "Valor",
        "B8": "Fecha de boda", "C8": date(2027, 6, 19),
        "B9": "Presupuesto total", "C9": 18000,
        "B10": "Invitados previstos", "C10": 120,
        "B11": "Pareja A", "C11": "Laura",
        "B12": "Pareja B", "C12": "Álvaro",
        "B13": "Moneda", "C13": "EUR",
        "E21": "Las celdas de entrada definen la boda. Mantén visibles estas listas para poder adaptar categorías, responsables y estados sin romper las fórmulas.",
    })
    for row in range(7, 19):
        sheet[f"E{row}"] = MONTHS[row - 7] if row - 7 < len(MONTHS) else None
    for row, value in enumerate(TYPES, 7):
        sheet[f"G{row}"] = value
    for row, value in enumerate(CATEGORIES, 7):
        sheet[f"I{row}"] = value
    for row, value in enumerate(PAYERS, 7):
        sheet[f"K{row}"] = value
    for row, value in enumerate(STATES, 7):
        sheet[f"M{row}"] = value
    sheet["C8"].number_format = "dd/mm/yyyy"
    sheet["C9"].number_format = "#,##0.00"
    sheet["C10"].number_format = "0"


def build_presupuesto(workbook) -> None:
    sheet = workbook["Presupuesto"]
    set_frame_text(sheet, "Presupuesto de boda por categoría", "Distribuye el coste previsto y conserva la lectura mensual para anticipar los desembolsos.")
    sheet["B6"] = "PRESUPUESTO DE BODA · CELDAS DE COLOR = ENTRADA"
    values = {
        "Lugar y ceremonia": [0, 0, 2000, 0, 1500, 0, 0, 0, 0, 0, 0, 0],
        "Banquete": [0, 0, 0, 0, 0, 5000, 0, 0, 0, 0, 0, 0],
        "Fotografía y vídeo": [0, 0, 0, 0, 0, 1800, 0, 0, 0, 0, 0, 0],
        "Música": [0, 0, 0, 0, 0, 1200, 0, 0, 0, 0, 0, 0],
        "Decoración y flores": [0, 0, 0, 0, 0, 1600, 0, 0, 0, 0, 0, 0],
        "Vestuario y belleza": [300, 800, 300, 0, 0, 600, 0, 0, 0, 0, 0, 0],
        "Papelería y detalles": [250, 250, 350, 350, 400, 500, 0, 0, 0, 0, 0, 0],
    }
    for row, category in enumerate(CATEGORIES, 8):
        sheet.cell(row, 2).value = category
        for col, amount in enumerate(values[category], 3):
            sheet.cell(row, col).value = amount
        sheet.cell(row, 15).value = f"=SUM(C{row}:N{row})"
    sheet["B15"] = "TOTAL DE LA BODA"
    for col in range(3, 16):
        letter = chr(64 + col)
        sheet.cell(15, col).value = f"=SUM({letter}8:{letter}14)"
    sheet["B18"] = "La matriz conserva el patrón mensual/anual de la base: categorías en filas, meses en columnas y total calculado a la derecha. Puedes cambiar los importes sin mover el encuadre."


def build_pagos_reales(workbook) -> None:
    sheet = workbook["Pagos reales"]
    set_frame_text(sheet, "Pagos reales de proveedores", "Cada importe se obtiene del registro de proveedores para comparar lo previsto con lo ya pagado.")
    sheet["B6"] = "PAGOS REALES · FÓRMULAS DESDE PROVEEDORES"
    for row, category_row in enumerate(range(8, 15), 8):
        sheet.cell(row, 2).value = f"='Presupuesto'!B{category_row}"
        for col in range(3, 15):
            letter = get_column_letter(col)
            sheet.cell(row, col).value = f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Pago",\'Proveedores\'!$E$8:$E$200,$B{row},\'Proveedores\'!$C$8:$C$200,{letter}$7)'
        sheet.cell(row, 15).value = f"=SUM(C{row}:N{row})"
    sheet["B15"] = "TOTAL PAGADO"
    for col in range(3, 16):
        letter = get_column_letter(col)
        sheet.cell(15, col).value = f"=SUM({letter}8:{letter}14)"
    sheet["B18"] = "La hoja sigue el patrón de presupuesto frente a real, pero centraliza los pagos en Proveedores para evitar duplicidades y facilitar la revisión de contratos."


def build_proveedores(workbook) -> None:
    sheet = workbook["Proveedores"]
    set_frame_text(sheet, "Registro de proveedores y pagos", "Anota cada contrato o desembolso una sola vez; las hojas de análisis se alimentan de este diario.")
    sheet["B6"] = "REGISTRO OPERATIVO · PROVEEDORES Y PAGOS"
    headers = ["Fecha", "Mes", "Tipo", "Categoría", "Concepto", "Importe", "Pagado por", "Estado"]
    for col, header in enumerate(headers, 2):
        sheet.cell(7, col).value = header
    samples = [
        [date(2027, 1, 10), "Enero", "Pago", "Lugar y ceremonia", "Reserva del espacio", 2000, "Pareja", "Pagado"],
        [date(2027, 2, 14), "Febrero", "Pago", "Vestuario y belleza", "Vestido y pruebas", 800, "Pareja", "Pagado"],
        [date(2027, 3, 4), "Marzo", "Pago", "Papelería y detalles", "Invitaciones", 350, "Familia", "Pendiente"],
        [date(2027, 4, 2), "Abril", "Pago", "Fotografía y vídeo", "Reserva fotógrafo", 600, "Pareja", "Pendiente"],
        [date(2027, 5, 20), "Mayo", "Ingreso", "Banquete", "Aportación familiar", 1500, "Familia", "Confirmado"],
        [date(2027, 6, 19), "Junio", "Ajuste", "Decoración y flores", "Fondo de imprevistos", 300, "Pareja", "Pendiente"],
    ]
    for row in range(8, 28):
        values = samples[row - 8] if row - 8 < len(samples) else [None] * 8
        for col, value in enumerate(values, 2):
            sheet.cell(row, col).value = value
    sheet["K6"] = "REGLAS DEL DIARIO"
    sheet["K7"] = "Una fila = un contrato, pago o ajuste. Usa Tipo y Estado para que los resúmenes distingan lo comprometido de lo abonado."
    sheet["K12"] = "Las listas desplegables mantienen consistentes categorías, responsables y estados en todo el libro."
    for row in range(8, 28):
        sheet.cell(row, 2).number_format = "dd/mm/yyyy"
        sheet.cell(row, 7).number_format = "#,##0.00"


def build_flujo_pagos(workbook) -> None:
    sheet = workbook["Flujo de pagos"]
    set_frame_text(sheet, "Flujo mensual de pagos", "La lectura anual combina presupuesto, ingresos, pagos reales, ajustes y saldo previsto.")
    sheet["B6"] = "LECTURA MENSUAL DE LA BODA"
    headers = ["Mes", "Presupuesto plan", "Ingresos", "Pagos previstos", "Pagos reales", "Ajustes", "Saldo previsto", "Comentarios", "Estado"]
    for col, header in enumerate(headers, 2):
        sheet.cell(7, col).value = header
    for row, month in enumerate(MONTHS, 8):
        letter = get_column_letter(row - 5)
        sheet.cell(row, 2).value = month
        sheet.cell(row, 3).value = f"=SUM('Presupuesto'!{letter}$8:{letter}$14)"
        sheet.cell(row, 4).value = f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Ingreso",\'Proveedores\'!$C$8:$C$200,$B{row})'
        sheet.cell(row, 5).value = f"=C{row}"
        sheet.cell(row, 6).value = f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Pago",\'Proveedores\'!$C$8:$C$200,$B{row})'
        sheet.cell(row, 7).value = f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Ajuste",\'Proveedores\'!$C$8:$C$200,$B{row})'
        sheet.cell(row, 8).value = f"=C{row}-F{row}+D{row}-G{row}"
        sheet.cell(row, 9).value = ""
    sheet["B20"] = "TOTAL ANUAL"
    for col in range(3, 9):
        letter = get_column_letter(col)
        sheet.cell(20, col).value = f"=SUM({letter}8:{letter}19)"
    sheet["I20"] = None
    sheet["J20"] = None
    sheet["B23"] = "El saldo previsto no sustituye una revisión contractual: sirve para ver cuándo se concentran los desembolsos y qué pagos quedan pendientes."
    if sheet._charts:
        sheet._charts[0].title = "Ingresos y pagos de la boda"


def build_resumen(workbook) -> None:
    sheet = workbook["Resumen"]
    set_frame_text(sheet, "El estado de la boda en una sola pantalla", "Panel de decisión para presupuesto, pagos, pendientes e invitados confirmados.")
    sheet["B5"] = "Fecha de boda"
    sheet["C5"] = "='Listas'!$C$8"
    sheet["C5"].number_format = "dd/mm/yyyy"
    cards = {
        "B6": "PRESUPUESTO TOTAL", "B7": "=SUM('Presupuesto'!C8:N14)",
        "E6": "PAGOS REALES", "E7": '=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Pago")',
        "H6": "PENDIENTE", "H7": "=B7-E7",
        "K6": "INVITADOS CONFIRMADOS", "K7": "=SUM('Invitados'!D8:D10)",
    }
    set_values(sheet, cards)
    sheet["B10"] = "PRESUPUESTO Y PAGOS POR CATEGORÍA"
    headers = ["Categoría", "Plan", "Pagado", "Pendiente", "% pagado"]
    for col, header in enumerate(headers, 2):
        sheet.cell(11, col).value = header
    for row, source_row in enumerate(range(8, 15), 12):
        sheet.cell(row, 2).value = f"='Presupuesto'!B{source_row}"
        sheet.cell(row, 3).value = f"=SUM('Presupuesto'!C{source_row}:N{source_row})"
        sheet.cell(row, 4).value = f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Pago",\'Proveedores\'!$E$8:$E$200,$B{row})'
        sheet.cell(row, 5).value = f"=C{row}-D{row}"
        sheet.cell(row, 6).value = f"=IFERROR(D{row}/C{row},0)"
    sheet["B19"] = "TOTAL"
    sheet["C19"] = "=SUM(C12:C18)"
    sheet["D19"] = "=SUM(D12:D18)"
    sheet["E19"] = "=SUM(E12:E18)"
    sheet["F19"] = "=IFERROR(D19/C19,0)"
    sheet["B22"] = "La diferencia es Plan - Pagado: positiva = aún queda presupuesto; negativa = la categoría ha superado lo previsto."
    if sheet._charts:
        chart = sheet._charts[0]
        chart.title = "Presupuesto frente a pagos"
        chart.dLbls = DataLabelList()
        chart.dLbls.showVal = False


def build_checklist(workbook) -> None:
    sheet = workbook["Checklist"]
    set_frame_text(sheet, "Checklist de la boda", "Convierte las decisiones importantes en tareas con responsable, fecha y estado.")
    sheet["B6"] = "TAREAS Y MOMENTOS CLAVE"
    headers = ["Tarea", "Área", "Responsable", "Fecha límite", "Estado", "Prioridad", "Días", "Notas", "Progreso"]
    for col, header in enumerate(headers, 2):
        sheet.cell(9, col).value = header
    rows = [
        ["Cerrar contrato del lugar", "Lugar", "Pareja", date(2027, 1, 31), "Confirmado", "Alta", "=MAX(E10-TODAY(),0)", "Revisar horarios y montaje", '=IF(F10="Confirmado",1,0)'],
        ["Enviar invitaciones", "Invitados", "Familia", date(2027, 3, 15), "En curso", "Alta", "=MAX(E11-TODAY(),0)", "Lista final pendiente", '=IF(F11="Confirmado",1,0)'],
        ["Confirmar menú y música", "Celebración", "Pareja", date(2027, 5, 20), "Pendiente", "Media", "=MAX(E12-TODAY(),0)", "Cita con proveedores", '=IF(F12="Confirmado",1,0)'],
    ]
    for row, values in enumerate(rows, 10):
        for col, value in enumerate(values, 2):
            sheet.cell(row, col).value = value
        sheet.cell(row, 5).number_format = "dd/mm/yyyy"
        sheet.cell(row, 8).number_format = "0"
        sheet.cell(row, 10).number_format = "0%"
    sheet["B16"] = "El checklist comparte la misma retícula de seguimiento de la base, pero cambia metas financieras por decisiones concretas de la boda."
    sheet.conditional_formatting._cf_rules.clear()
    sheet.conditional_formatting.add("J10:J12", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=WEDDING_PALETTE["TEAL"], showValue=True))


def build_invitados(workbook) -> None:
    sheet = workbook["Invitados"]
    set_frame_text(sheet, "Invitados y mesas", "Controla grupos, confirmaciones, preferencias de menú y distribución de mesas.")
    sheet["B6"] = "LISTA DE INVITADOS"
    headers = ["Grupo", "Invitados", "Confirmados", "Menú", "Mesa", "Estado", "Notas"]
    for col, header in enumerate(headers, 2):
        sheet.cell(7, col).value = header
    rows = [
        ["Familia pareja A", 35, 30, "Estándar", 1, "Confirmado", "Pendientes 5 respuestas"],
        ["Familia pareja B", 30, 26, "Vegetariano", 2, "Pendiente", "Revisar teléfonos"],
        ["Amigos", 55, 42, "Mixto", 3, "Pendiente", "Confirmar acompañantes"],
    ]
    for row, values in enumerate(rows, 8):
        for col, value in enumerate(values, 2):
            sheet.cell(row, col).value = value
    sheet["B11"] = "TOTAL"
    sheet["C11"] = "=SUM(C8:C10)"
    sheet["D11"] = "=SUM(D8:D10)"
    sheet["B14"] = "Las cifras de confirmados alimentan la tarjeta del Resumen y ayudan a decidir menú, mesas y cantidad de papelería."
    sheet.conditional_formatting._cf_rules.clear()
    sheet.conditional_formatting.add("D8:D10", DataBarRule(start_type="num", start_value=0, end_type="max", color=WEDDING_PALETTE["TEAL"], showValue=True))


def reset_tables(workbook) -> None:
    table_names = {
        "tblMovimientos": "tblProveedores",
        "tblFlujoCaja": "tblFlujoPagos",
        "tblMetas": "tblChecklist",
        "tblDeudas": "tblInvitados",
    }
    for sheet in workbook.worksheets:
        for table in sheet.tables.values():
            if table.name in table_names:
                table.name = table_names[table.name]
            table.tableStyleInfo = None


def build_workbook(output: Path) -> Path:
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(SOURCE.read_bytes())
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message=r"The image .*\.svg will be removed")
        warnings.filterwarnings("ignore", message=r"Data Validation extension is not supported")
        workbook = load_workbook(output)
    for old, new in SHEET_RENAMES.items():
        workbook[old].title = new
    retheme_workbook(workbook)
    replace_visible_text(workbook)
    build_inicio(workbook)
    build_listas(workbook)
    build_presupuesto(workbook)
    build_pagos_reales(workbook)
    build_proveedores(workbook)
    build_flujo_pagos(workbook)
    build_resumen(workbook)
    build_checklist(workbook)
    build_invitados(workbook)
    reset_tables(workbook)
    for sheet in workbook.worksheets:
        clear_validations(sheet)
    add_validation(workbook["Proveedores"], "=Listas!$E$7:$E$18", "C8:C200")
    add_validation(workbook["Proveedores"], "=Listas!$G$7:$G$9", "D8:D200")
    add_validation(workbook["Proveedores"], "=Listas!$I$7:$I$13", "E8:E200")
    add_validation(workbook["Proveedores"], "=Listas!$K$7:$K$9", "H8:H200")
    add_validation(workbook["Proveedores"], "=Listas!$M$7:$M$9", "I8:I200")
    add_validation(workbook["Checklist"], "=Listas!$K$7:$K$9", "D10:D100")
    add_validation(workbook["Checklist"], "=Listas!$M$7:$M$9", "F10:F100")
    add_validation(workbook["Invitados"], "=Listas!$M$7:$M$9", "G8:G100")
    replace_sheet_references(workbook)
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.properties.title = "Planificador de boda · base nativa"
    workbook.properties.subject = "Presupuesto, proveedores, pagos, invitados y checklist de boda"
    workbook.properties.creator = "Creador de Excels"
    workbook.save(output)
    workbook.close()
    embed_icon(output, output, ICON)
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description="Crea un planificador de boda heredando la base nativa aprobada")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    output = args.output if args.output.is_absolute() else ROOT / args.output
    # Keep this original entry point compatible while delegating the final save
    # to the native COM builder, which preserves drawings and SVG artwork.
    from create_wedding_planner_native import build_workbook as native_build_workbook
    print(f"WORKBOOK={native_build_workbook(output)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
