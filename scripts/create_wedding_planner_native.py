from __future__ import annotations

"""Create the wedding planner by editing the approved XLSX base in Excel itself.

The base workbook contains native shapes, SVG artwork, charts and tables.  Using
Excel COM for the edit keeps those objects intact; openpyxl is intentionally not
used for the final save because it cannot round-trip the native drawing layer.
"""

import argparse
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

import win32com.client as win32
from openpyxl.utils.datetime import to_excel

from model_quality import validate_model_file
from model_contracts import wedding_base_contract, write_model_behavior_markdown, write_model_contract_json


ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "templates" / "excel-first-base" / "master_workbook.xlsx"
ICON = ROOT / "assets" / "wedding-planner-icon.svg"
DEFAULT_OUTPUT = ROOT / "outputs" / "listings" / "wedding-planner-base" / "app" / "planificador-de-boda-base.xlsx"

PALETTE = {
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
    "164E63": PALETTE["NAVY"],
    "0F766E": PALETTE["TEAL"],
    "DDF3EF": PALETTE["MINT"],
    "477D8E": PALETTE["CORAL"],
    "5C8D94": PALETTE["AMBER"],
    "F4F8F7": PALETTE["CREAM"],
    "20343D": PALETTE["INK"],
    "60767D": PALETTE["MUTED"],
    "C6D4D8": PALETTE["LINE"],
    "E7F2EF": PALETTE["INPUT"],
    "E8F0F5": PALETTE["FORMULA"],
    "EDF4F5": PALETTE["PALE"],
    "FFFFFF": PALETTE["WHITE"],
}

SHEET_RENAMES = {
    "Gastos reales": "Pagos reales",
    "Movimientos": "Proveedores",
    "Flujo de caja": "Flujo de pagos",
    "Metas y fondo": "Checklist",
    "Deudas": "Invitados",
}

MONTHS = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
CATEGORIES = [
    "Lugar y ceremonia", "Banquete", "Fotografía y vídeo", "Música",
    "Decoración y flores", "Vestuario y belleza", "Papelería y detalles",
]
TYPES = ["Pago", "Ingreso", "Ajuste"]
PAYERS = ["Pareja", "Familia", "Proveedor"]
STATES = ["Pendiente", "Pagado", "Confirmado"]


def rgb(hex_color: str) -> int:
    value = hex_color.lstrip("#")
    red, green, blue = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return red + (green << 8) + (blue << 16)


def set_value(sheet, address: str, value) -> None:
    if isinstance(value, datetime):
        # Excel COM applies the machine timezone to Python datetimes.  Writing
        # the serial keeps the intended calendar date stable across locales.
        sheet.Range(address).Value2 = to_excel(value)
    else:
        sheet.Range(address).Value = value


def set_formula(sheet, address: str, formula: str) -> None:
    sheet.Range(address).Formula = formula


def clear_contents(sheet, address: str) -> None:
    target = sheet.Range(address)
    for row in range(target.Row, target.Row + target.Rows.Count):
        for col in range(target.Column, target.Column + target.Columns.Count):
            cell = sheet.Cells(row, col)
            if cell.MergeCells:
                # Merged blocks are rewritten through their top-left cell below;
                # clearing a partial merged range makes Excel reject the action.
                continue
            else:
                cell.ClearContents()


def retheme_package(path: Path) -> None:
    """Re-theme the OOXML style/chart definitions without round-tripping drawings."""
    temporary = path.with_suffix(".palette.tmp.xlsx")
    replacements = {}
    for old, new in OLD_TO_NEW.items():
        replacements[f"FF{old}".encode("ascii")] = f"FF{new}".encode("ascii")
        replacements[old.encode("ascii")] = new.encode("ascii")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if not item.filename.startswith("xl/media/"):
                for old, new in replacements.items():
                    data = data.replace(old, new)
            target.writestr(item, data)
    temporary.replace(path)


def retheme_workbook(workbook) -> None:
    """Apply the small native-only touches that are not represented in styles.xml."""
    for sheet in workbook.Worksheets:
        sheet.Tab.Color = rgb(PALETTE["TEAL"] if sheet.Name in {"Inicio", "Resumen", "Proveedores", "Listas"} else PALETTE["NAVY"])
        # The shared base deliberately has no divider beneath the row-7 headers.
        try:
            sheet.Rows(7).Borders(9).LineStyle = 0
        except Exception:
            pass


def set_frame(sheet, title: str, subtitle: str) -> None:
    set_value(sheet, "B2", f"PLANIFICADOR DE BODA · {sheet.Name.upper()}")
    set_value(sheet, "B3", title)
    set_value(sheet, "B4", subtitle)


def build_inicio(wb) -> None:
    sheet = wb.Worksheets("Inicio")
    set_frame(sheet, "Organiza tu boda con una visión clara", "Presupuesto, proveedores, pagos, invitados y tareas en un solo libro.")
    clear_contents(sheet, "B6:O25")
    values = {
        "B6": "CÓMO USARLO", "G6": "PARÁMETROS RÁPIDOS", "B14": "MAPA DEL LIBRO",
        "C7": "Define la celebración", "E7": "Fecha, presupuesto e invitados en Listas.",
        "C8": "Planifica el presupuesto", "E8": "Distribuye los importes por categoría y momento.",
        "C9": "Registra proveedores", "E9": "Anota contratos, pagos pendientes y responsables.",
        "C10": "Cierra cada decisión", "E10": "Usa Checklist, Invitados y Resumen para avanzar.",
        "G7": "Fecha de boda", "G8": "Presupuesto objetivo", "G9": "Invitados previstos", "G10": "Pareja A", "G11": "Pareja B",
        "B15": "Resumen", "C15": "Visión general del presupuesto, pagos, pendientes e invitados confirmados.",
        "B16": "Presupuesto", "C16": "Plan anual por categorías y calendario de desembolsos.",
        "B17": "Pagos reales", "C17": "Importes pagados y comparativa frente al plan.",
        "B18": "Proveedores", "C18": "Registro de contratos, pagos, responsables y estados.",
        "B19": "Flujo de pagos", "C19": "Lectura mensual de presupuesto, pagos y saldo previsto.",
        "B20": "Checklist", "C20": "Tareas clave con responsables, fechas y progreso.",
        "B21": "Invitados", "C21": "Conteo por grupo, confirmación, menú y mesa.",
        "B22": "Listas", "C22": "Parámetros y catálogos editables del libro.",
        "B25": "Consejo: duplica el archivo para cada boda y cambia primero las celdas de Listas. La estructura, la navegación y las fórmulas ya están preparadas para crecer sin perder el encuadre.",
    }
    for address, value in values.items():
        set_value(sheet, address, value)
    set_formula(sheet, "H7", "='Listas'!$C$8")
    set_formula(sheet, "H8", "='Listas'!$C$9")
    set_formula(sheet, "H9", "='Listas'!$C$10")
    set_formula(sheet, "H10", "='Listas'!$C$11")
    set_formula(sheet, "H11", "='Listas'!$C$12")
    sheet.Range("H7").NumberFormat = "dd/mm/aaaa"
    sheet.Range("H8").NumberFormat = "#,##0.00"
    sheet.Range("H9").NumberFormat = "0"


def build_listas(wb) -> None:
    sheet = wb.Worksheets("Listas")
    set_frame(sheet, "Parámetros y catálogos de la boda", "Cambia aquí los valores de entrada; las listas alimentan las validaciones del resto del libro.")
    clear_contents(sheet, "B6:M21")
    fixed = {
        "B6": "PARÁMETROS DE LA BODA", "E6": "MESES", "G6": "TIPOS", "I6": "CATEGORÍAS", "K6": "RESPONSABLES", "M6": "ESTADOS",
        "B7": "Parámetro", "C7": "Valor", "B8": "Fecha de boda", "B9": "Presupuesto objetivo", "B10": "Invitados previstos",
        "B11": "Pareja A", "B12": "Pareja B", "B13": "Moneda", "C8": datetime(2027, 6, 19), "C9": 18000,
        "C10": 120, "C11": "Laura", "C12": "Álvaro", "C13": "EUR",
        "E21": "Las celdas de entrada definen la boda. Mantén visibles estas listas para adaptar categorías, responsables y estados sin romper las fórmulas.",
    }
    for address, value in fixed.items():
        set_value(sheet, address, value)
    for index, value in enumerate(MONTHS, 7):
        set_value(sheet, f"E{index}", value)
    for index, value in enumerate(TYPES, 7):
        set_value(sheet, f"G{index}", value)
    for index, value in enumerate(CATEGORIES, 7):
        set_value(sheet, f"I{index}", value)
    for index, value in enumerate(PAYERS, 7):
        set_value(sheet, f"K{index}", value)
    for index, value in enumerate(STATES, 7):
        set_value(sheet, f"M{index}", value)
    sheet.Range("C8").NumberFormat = "dd/mm/aaaa"
    sheet.Range("C9").NumberFormat = "#,##0.00"
    sheet.Range("C10").NumberFormat = "0"


def build_presupuesto(wb) -> None:
    sheet = wb.Worksheets("Presupuesto")
    set_frame(sheet, "Presupuesto de boda por categoría", "Distribuye el coste previsto y conserva la lectura mensual para anticipar los desembolsos.")
    clear_contents(sheet, "B6:O18")
    set_value(sheet, "B6", "PRESUPUESTO DE BODA · CELDAS DE COLOR = ENTRADA")
    for col, header in enumerate(["Categoría", *MONTHS, "Total"], 2):
        set_value(sheet, f"{column_letter(col)}7", header)
    values = [
        ["Lugar y ceremonia", 0, 0, 2000, 0, 1500, 0, 0, 0, 0, 0, 0, 0],
        ["Banquete", 0, 0, 0, 0, 0, 5000, 0, 0, 0, 0, 0, 0],
        ["Fotografía y vídeo", 0, 0, 0, 0, 0, 1800, 0, 0, 0, 0, 0, 0],
        ["Música", 0, 0, 0, 0, 0, 1200, 0, 0, 0, 0, 0, 0],
        ["Decoración y flores", 0, 0, 0, 0, 0, 1600, 0, 0, 0, 0, 0, 0],
        ["Vestuario y belleza", 300, 800, 300, 0, 0, 600, 0, 0, 0, 0, 0, 0],
        ["Papelería y detalles", 250, 250, 350, 350, 400, 500, 0, 0, 0, 0, 0, 0],
    ]
    for row, values_row in enumerate(values, 8):
        set_value(sheet, f"B{row}", values_row[0])
        for col, value in enumerate(values_row[1:], 3):
            set_value(sheet, f"{column_letter(col)}{row}", value)
        set_formula(sheet, f"O{row}", f"=SUM(C{row}:N{row})")
    set_value(sheet, "B15", "TOTAL DE LA BODA")
    for col in range(3, 16):
        letter = column_letter(col)
        set_formula(sheet, f"{letter}15", f"=SUM({letter}8:{letter}14)")
    set_value(sheet, "B18", "La matriz conserva el patrón mensual/anual de la base: categorías en filas, meses en columnas y total calculado a la derecha. Puedes cambiar los importes sin mover el encuadre.")


def build_pagos_reales(wb) -> None:
    sheet = wb.Worksheets("Pagos reales")
    set_frame(sheet, "Pagos reales de proveedores", "Cada importe se obtiene del registro de proveedores para comparar lo previsto con lo ya pagado.")
    clear_contents(sheet, "B6:O18")
    set_value(sheet, "B6", "PAGOS REALES · FÓRMULAS DESDE PROVEEDORES")
    for col, header in enumerate(["Categoría", *MONTHS, "Total"], 2):
        set_value(sheet, f"{column_letter(col)}7", header)
    for row, source_row in enumerate(range(8, 15), 8):
        set_formula(sheet, f"B{row}", f"='Presupuesto'!B{source_row}")
        for col in range(3, 15):
            letter = column_letter(col)
            set_formula(sheet, f"{letter}{row}", f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Pago",\'Proveedores\'!$E$8:$E$200,$B{row},\'Proveedores\'!$C$8:$C$200,{letter}$7)')
        set_formula(sheet, f"O{row}", f"=SUM(C{row}:N{row})")
    set_value(sheet, "B15", "TOTAL PAGADO")
    for col in range(3, 16):
        letter = column_letter(col)
        set_formula(sheet, f"{letter}15", f"=SUM({letter}8:{letter}14)")
    set_value(sheet, "B18", "La hoja sigue el patrón de presupuesto frente a real, pero centraliza los pagos en Proveedores para evitar duplicidades y facilitar la revisión de contratos.")


def build_proveedores(wb) -> None:
    sheet = wb.Worksheets("Proveedores")
    set_frame(sheet, "Registro de proveedores y pagos", "Anota cada contrato o desembolso una sola vez; las hojas de análisis se alimentan de este diario.")
    clear_contents(sheet, "B6:O27")
    set_value(sheet, "B6", "REGISTRO OPERATIVO · PROVEEDORES Y PAGOS")
    headers = ["Fecha", "Mes", "Tipo", "Categoría", "Concepto", "Importe", "Pagado por", "Estado"]
    for col, header in enumerate(headers, 2):
        set_value(sheet, f"{column_letter(col)}7", header)
    samples = [
        [datetime(2027, 1, 10), "Enero", "Pago", "Lugar y ceremonia", "Reserva del espacio", 2000, "Pareja", "Pagado"],
        [datetime(2027, 2, 14), "Febrero", "Pago", "Vestuario y belleza", "Vestido y pruebas", 800, "Pareja", "Pagado"],
        [datetime(2027, 3, 4), "Marzo", "Pago", "Papelería y detalles", "Invitaciones", 350, "Familia", "Pendiente"],
        [datetime(2027, 4, 2), "Abril", "Pago", "Fotografía y vídeo", "Reserva fotógrafo", 600, "Pareja", "Pendiente"],
        [datetime(2027, 5, 20), "Mayo", "Ingreso", "Banquete", "Aportación familiar", 1500, "Familia", "Confirmado"],
        [datetime(2027, 6, 19), "Junio", "Ajuste", "Decoración y flores", "Fondo de imprevistos", 300, "Pareja", "Pendiente"],
    ]
    for row in range(8, 28):
        values = samples[row - 8] if row - 8 < len(samples) else [None] * 8
        for col, value in enumerate(values, 2):
            set_value(sheet, f"{column_letter(col)}{row}", value)
    set_value(sheet, "K6", "REGLAS DEL DIARIO")
    set_value(sheet, "K7", "Una fila = un contrato, pago o ajuste. Usa Tipo y Estado para que los resúmenes distingan lo comprometido de lo abonado.")
    set_value(sheet, "K12", "Las listas desplegables mantienen consistentes categorías, responsables y estados en todo el libro.")
    sheet.Range("B8:B27").NumberFormat = "dd/mm/aaaa"
    sheet.Range("G8:G27").NumberFormat = "#,##0.00"


def build_flujo_pagos(wb) -> None:
    sheet = wb.Worksheets("Flujo de pagos")
    set_frame(sheet, "Flujo mensual de pagos", "La lectura anual combina presupuesto, ingresos, pagos reales, ajustes y saldo previsto.")
    clear_contents(sheet, "B6:J23")
    set_value(sheet, "B6", "LECTURA MENSUAL DE LA BODA")
    headers = ["Mes", "Presupuesto plan", "Ingresos", "Pagos previstos", "Pagos reales", "Ajustes", "Saldo previsto", "Comentarios", "Estado"]
    for col, header in enumerate(headers, 2):
        set_value(sheet, f"{column_letter(col)}7", header)
    for row, month in enumerate(MONTHS, 8):
        letter = column_letter(row - 5)
        set_value(sheet, f"B{row}", month)
        set_formula(sheet, f"C{row}", f"=SUM('Presupuesto'!{letter}$8:{letter}$14)")
        set_formula(sheet, f"D{row}", f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Ingreso",\'Proveedores\'!$C$8:$C$200,$B{row})')
        set_formula(sheet, f"E{row}", f"=C{row}")
        set_formula(sheet, f"F{row}", f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Pago",\'Proveedores\'!$C$8:$C$200,$B{row})')
        set_formula(sheet, f"G{row}", f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Ajuste",\'Proveedores\'!$C$8:$C$200,$B{row})')
        set_formula(sheet, f"H{row}", f"=C{row}-F{row}+D{row}-G{row}")
    set_value(sheet, "B20", "TOTAL ANUAL")
    for col in range(3, 9):
        letter = column_letter(col)
        set_formula(sheet, f"{letter}20", f"=SUM({letter}8:{letter}19)")
    sheet.Range("C8:H20").NumberFormat = "#,##0.00"
    set_value(sheet, "B23", "El saldo previsto no sustituye una revisión contractual: sirve para ver cuándo se concentran los desembolsos y qué pagos quedan pendientes.")
    if sheet.ChartObjects().Count:
        sheet.ChartObjects(1).Chart.HasTitle = True
        sheet.ChartObjects(1).Chart.ChartTitle.Text = "Ingresos y pagos de la boda"


def build_resumen(wb) -> None:
    sheet = wb.Worksheets("Resumen")
    set_frame(sheet, "El estado de la boda en una sola pantalla", "Panel de decisión para presupuesto, pagos, pendientes e invitados confirmados.")
    clear_contents(sheet, "B5:M22")
    set_value(sheet, "B5", "Fecha de boda")
    set_formula(sheet, "C5", "='Listas'!$C$8")
    sheet.Range("C5").NumberFormat = "dd/mm/aaaa"
    cards = {
        "B6": "PRESUPUESTO PLANIFICADO", "E6": "PAGOS REALES", "H6": "PENDIENTE", "K6": "INVITADOS CONFIRMADOS",
    }
    for address, value in cards.items():
        set_value(sheet, address, value)
    set_formula(sheet, "B7", "=SUM('Presupuesto'!C8:N14)")
    set_formula(sheet, "E7", '=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Pago")')
    set_formula(sheet, "H7", "=B7-E7")
    set_formula(sheet, "K7", "='Invitados'!D11")
    sheet.Range("B7").NumberFormat = "#,##0.00"
    sheet.Range("E7:H7").NumberFormat = "#,##0.00"
    sheet.Range("K7").NumberFormat = "0"
    set_value(sheet, "B10", "PRESUPUESTO Y PAGOS POR CATEGORÍA")
    for col, header in enumerate(["Categoría", "Plan", "Pagado", "Pendiente", "% pagado"], 2):
        set_value(sheet, f"{column_letter(col)}11", header)
    for row, source_row in enumerate(range(8, 15), 12):
        set_formula(sheet, f"B{row}", f"='Presupuesto'!B{source_row}")
        set_formula(sheet, f"C{row}", f"=SUM('Presupuesto'!C{source_row}:N{source_row})")
        set_formula(sheet, f"D{row}", f'=SUMIFS(\'Proveedores\'!$G$8:$G$200,\'Proveedores\'!$D$8:$D$200,"Pago",\'Proveedores\'!$E$8:$E$200,$B{row})')
        set_formula(sheet, f"E{row}", f"=C{row}-D{row}")
        set_formula(sheet, f"F{row}", f"=IFERROR(D{row}/C{row},0)")
    set_value(sheet, "B19", "TOTAL")
    set_formula(sheet, "C19", "=SUM(C12:C18)")
    set_formula(sheet, "D19", "=SUM(D12:D18)")
    set_formula(sheet, "E19", "=SUM(E12:E18)")
    set_formula(sheet, "F19", "=IFERROR(D19/C19,0)")
    set_value(sheet, "B22", "La diferencia es Plan - Pagado: positiva = aún queda presupuesto; negativa = la categoría ha superado lo previsto.")
    if sheet.ChartObjects().Count:
        sheet.ChartObjects(1).Chart.HasTitle = True
        sheet.ChartObjects(1).Chart.ChartTitle.Text = "Presupuesto frente a pagos"


def build_checklist(wb) -> None:
    sheet = wb.Worksheets("Checklist")
    set_frame(sheet, "Checklist de la boda", "Convierte las decisiones importantes en tareas con responsable, fecha y estado.")
    clear_contents(sheet, "B6:J16")
    set_value(sheet, "B6", "TAREAS Y MOMENTOS CLAVE")
    for col, header in enumerate(["Tarea", "Área", "Responsable", "Fecha límite", "Estado", "Prioridad", "Días", "Notas", "Progreso"], 2):
        set_value(sheet, f"{column_letter(col)}9", header)
    rows = [
        ["Cerrar contrato del lugar", "Lugar", "Pareja", datetime(2027, 1, 31), "Confirmado", "Alta", "=MAX(E10-TODAY(),0)", "Revisar horarios y montaje", '=IF(F10="Confirmado",1,0)'],
        ["Enviar invitaciones", "Invitados", "Familia", datetime(2027, 3, 15), "Pendiente", "Alta", "=MAX(E11-TODAY(),0)", "Lista final pendiente", '=IF(F11="Confirmado",1,0)'],
        ["Confirmar menú y música", "Celebración", "Pareja", datetime(2027, 5, 20), "Pendiente", "Media", "=MAX(E12-TODAY(),0)", "Cita con proveedores", '=IF(F12="Confirmado",1,0)'],
    ]
    for row, values in enumerate(rows, 10):
        for col, value in enumerate(values, 2):
            address = f"{column_letter(col)}{row}"
            if isinstance(value, str) and value.startswith("="):
                set_formula(sheet, address, value)
            else:
                set_value(sheet, address, value)
    sheet.Range("E10:E12").NumberFormat = "dd/mm/aaaa"
    sheet.Range("H10:H12").NumberFormat = "0"
    sheet.Range("J10:J12").NumberFormat = "0%"
    set_value(sheet, "B16", "El checklist comparte la misma retícula de seguimiento de la base, pero cambia metas financieras por decisiones concretas de la boda.")


def build_invitados(wb) -> None:
    sheet = wb.Worksheets("Invitados")
    set_frame(sheet, "Invitados y mesas", "Controla grupos, confirmaciones, preferencias de menú y distribución de mesas.")
    clear_contents(sheet, "B6:H14")
    set_value(sheet, "B6", "LISTA DE INVITADOS")
    for col, header in enumerate(["Grupo", "Invitados previstos", "Confirmados", "Menú", "Mesa", "Estado", "Notas"], 2):
        set_value(sheet, f"{column_letter(col)}7", header)
    rows = [
        ["Familia pareja A", 35, 30, "Estándar", 1, "Confirmado", "Pendientes 5 respuestas"],
        ["Familia pareja B", 30, 26, "Vegetariano", 2, "Pendiente", "Revisar teléfonos"],
        ["Amigos", 55, 42, "Mixto", 3, "Pendiente", "Confirmar acompañantes"],
    ]
    for row, values in enumerate(rows, 8):
        for col, value in enumerate(values, 2):
            set_value(sheet, f"{column_letter(col)}{row}", value)
    set_value(sheet, "B11", "TOTAL")
    set_formula(sheet, "C11", "=SUM(C8:C10)")
    set_formula(sheet, "D11", "=SUM(D8:D10)")
    sheet.Range("C8:D11").NumberFormat = "0"
    sheet.Range("F8:F10").NumberFormat = "0"
    set_value(sheet, "B14", "Las cifras de confirmados alimentan la tarjeta del Resumen y ayudan a decidir menú, mesas y cantidad de papelería.")


def column_letter(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def rename_tables(wb) -> None:
    names = {
        "tblMovimientos": "tblProveedores",
        "tblFlujoCaja": "tblFlujoPagos",
        "tblMetas": "tblChecklist",
        "tblDeudas": "tblInvitados",
        "tblGastosReales": "tblPagosReales",
    }
    for sheet in wb.Worksheets:
        for index in range(1, sheet.ListObjects.Count + 1):
            table = sheet.ListObjects(index)
            if table.Name in names:
                table.Name = names[table.Name]


def resize_tables(wb) -> None:
    """Keep totals outside the data tables and give every table a stable schema."""
    targets = {
        "tblResumenCategorias": ("Resumen", "B11:F18"),
        "tblPresupuesto": ("Presupuesto", "B7:O14"),
        "tblPagosReales": ("Pagos reales", "B7:O14"),
        "tblProveedores": ("Proveedores", "B7:I27"),
        "tblFlujoPagos": ("Flujo de pagos", "B7:J19"),
        "tblChecklist": ("Checklist", "B9:J12"),
        "tblInvitados": ("Invitados", "B7:H10"),
        "tblParametros": ("Listas", "B7:C13"),
    }
    for name, (sheet_name, address) in targets.items():
        sheet = wb.Worksheets(sheet_name)
        table = sheet.ListObjects(name)
        table.Resize(sheet.Range(address))


def add_validation(sheet, cell_range: str, formula: str) -> None:
    target = sheet.Range(cell_range)
    try:
        target.Validation.Delete()
    except Exception:
        pass
    target.Validation.Add(3, 1, 1, formula)
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
    target.Validation.ErrorTitle = "Valor no válido"
    target.Validation.ErrorMessage = "Selecciona un valor de la lista."
    target.Validation.InputTitle = "Lista"
    target.Validation.InputMessage = "Usa una opción de las listas del libro."


def clear_generated_shapes(sheet) -> None:
    for index in range(sheet.Shapes.Count, 0, -1):
        shape = sheet.Shapes.Item(index)
        name = shape.Name
        if name == "FamilyBudgetIcon" or name.startswith("Nav_"):
            shape.Delete()


def add_navigation_and_icon(wb) -> None:
    items = [(sheet.Name, sheet.Name if len(sheet.Name) <= 12 else sheet.Name[:11] + "…") for sheet in wb.Worksheets]
    for sheet in wb.Worksheets:
        clear_generated_shapes(sheet)
        sheet.Rows(1).RowHeight = 26
        start = sheet.Range("B1").Left
        top = sheet.Range("B1").Top + 2
        for index, (target, label) in enumerate(items):
            button = sheet.Shapes.AddShape(5, start + index * 72, top, 68, 21)
            button.Name = f"Nav_{target.replace(' ', '_')}"
            button.Placement = 1
            button.Fill.ForeColor.RGB = rgb(PALETTE["NAVY"] if sheet.Name == target else PALETTE["MINT"])
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
            text.TextRange.Font.Fill.ForeColor.RGB = rgb(PALETTE["WHITE"] if sheet.Name == target else PALETTE["NAVY"])
            sheet.Hyperlinks.Add(button, "", f"'{target}'!A1")
        used = sheet.UsedRange
        last_column = used.Column + used.Columns.Count - 1
        anchor = sheet.Cells(2, last_column)
        left = anchor.Left + max(anchor.Width - 28, 2)
        picture = sheet.Shapes.AddPicture(str(ICON), 0, -1, left, anchor.Top + 1, 26, 26)
        picture.Name = "FamilyBudgetIcon"
        picture.Placement = 1


def prepare_workbook(wb) -> None:
    for old, new in SHEET_RENAMES.items():
        wb.Worksheets(old).Name = new
    retheme_workbook(wb)
    build_inicio(wb)
    build_listas(wb)
    build_presupuesto(wb)
    build_pagos_reales(wb)
    build_proveedores(wb)
    build_flujo_pagos(wb)
    build_resumen(wb)
    build_checklist(wb)
    build_invitados(wb)
    rename_tables(wb)
    resize_tables(wb)
    # Excel tables can propagate a calculated-column formula from the old
    # model when their range changes. Reapply every formula-backed screen after
    # the final table schema is fixed so each row keeps its intended references.
    build_presupuesto(wb)
    build_pagos_reales(wb)
    build_flujo_pagos(wb)
    build_resumen(wb)
    # Final resize must happen after formula writes; otherwise Excel can
    # auto-expand a table to absorb an adjacent TOTAL row.
    resize_tables(wb)
    add_validation(wb.Worksheets("Proveedores"), "C8:C200", "=Listas!$E$7:$E$18")
    add_validation(wb.Worksheets("Proveedores"), "D8:D200", "=Listas!$G$7:$G$9")
    add_validation(wb.Worksheets("Proveedores"), "E8:E200", "=Listas!$I$7:$I$13")
    add_validation(wb.Worksheets("Proveedores"), "H8:H200", "=Listas!$K$7:$K$9")
    add_validation(wb.Worksheets("Proveedores"), "I8:I200", "=Listas!$M$7:$M$9")
    add_validation(wb.Worksheets("Checklist"), "D10:D100", "=Listas!$K$7:$K$9")
    add_validation(wb.Worksheets("Checklist"), "F10:F100", "=Listas!$M$7:$M$9")
    add_validation(wb.Worksheets("Invitados"), "G8:G100", "=Listas!$M$7:$M$9")
    for sheet in wb.Worksheets:
        try:
            sheet.Rows(7).Borders(9).LineStyle = 0
        except Exception:
            pass
    add_navigation_and_icon(wb)
    try:
        wb.BuiltinDocumentProperties("Title").Value = "Planificador de boda · base nativa"
        wb.BuiltinDocumentProperties("Subject").Value = "Presupuesto, proveedores, pagos, invitados y checklist de boda"
        wb.BuiltinDocumentProperties("Author").Value = "Creador de Excels"
    except Exception:
        pass


def build_workbook(output: Path) -> Path:
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE, output)
    retheme_package(output)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(output.resolve()), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        try:
            excel.AutoCorrect.AutoFillFormulasInLists = False
        except Exception:
            pass
        prepare_workbook(workbook)
        excel.Calculation = -4105
        excel.CalculateFullRebuild()
        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
    expected_sheets = [
        "Inicio",
        "Resumen",
        "Presupuesto",
        "Pagos reales",
        "Proveedores",
        "Flujo de pagos",
        "Checklist",
        "Invitados",
        "Listas",
    ]
    contract = wedding_base_contract()
    artifact_dir = output.parent.parent
    behavior_path = artifact_dir / "model_behavior.md"
    write_model_behavior_markdown(behavior_path, contract)
    write_model_contract_json(artifact_dir / "model_contract.json", contract)
    quality = validate_model_file(
        output,
        expected_sheets=expected_sheets,
        model_contract=contract,
        model_contract_path=behavior_path,
    )
    if not quality["passed"]:
        raise RuntimeError(f"El modelo no supera la validación global: {quality['errors']}")
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description="Crea un planificador de boda heredando la base nativa aprobada")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    output = args.output if args.output.is_absolute() else ROOT / args.output
    print(f"WORKBOOK={build_workbook(output)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
