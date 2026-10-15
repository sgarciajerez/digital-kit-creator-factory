from __future__ import annotations

import hashlib
import json
import re
import time
from pathlib import Path
from typing import Literal

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pydantic import BaseModel, ConfigDict, Field

from factory_core import DATABASE_PATH, PROJECT_ROOT, create_product, update_product
from model_quality import validate_model_file


class ProductBrief(BaseModel):
    model_config = ConfigDict(extra="forbid")

    family: Literal["budget_planner", "family_budget_planner"] = "budget_planner"
    name: str = Field(min_length=3)
    audience: str = Field(min_length=3)
    problem: str = Field(min_length=5)
    currency: str = Field(default="EUR", min_length=3, max_length=3)
    categories: list[str] = Field(min_length=1)
    market_research: dict | None = None


class QualityReport(BaseModel):
    passed: bool
    score: float = Field(ge=0, le=100)
    checks: list[str]
    errors: list[str]
    sha256: str | None = None


DEFAULT_BRIEF = ProductBrief(
    name="Planificador de presupuesto personal",
    audience="Personas que quieren organizar gastos mensuales sin usar macros",
    problem="Necesitan comparar presupuesto y gasto real de forma sencilla y revisable",
    categories=["Vivienda", "Alimentación", "Transporte", "Salud", "Ocio", "Otros"],
)


def _header_style(cell) -> None:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center")


def generate_budget_workbook(output_path: Path, brief: ProductBrief) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    instructions = workbook.active
    instructions.title = "Instrucciones"
    budget = workbook.create_sheet("Presupuesto")
    preview = workbook.create_sheet("Vista previa")

    instructions.append([brief.name])
    instructions["A1"].font = Font(size=16, bold=True)
    instructions.append(["Producto ficticio de prueba; sustituye los ejemplos por tus datos."])
    instructions.append([])
    instructions.append(["Pasos"])
    _header_style(instructions["A4"])
    for number, text in enumerate(
        [
            "Introduce el presupuesto y el gasto real por categoría en la hoja Presupuesto.",
            "Revisa la desviación y las notas antes de tomar decisiones.",
            "La hoja Vista previa resume los totales; no sustituye asesoramiento financiero.",
        ],
        start=5,
    ):
        instructions.cell(number, 1, f"{number - 4}. {text}")
    instructions.column_dimensions["A"].width = 110

    headers = ["Categoría", "Presupuesto", "Gasto real", "Desviación", "Notas"]
    budget.append(headers)
    for cell in budget[1]:
        _header_style(cell)
    for row_number, category in enumerate(brief.categories, start=2):
        budget.cell(row_number, 1, category)
        budget.cell(row_number, 2, 0)
        budget.cell(row_number, 3, 0)
        budget.cell(row_number, 4, f"=B{row_number}-C{row_number}")
        for column in (2, 3, 4):
            budget.cell(row_number, column).number_format = "#,##0.00"
    total_row = len(brief.categories) + 2
    budget.cell(total_row, 1, "TOTAL")
    budget.cell(total_row, 2, f"=SUM(B2:B{total_row - 1})")
    budget.cell(total_row, 3, f"=SUM(C2:C{total_row - 1})")
    budget.cell(total_row, 4, f"=SUM(D2:D{total_row - 1})")
    for cell in budget[total_row]:
        cell.font = Font(bold=True)
    for column in (2, 3, 4):
        budget.cell(total_row, column).number_format = "#,##0.00"
    for column, width in {"A": 24, "B": 16, "C": 16, "D": 16, "E": 44}.items():
        budget.column_dimensions[column].width = width
    budget.freeze_panes = "B2"
    budget.auto_filter.ref = f"A1:E{total_row - 1}"

    if brief.family == "family_budget_planner":
        settings = workbook.create_sheet("Configuración")
        settings.append(["Configuración del hogar", "Valor"])
        for cell in settings[1]:
            _header_style(cell)
        settings.append(["Miembro A", ""])
        settings.append(["Miembro B", ""])
        settings.append(["Regla de reparto", "50/50 o personalizada"])
        settings.append(["Nota", "No constituye asesoramiento financiero"])
        settings.column_dimensions["A"].width = 24
        settings.column_dimensions["B"].width = 34

        shared = workbook.create_sheet("Gastos compartidos")
        shared.append(
            [
                "Fecha",
                "Concepto",
                "Categoría",
                "Pagado por",
                "Importe",
                "Parte miembro A",
                "Parte miembro B",
                "Saldo A/B",
            ]
        )
        for cell in shared[1]:
            _header_style(cell)
        for row_number in range(2, 32):
            shared.cell(row_number, 6, f'=IF(E{row_number}="",0,E{row_number}/2)')
            shared.cell(row_number, 7, f'=IF(E{row_number}="",0,E{row_number}/2)')
            shared.cell(row_number, 8, f'=IF(D{row_number}="",0,IF(D{row_number}=Configuración!B2,F{row_number}-G{row_number},G{row_number}-F{row_number}))')
        for row_number in range(2, 32):
            for column in (5, 6, 7, 8):
                shared.cell(row_number, column).number_format = "#,##0.00"
        for column, width in {"A": 14, "B": 28, "C": 22, "D": 20, "E": 14, "F": 18, "G": 18, "H": 14}.items():
            shared.column_dimensions[column].width = width
        shared.freeze_panes = "A2"
        shared.auto_filter.ref = "A1:H31"

    preview.append(["Resumen", "Valor"])
    _header_style(preview[1][0])
    _header_style(preview[1][1])
    preview.append(["Presupuesto total", f"=Presupuesto!B{total_row}"])
    preview.append(["Gasto real total", f"=Presupuesto!C{total_row}"])
    preview.append(["Desviación total", f"=Presupuesto!D{total_row}"])
    preview.append(["Moneda", brief.currency.upper()])
    preview.column_dimensions["A"].width = 26
    preview.column_dimensions["B"].width = 22
    for row_number in (2, 3, 4):
        preview.cell(row_number, 2).number_format = "#,##0.00"

    workbook.properties.title = brief.name
    workbook.properties.subject = "Plantilla de planificación presupuestaria"
    workbook.save(output_path)
    return output_path


def validate_workbook(path: Path, expected_family: str = "budget_planner") -> QualityReport:
    checks: list[str] = []
    errors: list[str] = []
    sha256 = hashlib.sha256(path.read_bytes()).hexdigest() if path.exists() else None
    if not path.exists():
        return QualityReport(passed=False, score=0, checks=[], errors=["No existe el archivo"])

    try:
        workbook = load_workbook(path, data_only=False, read_only=False)
    except Exception as error:
        return QualityReport(passed=False, score=0, checks=[], errors=[f"No se puede abrir: {error}"])

    required_sheets = {"Instrucciones", "Presupuesto", "Vista previa"}
    if expected_family == "family_budget_planner":
        required_sheets |= {"Configuración", "Gastos compartidos"}
    if required_sheets <= set(workbook.sheetnames):
        checks.append("hojas requeridas")
    else:
        errors.append("Faltan hojas requeridas")

    budget = workbook["Presupuesto"]
    if budget.max_row >= 3 and budget["D2"].value == "=B2-C2":
        checks.append("fórmulas de desviación")
    else:
        errors.append("No se encontró la fórmula de desviación")
    if any(
        isinstance(cell.value, str) and cell.value.startswith("=SUM(")
        for row in budget.iter_rows()
        for cell in row
    ):
        checks.append("fórmulas de totales")
    else:
        errors.append("No se encontraron fórmulas de totales")
    if budget.freeze_panes == "B2":
        checks.append("inmovilización de encabezado")
    else:
        errors.append("Falta inmovilizar el encabezado")
    workbook.close()

    model_report = validate_model_file(path, expected_sheets=required_sheets)
    checks.extend(model_report.get("checks", []))
    errors.extend(model_report.get("errors", []))

    score = round(100 * len(checks) / max(len(checks) + len(errors), 1), 2)
    return QualityReport(
        passed=not errors,
        score=score,
        checks=checks,
        errors=errors,
        sha256=sha256,
    )


def create_first_product(
    database_path: Path,
    output_directory: Path,
    brief: ProductBrief = DEFAULT_BRIEF,
    output_filename: str = "planificador_presupuesto_0.1.0.xlsx",
    opportunity_id: int | None = None,
) -> tuple[int, QualityReport, Path]:
    started = time.perf_counter()
    product_id = create_product(
        database_path,
        name=brief.name,
        product_type="xlsx",
        opportunity_id=opportunity_id,
        version="0.1.0",
        status="draft",
        brief_json=brief.model_dump_json(),
    )
    safe_filename = re.sub(r"[^a-zA-Z0-9._-]+", "_", output_filename)
    output_path = output_directory / safe_filename
    try:
        generate_budget_workbook(output_path, brief)
        report = validate_workbook(output_path, expected_family=brief.family)
        update_product(
            database_path,
            product_id,
            output_path=str(output_path),
            quality_score=report.score,
            quality_report=report.model_dump_json(),
            production_time_seconds=round(time.perf_counter() - started),
            status="draft" if report.passed else "needs_review",
        )
    except Exception as error:
        update_product(
            database_path,
            product_id,
            quality_report=json.dumps({"passed": False, "errors": [str(error)]}),
            production_time_seconds=round(time.perf_counter() - started),
            status="error",
        )
        raise
    return product_id, report, output_path


if __name__ == "__main__":
    product_id, report, output_path = create_first_product(
        DATABASE_PATH,
        PROJECT_ROOT / "outputs" / "products",
    )
    print(f"Producto generado: {product_id}")
    print(f"Validación: {'OK' if report.passed else 'REVISAR'} ({report.score:.2f}/100)")
    print(f"Archivo: {output_path}")
