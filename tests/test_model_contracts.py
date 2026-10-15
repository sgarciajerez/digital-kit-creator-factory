from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.table import Table

from model_contracts import render_model_behavior_markdown
from model_quality import validate_model_contract, validate_model_behavior_markdown


def _contract():
    return {
        "contract_id": "test-v1",
        "title": "Modelo de prueba",
        "purpose": "Registrar importes y calcular un saldo.",
        "sheets": ["Data"],
        "tables": [{
            "sheet": "Data",
            "ref": "A1:C2",
            "headers": [
                {"name": "Item", "role": "input", "kind": "text"},
                {"name": "Amount", "role": "input", "kind": "amount"},
                {"name": "Balance", "role": "derived", "kind": "amount"},
            ],
            "required": True,
        }],
        "formula_rules": [{
            "sheet": "Data",
            "cell": "C2",
            "contains": "B2",
            "description": "Balance follows Amount.",
        }],
        "relationships": ["Balance follows Amount."],
        "cases": [{
            "name": "balance changes",
            "purpose": "Comprueba la conexión principal.",
            "inputs": [{"sheet": "Data", "cell": "B2", "value": 10}],
            "assertions": [{"sheet": "Data", "cell": "C2", "operator": "equals", "value": 10}],
        }],
        "failure_criteria": ["La fórmula de Balance desaparece."],
    }


def test_model_behavior_guide_is_normative_and_contains_expected_case(tmp_path: Path) -> None:
    contract = _contract()
    path = tmp_path / "model_behavior.md"
    path.write_text(render_model_behavior_markdown(contract), encoding="utf-8")

    report = validate_model_behavior_markdown(path, contract)

    assert report["passed"] is True
    assert "error observado" in path.read_text(encoding="utf-8")


def test_model_contract_fails_when_table_schema_or_formula_does_not_match(tmp_path: Path) -> None:
    contract = _contract()
    path = tmp_path / "model_behavior.md"
    path.write_text(render_model_behavior_markdown(contract), encoding="utf-8")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["Item", "Amount", "Wrong header"])
    sheet.append(["Example", 10, 0])
    sheet.add_table(Table(displayName="tblData", ref="A1:C2"))

    report = validate_model_contract(workbook, contract, path)

    assert report["passed"] is False
    assert any("cabeceras reales" in error for error in report["errors"])
    assert any("falta fórmula" in error for error in report["errors"])

