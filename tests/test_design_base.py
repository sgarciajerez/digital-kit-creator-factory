from pathlib import Path

from validate_design_base import validate_base


def test_reusable_design_base_matches_the_generated_workbook(tmp_path: Path) -> None:
    root = Path(__file__).resolve().parent.parent
    base = root / "templates" / "excel-first-base"
    report = validate_base(
        base / "design_spec.json",
        base / "master_workbook.xlsx",
        base / "master_workbook.xlsx",
        tmp_path / "design_contract_report.json",
    )
    assert report["status"] == "passed"
    assert report["checks"]["no_freeze_panes"] is True
    assert report["workbook"]["formula_count"] >= 12
    assert report["workbook"]["validation_count"] >= 1
    assert report["checks"]["chart_contract"] is True
