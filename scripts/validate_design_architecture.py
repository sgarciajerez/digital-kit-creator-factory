from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

from design_architecture import ROOT, load_design_system, validate_design_spec
from model_quality import validate_model_behavior_markdown
from model_contracts import contract_for_product


def active_product_ids() -> list[int]:
    pipeline = json.loads((ROOT / "config" / "FACTORY_PIPELINE.json").read_text(encoding="utf-8"))
    return [int(value) for value in pipeline.get("active_product_ids", [])]


def validate_product(product_id: int, write_report: bool = True) -> tuple[bool, dict]:
    product = ROOT / "outputs" / "listings" / f"product-{product_id}"
    spec_path = product / "design_spec.json"
    errors: list[str] = []
    spec: dict = {}
    if not spec_path.is_file():
        errors.append("falta design_spec.json")
    else:
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        errors.extend(validate_design_spec(spec, load_design_system()))
        contract_path = product / str(spec.get("model_contract_file", "model_contract.json"))
        behavior_path = product / str(spec.get("model_behavior_file", "model_behavior.md"))
        if not contract_path.is_file() or not behavior_path.is_file():
            errors.append("falta el contrato funcional o model_behavior.md")
        elif spec.get("model_contract"):
            behavior_report = validate_model_behavior_markdown(behavior_path, contract_for_product(product_id))
            errors.extend(behavior_report.get("errors", []))
        master_file = ROOT / str(spec.get("master_workbook", "templates/excel-first-base/master_workbook.xlsx"))
        if not master_file.is_file():
            errors.append("falta templates/excel-first-base/master_workbook.xlsx")
    report = {
        "schema_version": "2.0.0",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "product_id": product_id,
        "status": "passed" if not errors else "failed",
        "errors": errors,
        "templates": {tab: data.get("template") for tab, data in (spec.get("screen_contracts") or {}).items()},
        "source_lineage": spec.get("source_lineage", []),
        "model_base": spec.get("model_base"),
        "next_stage": "app_generation" if not errors else "repair_native_design_contract",
    }
    if write_report:
        product.mkdir(parents=True, exist_ok=True)
        (product / "design_architecture_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return not errors, report


def main() -> int:
    parser = argparse.ArgumentParser(description="Valida la arquitectura nativa de la factoria")
    parser.add_argument("--product", type=int, help="Validar un solo producto")
    args = parser.parse_args()
    product_ids = [args.product] if args.product else active_product_ids()
    failed = False
    for product_id in product_ids:
        passed, report = validate_product(product_id)
        failed = failed or not passed
        print(f"product-{product_id}: DESIGN_ARCHITECTURE={report['status'].upper()}")
        for error in report["errors"]:
            print(f"  - {error}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
