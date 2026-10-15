from __future__ import annotations

import argparse
import json
from pathlib import Path

from excel_template_factory import ROOT, create_template_workbook
from factory_config import all_product_ids
from model_quality import validate_model_file
from model_contracts import contract_for_product, contract_for_slug, load_model_contract, write_app_usage_markdown, write_model_behavior_markdown, write_model_contract_json


def main() -> int:
    parser = argparse.ArgumentParser(description="Crea un XLSX desde el master_workbook y el contrato Excel-first")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--product", type=int, choices=all_product_ids())
    source.add_argument("--spec", type=Path, help="Contrato JSON de diseño para crear un XLSX derivado")
    parser.add_argument("--output", type=Path, help="Ruta destino opcional")
    parser.add_argument(
        "--functional-qa",
        action="store_true",
        help="Ejecuta los casos funcionales declarados usando Excel instalado",
    )
    args = parser.parse_args()
    if args.spec:
        spec_path = args.spec if args.spec.is_absolute() else ROOT / args.spec
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        output = args.output or spec_path.parent / "master_workbook-copy.xlsx"
    else:
        folder = ROOT / "outputs" / "listings" / f"product-{args.product}"
        spec = json.loads((folder / "design_spec.json").read_text(encoding="utf-8"))
        output = args.output or folder / "app" / f"{spec['slug']}-simple-template.xlsx"
    model_contract = spec.get("model_contract")
    if model_contract is None:
        declared_contract = spec.get("model_contract_file")
        if declared_contract:
            candidate = (spec_path.parent if args.spec else output.parent.parent) / str(declared_contract)
            if candidate.is_file():
                model_contract = load_model_contract(candidate)
    if model_contract is None and spec.get("product_id") is not None:
        try:
            model_contract = contract_for_product(int(spec["product_id"]))
        except KeyError:
            model_contract = None
    if model_contract is None:
        model_contract = contract_for_slug(str(spec.get("slug", "")))
    model_behavior_path = spec_path.parent / "model_behavior.md" if args.spec else output.parent.parent / "model_behavior.md"
    # The master workbook is the only visual and structural base. The final
    # product builder is the stage that must pass the product model contract.
    enforce_model_contract = True
    if model_contract and enforce_model_contract:
        write_model_behavior_markdown(model_behavior_path, model_contract)
        write_model_contract_json(model_behavior_path.with_name("model_contract.json"), model_contract)
    output.parent.mkdir(parents=True, exist_ok=True)
    guide_markdown = write_app_usage_markdown(output.parent / "README.md", spec, model_contract or {})
    guide_filename = str(spec.get("app_user_guide_pdf_filename", f"{spec.get('slug', 'excel')}-user-guide.pdf"))
    if not guide_filename.lower().endswith(".pdf") or Path(guide_filename).name != guide_filename:
        raise ValueError("app_user_guide_pdf_filename debe ser un nombre PDF sin rutas")
    from product_guide_pdf import generate_user_guide_pdf
    generate_user_guide_pdf(guide_markdown, output.parent / guide_filename, spec)
    create_template_workbook(spec, spec.get("content_seed")).save(output)
    model_report = validate_model_file(
        output,
        expected_sheets=spec.get("tabs", []),
        functional_contract=spec.get("functional_contract"),
        run_functional=args.functional_qa,
        model_contract=model_contract,
        model_contract_path=model_behavior_path if model_contract else None,
    )
    report_path = output.with_name(f"{output.stem}.model_quality.json")
    report_path.write_text(json.dumps(model_report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if not model_report["passed"]:
        print(f"MODEL_QUALITY=FAILED report={report_path}")
        return 1
    print(f"TEMPLATE_XLSX={output}")
    print(f"MODEL_QUALITY=PASSED report={report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
