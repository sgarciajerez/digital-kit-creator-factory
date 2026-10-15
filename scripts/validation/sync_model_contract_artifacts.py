from __future__ import annotations

"""Create the normative model guides for bases and existing product sheets."""

import json
from pathlib import Path

from model_contracts import (
    contract_for_product,
    contract_for_slug,
    family_base_contract,
    wedding_base_contract,
    write_model_behavior_markdown,
    write_model_contract_json,
)


ROOT = Path(__file__).resolve().parents[2]


def sync(folder: Path, contract: dict) -> None:
    write_model_behavior_markdown(folder / "model_behavior.md", contract)
    write_model_contract_json(folder / "model_contract.json", contract)
    spec_path = folder / "design_spec.json"
    if spec_path.is_file():
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        spec["model_behavior_file"] = "model_behavior.md"
        spec["model_contract_file"] = "model_contract.json"
        spec["model_behavior_source"] = "declared_model_contract_before_xlsx"
        spec["user_model_guide"] = {
            "file": "model_behavior.md",
            "purpose": "Guía de uso y de verificación funcional independiente del XLSX.",
            "edit_inputs_only": True,
            "fail_on_contract_mismatch": True,
        }
        spec["model_contract"] = contract
        spec_path.write_text(json.dumps(spec, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> int:
    sync(ROOT / "templates" / "excel-first-base", family_base_contract())
    sync(ROOT / "outputs" / "listings" / "family-budget-sources-base", family_base_contract())
    sync(ROOT / "outputs" / "listings" / "wedding-planner-base", wedding_base_contract())
    for product_id in (8, 9, 10, 11):
        folder = ROOT / "outputs" / "listings" / f"product-{product_id}"
        if folder.is_dir():
            sync(folder, contract_for_product(product_id))
    print("MODEL_CONTRACT_ARTIFACTS=SYNCED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

