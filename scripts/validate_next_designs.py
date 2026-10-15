from __future__ import annotations

import json
import re
from pathlib import Path

from factory_config import active_product_ids
from validate_design_architecture import validate_product


ROOT = Path(__file__).resolve().parent.parent
PRODUCT_IDS = active_product_ids()


def main() -> int:
    errors: list[str] = []
    master_references = 0
    markdown_specs = 0
    global_spec = ROOT / "rules" / "FACTORY_NATIVE_SOURCE_WORKFLOW.md"
    if not global_spec.is_file():
        errors.append("Falta FACTORY_NATIVE_SOURCE_WORKFLOW.md")
    for product_id in PRODUCT_IDS:
        folder = ROOT / "outputs" / "listings" / f"product-{product_id}"
        spec_path = folder / "design_spec.json"
        app_spec_path = folder / "app_design_spec.md"
        review_path = folder / "design_review.json"
        if not spec_path.is_file() or not review_path.is_file() or not app_spec_path.is_file():
            errors.append(f"Faltan metadatos de product-{product_id}")
            continue
        markdown_specs += 1
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        architecture_ok, architecture_report = validate_product(product_id)
        if not architecture_ok:
            errors.extend(f"product-{product_id}: {error}" for error in architecture_report["errors"])
        review = json.loads(review_path.read_text(encoding="utf-8"))
        app_markdown = app_spec_path.read_text(encoding="utf-8")
        behavior_path = folder / str(spec.get("model_behavior_file", "model_behavior.md"))
        if not behavior_path.is_file():
            errors.append(f"product-{product_id} no tiene model_behavior.md")
        app_guide_path = folder / str(spec.get("app_user_guide_file", "app/README.md"))
        expected_language = str(spec.get("design_language") or "en").casefold()
        if not app_guide_path.is_file():
            errors.append(f"product-{product_id} no tiene la guia de uso dentro de app")
        elif f"**Language:** `{expected_language}`" not in app_guide_path.read_text(encoding="utf-8"):
            errors.append(f"product-{product_id} tiene una guia app con idioma incorrecto")
        if "## Functional model guide" not in app_markdown:
            errors.append(f"product-{product_id} no incluye la guia funcional en su ficha")
        section = re.search(r"## Tabs \(order\)(.*?)(?=\n## |\Z)", app_markdown, flags=re.S)
        markdown_tabs = re.findall(r"^- `([^`]+)`$", section.group(1), flags=re.M) if section else []
        if markdown_tabs != spec.get("tabs", []):
            errors.append(f"product-{product_id} no coincide con sus pestanas")
        if "app_design_spec.md" not in str(spec.get("source_of_truth", "")):
            errors.append(f"product-{product_id} debe conservar la ficha como contrato editorial")
        if not spec.get("source_lineage") or not spec.get("model_base"):
            errors.append(f"product-{product_id} no declara sources y modelo base")
        identity = spec.get("visual_identity", {})
        if not identity.get("name") or not identity.get("motif"):
            errors.append(f"product-{product_id} no tiene identidad visual declarada")
        palette = spec.get("palette", {})
        required_palette = {"background", "ink", "accent", "secondary", "highlight"}
        if set(palette) < required_palette or len({str(palette.get(key)).casefold() for key in required_palette}) != 5:
            errors.append(f"product-{product_id} no tiene una paleta visual de cinco colores validos")
        if review.get("status") not in {"pending_human_review", "approved"}:
            errors.append(f"product-{product_id} tiene un estado de revision invalido")
        master_references += 1 if (ROOT / str(spec.get("master_workbook", "templates/excel-first-base/master_workbook.xlsx"))).is_file() else 0
        if spec.get("native_reference_mode") != "master_workbook_contract":
            errors.append(f"product-{product_id} no usa el contrato del master_workbook.xlsx")
        if not (ROOT / str(spec.get("master_workbook", "templates/excel-first-base/master_workbook.xlsx"))).is_file():
            errors.append(f"product-{product_id} no declara un master_workbook.xlsx existente")
    if errors:
        print("DESIGN_BATCH_QA=FAILED")
        print("\n".join(errors))
        return 1
    print("DESIGN_BATCH_QA=APPROVED")
    print(f"Products={len(PRODUCT_IDS)}")
    print(f"MasterReferences={master_references}")
    print(f"MarkdownSpecs={markdown_specs}")
    print("HumanReview=pending_for_user_validation")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
