from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

from app_validation_rules import validate_product_rules
from factory_config import active_product_ids


ROOT = Path(__file__).resolve().parent.parent
OUT_ROOT = ROOT / "outputs" / "listings"
PRODUCT_IDS = active_product_ids()
THRESHOLD = 95


def inspect_native_package(folder: Path, spec: dict) -> dict:
    master = ROOT / str(spec.get("master_workbook", "templates/excel-first-base/master_workbook.xlsx"))
    contract = folder / str(spec.get("model_contract_file", "model_contract.json"))
    guide = folder / str(spec.get("model_behavior_file", "model_behavior.md"))
    app_guide = folder / str(spec.get("app_user_guide_file", "app/README.md"))
    workbooks = sorted((folder / "app").glob("*.xlsx"))
    palette = spec.get("palette") or {}
    identity = spec.get("visual_identity") or {}
    return {
        "master_workbook_exists": master.is_file(),
        "model_contract_exists": contract.is_file(),
        "model_behavior_exists": guide.is_file(),
        "app_user_guide_exists": app_guide.is_file(),
        "final_workbook_exists": bool(workbooks),
        "visual_identity_declared": bool(identity.get("name") and identity.get("motif")),
        "palette_declared": len({str(palette.get(key)).casefold() for key in ("background", "ink", "accent", "secondary", "highlight")}) == 5,
        "source_lineage_declared": bool(spec.get("source_lineage")),
        "native_model_declared": bool(spec.get("model_base")),
        "native_contract_mode": spec.get("native_reference_mode") == "master_workbook_contract",
    }


def score_functionality(spec: dict, native: dict) -> tuple[int, list[str]]:
    features = spec.get("functional_scope", [])
    tabs = spec.get("tabs", [])
    score = min(20, len(features) * 3)
    evidence = [f"Documented capabilities: {len(features)}"]
    if len(tabs) >= 7:
        score += 5
        evidence.append("The workbook has a complete multi-screen structure")
    if native["final_workbook_exists"] and native["model_contract_exists"] and native["model_behavior_exists"] and native["app_user_guide_exists"]:
        score += 5
        evidence.append("The final XLSX, model contract and independent guide are present")
    if {"Start here", "Dashboard", "Settings"}.issubset(tabs):
        score += 5
        evidence.append("The flow includes onboarding, operating view and settings")
    return min(35, score), evidence


def score_design(spec: dict, native: dict, folder: Path) -> tuple[int, list[str]]:
    checks = [
        (native["master_workbook_exists"], "The shared master_workbook.xlsx is present"),
        (native["native_model_declared"] and native["source_lineage_declared"], "The design declares its model and source lineage"),
        (native["native_contract_mode"], "The design uses the native XLSX contract mode"),
        (native["visual_identity_declared"] and native["palette_declared"], "A product-specific visual identity and five-colour palette are declared"),
    ]
    score = sum(7 for passed, _ in checks if passed)
    evidence = [message for passed, message in checks if passed]
    if spec.get("visual_acceptance", {}).get("native_geometry_required"):
        score += 2
        evidence.append("Native row, column, table and object geometry is required")
    return min(30, score), evidence


def score_utility(spec: dict) -> tuple[int, list[str]]:
    score = 0
    evidence: list[str] = []
    if spec.get("target_buyer"):
        score += 5
        evidence.append("The intended buyer and use case are explicit")
    if spec.get("title") and spec.get("subtitle"):
        score += 5
        evidence.append("The product promise is understandable")
    if len(spec.get("functional_scope", [])) >= 7:
        score += 5
        evidence.append("The feature set is substantial")
    if spec.get("signature_chart", {}).get("type"):
        score += 5
        evidence.append("The primary visual communicates the product promise")
    if len(spec.get("tabs", [])) >= 7:
        score += 5
        evidence.append("The structure supports a repeatable workflow")
    return min(25, score), evidence


def improvement_plan(spec: dict) -> tuple[int, list[str], list[str]]:
    improvements = [
        "Create the Spanish version while preserving the native grid and table contracts.",
        "Validate formulas, empty states and edge cases once the XLSX exists.",
        "Review model-library feedback after the first user test.",
    ]
    return 10, improvements, ["The improvements are bounded and scheduled after the core workbook gate."]


def evaluate_product(product_id: int) -> dict:
    folder = OUT_ROOT / f"product-{product_id}"
    spec = json.loads((folder / "design_spec.json").read_text(encoding="utf-8"))
    native = inspect_native_package(folder, spec)
    functionality_score, functionality_evidence = score_functionality(spec, native)
    design_score, design_evidence = score_design(spec, native, folder)
    utility_score, utility_evidence = score_utility(spec)
    improvement_score, improvements, improvement_evidence = improvement_plan(spec)
    automatic_rules = validate_product_rules(folder, spec)
    total = functionality_score + design_score + utility_score + improvement_score
    passed = total >= THRESHOLD and automatic_rules["all_passed"]
    return {
        "product_id": product_id,
        "slug": spec.get("slug"),
        "title": spec.get("title"),
        "evaluated_at": datetime.now(timezone.utc).isoformat(),
        "threshold": THRESHOLD,
        "score": total,
        "status": "ready_for_app_generation" if passed else "needs_design_improvement",
        "next_stage": "generate_app" if passed else "return_to_design",
        "categories": {
            "functionality": {"score": functionality_score, "max": 35, "evidence": functionality_evidence},
            "design": {"score": design_score, "max": 30, "evidence": design_evidence},
            "utility": {"score": utility_score, "max": 25, "evidence": utility_evidence},
            "improvements": {"score": improvement_score, "max": 10, "evidence": improvement_evidence},
        },
        "native_checks": native,
        "improvement_plan": improvements,
        "human_review_required": spec.get("visual_acceptance", {}).get("human_review_required", True),
        "automatic_rules": automatic_rules,
    }


def main() -> int:
    reports: list[dict] = []
    for product_id in PRODUCT_IDS:
        report = evaluate_product(product_id)
        reports.append(report)
        (OUT_ROOT / f"product-{product_id}" / "auto_evaluation.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"product-{product_id}: score={report['score']}/100 status={report['status']}")
    queue = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "threshold": THRESHOLD,
        "ready_products": [report["product_id"] for report in reports if report["score"] >= THRESHOLD and report["automatic_rules"]["all_passed"]],
        "blocked_products": [report["product_id"] for report in reports if report["score"] < THRESHOLD or not report["automatic_rules"]["all_passed"]],
        "reports": [f"outputs/listings/product-{report['product_id']}/auto_evaluation.json" for report in reports],
    }
    aggregate = {"generated_at": datetime.now(timezone.utc).isoformat(), "threshold": THRESHOLD, "summary": {"products": len(reports), "ready_for_app_generation": len(queue["ready_products"]), "needs_design_improvement": len(queue["blocked_products"])}, "products": reports, "queue_file": "outputs/listings/app_generation_queue.json"}
    aggregate_path = OUT_ROOT / "auto_evaluation.json"
    aggregate_path.write_text(json.dumps(aggregate, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (OUT_ROOT / "app_generation_queue.json").write_text(json.dumps(queue, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"AUTO_EVALUATION={'PASSED' if not queue['blocked_products'] else 'PARTIAL'}")
    print(f"ReadyForApp={queue['ready_products']}")
    print(f"NeedsImprovement={queue['blocked_products']}")
    if "--no-open" not in sys.argv and os.name == "nt":
        try:
            os.startfile(str(aggregate_path))
        except OSError:
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
