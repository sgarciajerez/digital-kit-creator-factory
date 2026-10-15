from __future__ import annotations

import json
import argparse
import math
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path
from statistics import median

from listing_package import ListingDraft, validate_listing_draft
from etsy_client import load_env
from etsy_draft import build_etsy_draft_request
from factory_config import active_product_ids
from artifact_integrity import approval_is_current, sha256
from model_contracts import contract_for_product


ROOT = Path(__file__).resolve().parent.parent
OUT_ROOT = ROOT / "outputs" / "listings"
PRICING_PATH = ROOT / "config" / "competition_pricing.json"
PRODUCT_IDS = active_product_ids()
MAX_PRICE_EUR = 3.0


def launch_price(observed: list[float], discount: float = 0.08) -> tuple[float, float, float, list[float]]:
    if not observed or any(not math.isfinite(float(value)) or float(value) <= 0 for value in observed):
        raise ValueError("Los comparables deben contener precios positivos y finitos")
    capped_observed = [float(value) for value in observed]
    reference = float(median(capped_observed))
    raw = reference * (1 - discount)
    # Keep the price just below the target and avoid rounding up over it.
    price = min(MAX_PRICE_EUR, math.floor(raw * 100) / 100)
    return reference, raw, price, capped_observed


def tags_for(spec: dict) -> list[str]:
    candidates = {
        "wedding-planner": ["wedding planner", "wedding budget", "excel template", "guest list", "wedding checklist", "vendor tracker", "digital download"],
        "debt-payoff-planner": ["debt payoff", "debt snowball", "debt avalanche", "excel planner", "payment tracker", "debt spreadsheet", "digital download"],
        "small-business-cashflow": ["cash flow tracker", "small business", "income expense", "excel template", "invoice tracker", "business budget", "digital download"],
        "meal-planner-grocery-budget": ["meal planner", "grocery list", "food budget", "excel template", "meal prep", "recipe planner", "digital download"],
        "home-renovation-planner": ["home renovation", "renovation planner", "remodel budget", "excel template", "project tracker", "contractor quotes", "digital download"],
    }.get(spec["slug"], ["excel template", "digital download", "planner"])
    return candidates[:13]


def description_for(spec: dict, price: float) -> str:
    features = "\n".join(f"• {item}" for item in spec["functional_scope"])
    return (
        f"{spec['title']} — a connected Excel workbook for {spec['target_buyer'].lower()}\n\n"
        f"{spec['subtitle']}\n\n"
        f"WHAT'S INCLUDED\n{features}\n\n"
        "HOW TO USE\n1. Open the Start here sheet.\n2. Enter your inputs in the highlighted cells.\n"
        "3. Review the dashboard and connected views.\n\n"
        "Digital download in XLSX format. This is an organisational planning tool, not financial, tax, medical or legal advice."
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Prepara paquetes locales de productos validados")
    parser.add_argument("--product", type=int, help="Producto existente; sin argumento usa el catálogo activo")
    args = parser.parse_args()
    pricing = json.loads(PRICING_PATH.read_text(encoding="utf-8"))
    env_values = load_env()
    shop_id = int(env_values["ETSY_SHOP_ID"]) if env_values.get("ETSY_SHOP_ID", "").isdigit() else None
    taxonomy_id = int(env_values["ETSY_TAXONOMY_ID"]) if env_values.get("ETSY_TAXONOMY_ID", "").isdigit() else None
    results = []
    for product_id in ((args.product,) if args.product else PRODUCT_IDS):
        folder = OUT_ROOT / f"product-{product_id}"
        spec = json.loads((folder / "design_spec.json").read_text(encoding="utf-8"))
        app_report = json.loads((folder / "app_generation_report.json").read_text(encoding="utf-8"))
        evaluation = json.loads((folder / "auto_evaluation.json").read_text(encoding="utf-8"))
        customer = app_report.get("customer_copy", {})
        from preflight_visual_contract import evaluate
        native_passed, _ = evaluate(product_id)
        clean_file = ROOT / customer.get("file", "missing-customer-copy")
        clean_current = clean_file.is_file() and customer.get("validation", {}).get("passed") and customer.get("sha256") == sha256(clean_file)
        if (
            evaluation.get("status") != "ready_for_app_generation"
            or not native_passed
            or not clean_current
            or not approval_is_current(app_report, ROOT / app_report["file"], ROOT / spec.get("master_workbook", "templates/excel-first-base/master_workbook.xlsx"), contract_for_product(product_id), spec)
        ):
            stale_manifest = folder / "draft_package" / "listing.json"
            if stale_manifest.exists():
                stale = json.loads(stale_manifest.read_text(encoding="utf-8"))
                stale["status"] = "blocked_by_native_xlsx_qa"
                stale["next_action"] = "Reconcile the XLSX with model_contract.md and the native structural audit."
                stale_manifest.write_text(json.dumps(stale, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            results.append({"product_id": product_id, "status": "blocked_by_quality_gate"})
            print(f"product-{product_id}: draft blocked")
            continue
        pricing_input = (pricing.get("products") or {}).get(spec["slug"])
        if not pricing_input:
            results.append({
                "product_id": product_id,
                "status": "blocked_by_missing_pricing",
                "reason": f"No existe una ficha de comparables para {spec['slug']} en competition_pricing.json",
            })
            print(f"product-{product_id}: draft blocked; falta pricing para {spec['slug']}")
            continue
        reference, raw, price, capped_observed = launch_price(pricing_input["observed_prices"])
        draft_dir = folder / "draft_package"
        draft_dir.mkdir(parents=True, exist_ok=True)
        source_file = ROOT / app_report["file"]
        local_file = draft_dir / source_file.name
        shutil.copy2(source_file, local_file)
        shutil.copy2(clean_file, draft_dir / clean_file.name)
        shutil.copy2(folder / "app" / "README.md", draft_dir / "README.md")
        behavior_file = folder / str(spec.get("model_behavior_file", "model_behavior.md"))
        model_qa_file = folder / "model_behavior_qa.json"
        if not behavior_file.is_file():
            results.append({"product_id": product_id, "status": "blocked_by_missing_model_behavior"})
            print(f"product-{product_id}: draft blocked; falta {behavior_file.name}")
            continue
        if not model_qa_file.is_file():
            results.append({"product_id": product_id, "status": "blocked_by_missing_model_behavior_qa"})
            print(f"product-{product_id}: draft blocked; falta {model_qa_file.name}")
            continue
        shutil.copy2(behavior_file, draft_dir / behavior_file.name)
        shutil.copy2(model_qa_file, draft_dir / model_qa_file.name)
        draft = ListingDraft(
            title=spec["title"],
            description=description_for(spec, price),
            tags=tags_for(spec),
            price=price,
            currency="EUR",
            policy_reviewed=True,
            rights_confirmed=True,
        )
        validation = validate_listing_draft(draft)
        manifest = {
            "product_id": product_id,
            "platform": "etsy-local-draft",
            "status": "draft_ready" if validation.passed else "needs_review",
            "activation_requested": False,
            "draft": draft.model_dump(),
            "validation": validation.model_dump(),
            "local_file": local_file.name,
            "blank_file": clean_file.name,
            "usage_guide": "README.md",
            "artifact_identity": app_report["artifact_identity"],
            "blank_sha256": customer["sha256"],
            "model_behavior_file": behavior_file.name,
            "model_behavior_qa_file": model_qa_file.name,
            "image_files": [],
            "video_file": None,
            "quality_basis": "native_xlsx_contract_and_functional_model_qa",
            "pricing": {
                "observed_currency": pricing["currency_observed"],
                "shop_currency": draft.currency,
                "maximum_price": MAX_PRICE_EUR,
                "capped_observed_prices": capped_observed,
                "median_comparable_price": round(reference, 2),
                "discount_rate": 0.08,
                "raw_target_price": round(raw, 2),
                "recommended_launch_price": price,
                "sources": pricing_input["sources"],
                "note": pricing_input["note"],
            },
            "created_at": datetime.now(timezone.utc).isoformat(),
            "next_action": "Review the XLSX and listing copy manually before any external upload; activation remains disabled.",
        }
        if shop_id and taxonomy_id:
            etsy_request = build_etsy_draft_request(draft, shop_id, taxonomy_id, local_file.name)
            (draft_dir / "etsy_draft_request.json").write_text(json.dumps(etsy_request.model_dump(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            manifest["etsy_draft_request"] = "etsy_draft_request.json"
        else:
            manifest["etsy_draft_request"] = None
        (draft_dir / "listing.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        (draft_dir / "pricing.json").write_text(json.dumps(manifest["pricing"], ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        results.append({"product_id": product_id, "status": manifest["status"], "price": price, "draft": str((draft_dir / "listing.json").relative_to(ROOT))})
        print(f"product-{product_id}: draft={manifest['status']} price={price:.2f} EUR")
    batch = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "currency": "EUR",
        "products": results,
        "activation_requested": False,
        "etsy_credentials_configured": bool(shop_id and taxonomy_id),
        "etsy_publish_enabled": env_values.get("ETSY_PUBLISH_ENABLED", "").casefold() in {"1", "true", "yes", "on"},
        "etsy_dry_run": env_values.get("ETSY_DRY_RUN", "").casefold() in {"1", "true", "yes", "on"},
    }
    (OUT_ROOT / "draft_generation_report.json").write_text(json.dumps(batch, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    aggregate_path = OUT_ROOT / "auto_evaluation.json"
    if aggregate_path.exists():
        aggregate = json.loads(aggregate_path.read_text(encoding="utf-8"))
        batch_generation_path = OUT_ROOT / "app_generation_report.json"
        if batch_generation_path.is_file():
            aggregate["app_generation"] = json.loads(batch_generation_path.read_text(encoding="utf-8"))
        else:
            # A targeted factory run writes the product report only. Keep the
            # aggregate package valid without requiring a batch side effect.
            product_reports = []
            for item in results:
                product_report_path = OUT_ROOT / f"product-{item['product_id']}" / "app_generation_report.json"
                if product_report_path.is_file():
                    product_reports.append(json.loads(product_report_path.read_text(encoding="utf-8")))
            aggregate["app_generation"] = {"products": product_reports, "all_passed": all(report.get("validation", {}).get("passed", False) for report in product_reports)}
        aggregate["draft_generation"] = batch
        aggregate_path.write_text(json.dumps(aggregate, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return 0 if all(item["status"] == "draft_ready" for item in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
