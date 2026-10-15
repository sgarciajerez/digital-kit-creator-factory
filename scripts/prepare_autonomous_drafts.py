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
from digital_kit import product_kit_policy
from etsy_price_research import search_etsy_prices


ROOT = Path(__file__).resolve().parent.parent
OUT_ROOT = ROOT / "outputs" / "listings"
PRICING_PATH = ROOT / "config" / "competition_pricing.json"
PRODUCT_IDS = active_product_ids()
def launch_price(
    observed: list[float],
    discount: float = 0.0,
    observed_to_shop_currency: float = 1.0,
) -> tuple[float, float, float, list[float]]:
    if not math.isfinite(float(discount)) or not 0 <= float(discount) < 1:
        raise ValueError("El descuento debe estar entre 0 y menos de 100 %")
    if not observed or any(not math.isfinite(float(value)) or float(value) <= 0 for value in observed):
        raise ValueError("Los comparables deben contener precios positivos y finitos")
    if not math.isfinite(float(observed_to_shop_currency)) or observed_to_shop_currency <= 0:
        raise ValueError("El tipo de conversión debe ser positivo y finito")
    observed_prices = [float(value) for value in observed]
    converted = [value * observed_to_shop_currency for value in observed_prices]
    reference = round(float(median(converted)), 2)
    raw = round(reference * (1 - discount), 2)
    price = math.floor(raw * 100) / 100
    return reference, raw, price, observed_prices


def pricing_queries_for(pricing_input: dict, seo_data: dict | None, slug: str) -> list[str]:
    """Prefer explicit Etsy queries, then product research, then its product name."""
    explicit = pricing_input.get("search_queries") or []
    if explicit:
        return list(dict.fromkeys(str(value).strip() for value in explicit if str(value).strip()))[:3]
    if seo_data:
        hypotheses = seo_data.get("keyword_hypotheses") or []
        phrases = [str(item.get("phrase", "")).strip() for item in hypotheses if isinstance(item, dict)]
        phrases = [phrase for phrase in phrases if phrase]
        if phrases:
            return list(dict.fromkeys(phrases))[:3]
        signals = seo_data.get("market_signals") or []
        queries = [str(item.get("query", "")).strip() for item in signals if isinstance(item, dict)]
        queries = [query for query in queries if query]
        if queries:
            return list(dict.fromkeys(queries))[:3]
    return [f"{slug.replace('-', ' ')} spreadsheet"]


def tags_for(spec: dict) -> list[str]:
    candidates = {
        "wedding-planner": ["wedding planner", "wedding budget", "excel template", "guest list", "wedding checklist", "vendor tracker", "digital download"],
        "debt-payoff-planner": ["debt payoff", "debt snowball", "debt avalanche", "excel planner", "payment tracker", "debt spreadsheet", "digital download"],
        "small-business-cashflow": ["cash flow tracker", "small business", "income expense", "excel template", "invoice tracker", "business budget", "digital download"],
        "meal-planner-grocery-budget": ["meal planner", "grocery list", "food budget", "excel template", "meal prep", "recipe planner", "digital download"],
        "home-renovation-planner": ["home renovation", "renovation planner", "remodel budget", "excel template", "project tracker", "contractor quotes", "digital download"],
        "moving-home-planner": ["moving planner", "moving checklist", "home move planner", "moving budget", "relocation plan", "excel template", "move timeline"],
    }.get(spec["slug"], ["excel template", "digital download", "planner"])
    return candidates[:13]


def description_for(spec: dict, price: float, kit_report: dict | None = None) -> str:
    features = "\n".join(f"• {item}" for item in spec["functional_scope"])
    kit_report = kit_report or {}
    included = ["Worked-example Excel workbook", "Clean blank Excel workbook", "PDF user guide"]
    if kit_report.get("ebook_included"):
        included.append("Topic-specific e-book PDF")
    html_tool = kit_report.get("interactive_html_tool") or {}
    if html_tool.get("passed"):
        included.append(f"Offline interactive HTML tool: {html_tool.get('title', 'Topic calculator')} (opens in a browser; no internet required)")
    included.extend(module["title"] for module in kit_report.get("selected_modules", []))
    included_text = "\n".join(f"- {item}" for item in included)
    return (
        f"{spec['title']} — a connected Excel workbook for {spec['target_buyer'].lower()}\n\n"
        f"{spec['subtitle']}\n\n"
        f"WHAT'S INCLUDED\n{features}\n\nDIGITAL KIT FILES\n{included_text}\n\n"
        "HOW TO USE\n1. Open the Start here sheet.\n2. Enter your inputs in the highlighted cells.\n"
        "3. Review the dashboard and connected views.\n\n"
        "Instant digital ZIP download containing the Excel workbooks, PDF resources and any interactive HTML tool listed above. No physical item is shipped. "
        "This is an organisational planning tool, not financial, tax, medical or legal advice."
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Prepara paquetes locales de productos validados")
    parser.add_argument("--product", type=int, help="Producto existente; sin argumento usa el catálogo activo")
    args = parser.parse_args()
    pricing = json.loads(PRICING_PATH.read_text(encoding="utf-8"))
    shop_currency = str(pricing.get("shop_currency", "EUR")).upper()
    env_values = load_env()
    shop_id = int(env_values["ETSY_SHOP_ID"]) if env_values.get("ETSY_SHOP_ID", "").isdigit() else None
    taxonomy_id = int(env_values["ETSY_TAXONOMY_ID"]) if env_values.get("ETSY_TAXONOMY_ID", "").isdigit() else None
    results = []
    for product_id in ((args.product,) if args.product else PRODUCT_IDS):
        folder = OUT_ROOT / f"product-{product_id}"
        spec = json.loads((folder / "design_spec.json").read_text(encoding="utf-8"))
        kit_policy = product_kit_policy(spec.get("slug", ""))
        if not kit_policy or kit_policy.get("product_readiness", "ready") != "ready":
            stale_manifest = folder / "draft_package" / "listing.json"
            if stale_manifest.is_file():
                stale = json.loads(stale_manifest.read_text(encoding="utf-8"))
                stale["status"] = kit_policy.get("product_readiness", "blocked_by_missing_kit_profile") if kit_policy else "blocked_by_missing_kit_profile"
                stale["next_action"] = kit_policy.get("readiness_note", "Declare and approve a digital-kit profile before packaging.")
                stale_manifest.write_text(json.dumps(stale, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            status = kit_policy.get("product_readiness", "blocked_by_missing_kit_profile") if kit_policy else "blocked_by_missing_kit_profile"
            results.append({"product_id": product_id, "status": status})
            print(f"product-{product_id}: draft blocked by kit readiness")
            continue
        app_report = json.loads((folder / "app_generation_report.json").read_text(encoding="utf-8"))
        evaluation = json.loads((folder / "auto_evaluation.json").read_text(encoding="utf-8"))
        customer = app_report.get("customer_copy", {})
        from preflight_visual_contract import evaluate
        native_passed, _ = evaluate(product_id)
        clean_file = ROOT / customer.get("file", "missing-customer-copy")
        clean_current = clean_file.is_file() and customer.get("validation", {}).get("passed") and customer.get("sha256") == sha256(clean_file)
        ebook_record = app_report.get("ebook") or {}
        ebook_file = ROOT / ebook_record.get("file", "missing-ebook.pdf")
        ebook_cover = Path(ebook_record.get("cover_image", "missing-cover.png"))
        ebook_source = Path(ebook_record.get("content_file", "missing-content.json"))
        ebook_current = not ebook_record.get("included") or (
            ebook_record.get("passed")
            and ebook_file.is_file()
            and ebook_record.get("sha256") == sha256(ebook_file)
            and ebook_cover.is_file()
            and ebook_record.get("cover_sha256") == sha256(ebook_cover)
            and ebook_source.is_file()
            and ebook_record.get("content_sha256") == sha256(ebook_source)
        )
        guide_record = app_report.get("user_guide_pdf") or {}
        guide_pdf = ROOT / guide_record.get("file", "missing-user-guide.pdf")
        guide_markdown = folder / "app" / "README.md"
        user_guide_current = (
            guide_record.get("passed")
            and guide_pdf.is_file()
            and guide_record.get("sha256") == sha256(guide_pdf)
            and guide_markdown.is_file()
            and guide_record.get("markdown_sha256") == sha256(guide_markdown)
        )
        kit_record = app_report.get("digital_kit") or {}
        kit_archive = ROOT / kit_record.get("file", "missing-digital-kit.zip")
        kit_manifest = ROOT / kit_record.get("manifest_file", "missing-digital-kit-manifest.json")
        kit_manifest_data = {}
        if kit_manifest.is_file():
            try:
                kit_manifest_data = json.loads(kit_manifest.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                kit_manifest_data = {}
        catalog_path = ROOT / kit_manifest_data.get("catalog_file", "missing-digital-kit-catalog.json")
        module_sources = kit_manifest_data.get("module_sources", [])
        html_tool_source = (kit_manifest_data.get("interactive_html_tool") or {}).get("generator_file")
        html_tool_source_path = ROOT / html_tool_source if html_tool_source else None
        html_tool_source_current = (
            html_tool_source_path is None
            or (html_tool_source_path.is_file() and (kit_manifest_data.get("interactive_html_tool") or {}).get("generator_sha256") == sha256(html_tool_source_path))
        )
        kit_sources_current = (
            catalog_path.is_file()
            and kit_manifest_data.get("catalog_sha256") == sha256(catalog_path)
            and html_tool_source_current
            and all(
                (ROOT / item.get("file", "missing-module-source.json")).is_file()
                and item.get("sha256") == sha256(ROOT / item.get("file", "missing-module-source.json"))
                for item in module_sources
            )
        )
        kit_current = (
            kit_record.get("passed")
            and kit_archive.is_file()
            and kit_record.get("sha256") == sha256(kit_archive)
            and kit_manifest.is_file()
            and kit_record.get("manifest_sha256") == sha256(kit_manifest)
            and kit_sources_current
        )
        if (
            evaluation.get("status") != "ready_for_app_generation"
            or not native_passed
            or not clean_current
            or not ebook_current
            or not user_guide_current
            or not kit_current
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
        seo_data = None
        seo_path = folder / "seo_keyword_research.json"
        if product_id >= 12:
            try:
                seo_data = json.loads(seo_path.read_text(encoding="utf-8"))
                seo_title = str(seo_data["listing"]["title"])
                seo_tags = list(seo_data["listing"]["tags"])
                if not seo_title or len(seo_title) > 140 or len(seo_tags) != 13 or any(len(tag) > 20 for tag in seo_tags):
                    raise ValueError("title/tags do not meet current Etsy limits")
            except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
                results.append({"product_id": product_id, "status": "blocked_by_missing_or_invalid_seo_research", "reason": str(exc)})
                print(f"product-{product_id}: draft blocked; revisa seo_keyword_research.json")
                continue
        search_queries = pricing_queries_for(pricing_input, seo_data, spec["slug"])
        try:
            price_research = search_etsy_prices(
                search_queries,
                env_values,
                shop_id=shop_id,
                target_currency=shop_currency,
                limit_per_query=int(pricing.get("sample_size_per_query", 40)),
                minimum_comparables=int(pricing.get("minimum_comparables", 5)),
            )
        except (OSError, RuntimeError, ValueError, KeyError) as exc:
            results.append({
                "product_id": product_id,
                "status": "blocked_by_missing_live_etsy_price_research",
                "reason": str(exc),
                "search_queries": search_queries,
            })
            print(f"product-{product_id}: draft blocked; investigación de precios Etsy no disponible: {exc}")
            continue
        price = float(price_research["recommended_launch_price"])
        reference = float(price_research["market_band"]["median"])
        raw = price
        observed_prices = [float(item["price_in_shop_currency"]) for item in price_research["observations"]]
        draft_dir = folder / "draft_package"
        draft_dir.mkdir(parents=True, exist_ok=True)
        visual_identity = spec.get("visual_identity") or {}
        artwork_ref = str(visual_identity.get("marketing_artwork_file") or "").strip()
        packaged_artwork = None
        if artwork_ref:
            artwork_relative = Path(artwork_ref)
            source_ref = str(visual_identity.get("marketing_artwork_source") or artwork_ref)
            source_relative = Path(source_ref)
            source_artwork = ROOT / source_relative
            if (
                artwork_relative.is_absolute() or ".." in artwork_relative.parts
                or source_relative.is_absolute() or ".." in source_relative.parts
                or not source_artwork.is_file()
            ):
                results.append({"product_id": product_id, "status": "blocked_by_missing_marketing_artwork", "reason": artwork_ref})
                print(f"product-{product_id}: draft blocked; falta el arte de marketing {artwork_ref}")
                continue
            packaged_artwork = draft_dir / artwork_relative
            packaged_artwork.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_artwork, packaged_artwork)
        source_file = ROOT / app_report["file"]
        local_workbook = draft_dir / source_file.name
        shutil.copy2(source_file, local_workbook)
        shutil.copy2(clean_file, draft_dir / clean_file.name)
        shutil.copy2(folder / "app" / "README.md", draft_dir / "README.md")
        local_user_guide = draft_dir / guide_pdf.name
        shutil.copy2(guide_pdf, local_user_guide)
        local_kit = draft_dir / kit_archive.name
        shutil.copy2(kit_archive, local_kit)
        html_tool_manifest = None
        html_tool_record = kit_record.get("interactive_html_tool") or {}
        if html_tool_record.get("passed"):
            html_tool_name = Path(str(html_tool_record.get("file") or "")).name
            html_tool_source = folder / "app" / html_tool_name
            if not html_tool_name or not html_tool_source.is_file():
                results.append({"product_id": product_id, "status": "blocked_by_missing_html_tool_source"})
                print(f"product-{product_id}: draft blocked; falta la herramienta HTML incluida en el ZIP")
                continue
            html_tool_package_path = draft_dir / "marketing_assets" / "source_calculator.html"
            html_tool_package_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(html_tool_source, html_tool_package_path)
            html_tool_manifest = {
                **html_tool_record,
                "filename": html_tool_name,
                "file": html_tool_package_path.relative_to(draft_dir).as_posix(),
            }
        local_ebook = None
        if ebook_record.get("included"):
            local_ebook = draft_dir / ebook_file.name
            shutil.copy2(ebook_file, local_ebook)
        local_cover_image = None
        if ebook_record.get("included") and ebook_cover.is_file():
            local_cover_image = draft_dir / ebook_cover.name
            shutil.copy2(ebook_cover, local_cover_image)
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
            title=seo_data["listing"]["title"] if seo_data else spec["title"],
            description=description_for(spec, price, kit_record),
            tags=seo_data["listing"]["tags"] if seo_data else tags_for(spec),
            price=price,
            currency=shop_currency,
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
            "local_file": local_kit.name,
            "digital_kit_zip": local_kit.name,
            "interactive_html_tool": html_tool_manifest,
            "digital_kit_sha256": kit_record["sha256"],
            "workbook_file": local_workbook.name,
            "blank_file": clean_file.name,
            "usage_guide": "README.md",
            "usage_guide_pdf": local_user_guide.name,
            "usage_guide_pdf_sha256": guide_record["sha256"],
            "ebook_file": local_ebook.name if local_ebook else None,
            "ebook_sha256": ebook_record.get("sha256") if local_ebook else None,
            "ebook_cover_image": local_cover_image.name if local_cover_image else None,
            "marketing_artwork_file": packaged_artwork.relative_to(draft_dir).as_posix() if packaged_artwork else None,
            "artifact_identity": app_report["artifact_identity"],
            "blank_sha256": customer["sha256"],
            "model_behavior_file": behavior_file.name,
            "model_behavior_qa_file": model_qa_file.name,
            "image_files": [local_cover_image.name] if local_cover_image else [],
            "cover_image": local_cover_image.name if local_cover_image else None,
            "video_file": None,
            "quality_basis": "native_xlsx_contract_and_functional_model_qa",
            "pricing": {
                **price_research,
                "observed_currency": "mixed; each observation retains its source currency",
                "shop_currency": draft.currency,
                "observed_prices": observed_prices,
                "median_comparable_price": round(reference, 2),
                "raw_target_price": round(raw, 2),
                "recommended_launch_price": price,
                "sources": list(dict.fromkeys(item["url"] for item in price_research["observations"])),
                "note": "Precio recomendado desde la mediana de listings Etsy activos comparables. El rango central se expresa como Q1–Q3; cada observación y búsqueda queda registrada para revisión.",
                "research_context": pricing_input.get("note", ""),
            },
            "created_at": datetime.now(timezone.utc).isoformat(),
            "next_action": "Review the XLSX and listing copy manually before any external upload; activation remains disabled.",
        }
        if shop_id and taxonomy_id:
            etsy_request = build_etsy_draft_request(draft, shop_id, taxonomy_id, local_kit.name)
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
