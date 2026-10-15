from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from time import perf_counter

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
import scripts  # noqa: F401 - initializes categorized legacy module paths

from factory_config import PIPELINE_PATH, active_product_ids, all_product_ids, finalized_product_ids


def run(script: str, *args: str, env: dict[str, str] | None = None) -> int:
    return int(subprocess.run([sys.executable, "-m", script, *args], cwd=ROOT, env=env, check=False).returncode)


def run_stage(label: str, script: str, *args: str, records: list[dict] | None = None, env: dict[str, str] | None = None) -> bool:
    print(f"FACTORY_STAGE={label}")
    started = perf_counter()
    code = run(script, *args, env=env)
    duration = round(perf_counter() - started, 3)
    if records is not None:
        records.append({"label": label, "script": script, "args": list(args), "status": "passed" if code == 0 else "failed", "return_code": code, "duration_seconds": duration})
    if code:
        print(f"FACTORY_STAGE={label}_FAILED code={code} duration={duration}s")
        return False
    print(f"FACTORY_STAGE={label}_PASSED duration={duration}s")
    return True


def auto_publish_candidates(product_ids: tuple[int, ...]) -> tuple[int, ...]:
    config = json.loads(PIPELINE_PATH.read_text(encoding="utf-8"))
    policy = config.get("auto_publish_policy", {})
    if not policy.get("enabled") or not policy.get("requires_passed_factory_gates", True):
        return ()
    minimum_new_id = int(policy.get("minimum_new_product_id", 2**31))
    one_time = {int(value) for value in policy.get("one_time_product_ids", [])}
    return tuple(product_id for product_id in product_ids if product_id in one_time or product_id >= minimum_new_id)


def publish_product(product_id: int, records: list[dict]) -> bool:
    manifest_path = ROOT / "outputs" / "listings" / f"product-{product_id}" / "draft_package" / "listing.json"
    if not manifest_path.is_file():
        print(f"FACTORY_AUTOPUBLISH=SKIPPED product={product_id} reason=missing_manifest")
        return False
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("etsy_status") == "active":
        print(f"FACTORY_AUTOPUBLISH=SKIPPED product={product_id} reason=already_active")
        return True

    publish_env = os.environ.copy()
    publish_env["ETSY_DRY_RUN"] = "false"
    publish_env["ETSY_PUBLISH_ENABLED"] = "true"
    manifest_arg = str(manifest_path)
    if not manifest.get("etsy_listing_id"):
        if not run_stage(f"PRODUCT_{product_id}_ETSY_DRAFT_UPLOAD", "scripts.etsy.etsy_publish", "--manifest", manifest_arg, "--execute", records=records, env=publish_env):
            return False
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("etsy_status") != "active":
        if not run_stage(f"PRODUCT_{product_id}_ETSY_ACTIVATE", "scripts.etsy.etsy_activate", "--manifest", manifest_arg, "--execute", records=records, env=publish_env):
            return False
    return True


def consume_one_time_publish_permission(product_id: int) -> None:
    config = json.loads(PIPELINE_PATH.read_text(encoding="utf-8"))
    policy = config.get("auto_publish_policy", {})
    one_time = [int(value) for value in policy.get("one_time_product_ids", [])]
    if product_id not in one_time:
        return
    policy["one_time_product_ids"] = [value for value in one_time if value != product_id]
    config["auto_publish_policy"] = policy
    PIPELINE_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def open_report() -> None:
    report = ROOT / "outputs" / "listings" / "auto_evaluation.json"
    if os.name == "nt" and report.exists():
        try:
            os.startfile(str(report))
        except OSError:
            pass


def main() -> int:
    parser = argparse.ArgumentParser(description="Ejecuta la factoría con validación directa del XLSX.")
    parser.add_argument("--product", type=int, choices=all_product_ids(), help="Limitar la ejecución a un producto")
    parser.add_argument("--no-open", action="store_true", help="No abrir auto_evaluation.json al terminar")
    args = parser.parse_args()

    failures: list[str] = []
    stage_records: list[dict] = []
    published_products: list[int] = []
    run_started_at = datetime.now(timezone.utc).isoformat()
    active_products = active_product_ids()
    is_finalized_target = args.product in finalized_product_ids() if args.product else False

    if not run_stage("AUTO_EVALUATION", "scripts.validation.evaluate_excel_products", "--no-open", records=stage_records):
        failures.append("auto_evaluation")
    architecture_args = ("--product", str(args.product)) if args.product else ()
    architecture_ok = True
    if not is_finalized_target and not run_stage("DESIGN_ARCHITECTURE", "scripts.validation.validate_design_architecture", *architecture_args, records=stage_records):
        failures.append("design_architecture")
        architecture_ok = False
    generation_args = ("--product", str(args.product)) if args.product else ()
    if architecture_ok and not run_stage("APP_GENERATION", "scripts.generation.generate_autonomous_apps", *generation_args, records=stage_records):
        failures.append("app_generation")
    elif not architecture_ok:
        print("FACTORY_STAGE=APP_GENERATION_SKIPPED reason=design_architecture")

    target_products = (args.product,) if args.product else active_products
    if not failures:
        for product_id in target_products:
            semantic_label = f"PRODUCT_{product_id}_SEMANTIC_CONTRACT"
            if not run_stage(semantic_label, "scripts.validation.audit_design_contract", "--product", str(product_id), records=stage_records):
                failures.append(semantic_label)
                continue
            native_label = f"PRODUCT_{product_id}_NATIVE_XLSX_PREFLIGHT"
            if not run_stage(native_label, "scripts.validation.preflight_visual_contract", "--product", str(product_id), records=stage_records):
                failures.append(native_label)
    else:
        stage_records.append({"label": "NATIVE_XLSX_QA", "script": "scripts.validation.preflight_visual_contract", "args": [], "status": "skipped", "return_code": None, "duration_seconds": 0.0, "reason": "upstream_quality_gate"})
        print("FACTORY_STAGE=NATIVE_XLSX_QA_SKIPPED reason=upstream_quality_gate")

    if failures:
        stage_records.append({"label": "DRAFT_PACKAGE", "script": "scripts.generation.prepare_autonomous_drafts", "args": [], "status": "skipped", "return_code": None, "duration_seconds": 0.0, "reason": "upstream_quality_gate"})
        print("FACTORY_STAGE=DRAFT_PACKAGE_SKIPPED reason=upstream_quality_gate")
    elif not run_stage("DRAFT_PACKAGE", "scripts.generation.prepare_autonomous_drafts", *generation_args, records=stage_records):
        failures.append("draft_package")
    elif not run_stage("MARKETING_ASSETS", "scripts.marketing.marketing_assets", *generation_args, records=stage_records):
        failures.append("marketing_assets")
    else:
        for product_id in target_products:
            manifest_path = ROOT / "outputs" / "listings" / f"product-{product_id}" / "draft_package" / "listing.json"
            if not run_stage(
                f"PRODUCT_{product_id}_MARKETING_QA",
                "scripts.validation.validate_marketing_assets",
                "--manifest",
                str(manifest_path),
                records=stage_records,
            ):
                failures.append(f"product_{product_id}_marketing_qa")

    if not failures:
        publish_targets = auto_publish_candidates(target_products)
        for product_id in publish_targets:
            if not publish_product(product_id, stage_records):
                failures.append(f"product_{product_id}_etsy_autopublish")
                continue
            published_products.append(product_id)
            consume_one_time_publish_permission(product_id)

    final_status = "blocked" if failures else ("published" if published_products else "ready_for_draft_upload")
    (ROOT / "outputs" / "listings" / "factory_run_report.json").write_text(
        json.dumps({"schema_version": "2.1.0", "started_at": run_started_at, "completed_at": datetime.now(timezone.utc).isoformat(), "status": final_status, "active_products": list(active_products), "target_product": args.product, "published_products": published_products, "failures": failures, "stages": stage_records, "total_duration_seconds": round(sum(item["duration_seconds"] for item in stage_records), 3)}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    if not args.no_open:
        open_report()
    if failures:
        print(f"AUTONOMOUS_FACTORY=BLOCKED stages={','.join(failures)}")
        return 1
    print("AUTONOMOUS_FACTORY=PUBLISHED" if publish_targets else "AUTONOMOUS_FACTORY=READY_FOR_DRAFT_UPLOAD")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
