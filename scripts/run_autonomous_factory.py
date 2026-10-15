from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from time import perf_counter

from factory_config import active_product_ids, all_product_ids, finalized_product_ids


ROOT = Path(__file__).resolve().parent.parent


def run(script: str, *args: str) -> int:
    return int(subprocess.run([sys.executable, str(ROOT / "scripts" / script), *args], cwd=ROOT, check=False).returncode)


def run_stage(label: str, script: str, *args: str, records: list[dict] | None = None) -> bool:
    print(f"FACTORY_STAGE={label}")
    started = perf_counter()
    code = run(script, *args)
    duration = round(perf_counter() - started, 3)
    if records is not None:
        records.append({"label": label, "script": script, "args": list(args), "status": "passed" if code == 0 else "failed", "return_code": code, "duration_seconds": duration})
    if code:
        print(f"FACTORY_STAGE={label}_FAILED code={code} duration={duration}s")
        return False
    print(f"FACTORY_STAGE={label}_PASSED duration={duration}s")
    return True


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
    run_started_at = datetime.now(timezone.utc).isoformat()
    active_products = active_product_ids()
    is_finalized_target = args.product in finalized_product_ids() if args.product else False

    if not run_stage("AUTO_EVALUATION", "evaluate_excel_products.py", "--no-open", records=stage_records):
        failures.append("auto_evaluation")
    architecture_args = ("--product", str(args.product)) if args.product else ()
    architecture_ok = True
    if not is_finalized_target and not run_stage("DESIGN_ARCHITECTURE", "validate_design_architecture.py", *architecture_args, records=stage_records):
        failures.append("design_architecture")
        architecture_ok = False
    generation_args = ("--product", str(args.product)) if args.product else ()
    if architecture_ok and not run_stage("APP_GENERATION", "generate_autonomous_apps.py", *generation_args, records=stage_records):
        failures.append("app_generation")
    elif not architecture_ok:
        print("FACTORY_STAGE=APP_GENERATION_SKIPPED reason=design_architecture")

    target_products = (args.product,) if args.product else active_products
    if not failures:
        for product_id in target_products:
            semantic_label = f"PRODUCT_{product_id}_SEMANTIC_CONTRACT"
            if not run_stage(semantic_label, "audit_design_contract.py", "--product", str(product_id), records=stage_records):
                failures.append(semantic_label)
                continue
            native_label = f"PRODUCT_{product_id}_NATIVE_XLSX_PREFLIGHT"
            if not run_stage(native_label, "preflight_visual_contract.py", "--product", str(product_id), records=stage_records):
                failures.append(native_label)
    else:
        stage_records.append({"label": "NATIVE_XLSX_QA", "script": "preflight_visual_contract.py", "args": [], "status": "skipped", "return_code": None, "duration_seconds": 0.0, "reason": "upstream_quality_gate"})
        print("FACTORY_STAGE=NATIVE_XLSX_QA_SKIPPED reason=upstream_quality_gate")

    if failures:
        stage_records.append({"label": "DRAFT_PACKAGE", "script": "prepare_autonomous_drafts.py", "args": [], "status": "skipped", "return_code": None, "duration_seconds": 0.0, "reason": "upstream_quality_gate"})
        print("FACTORY_STAGE=DRAFT_PACKAGE_SKIPPED reason=upstream_quality_gate")
    elif not run_stage("DRAFT_PACKAGE", "prepare_autonomous_drafts.py", *generation_args, records=stage_records):
        failures.append("draft_package")

    final_status = "blocked" if failures else "ready_for_draft_upload"
    (ROOT / "outputs" / "listings" / "factory_run_report.json").write_text(
        json.dumps({"schema_version": "2.0.0", "started_at": run_started_at, "completed_at": datetime.now(timezone.utc).isoformat(), "status": final_status, "active_products": list(active_products), "target_product": args.product, "failures": failures, "stages": stage_records, "total_duration_seconds": round(sum(item["duration_seconds"] for item in stage_records), 3)}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    if not args.no_open:
        open_report()
    if failures:
        print(f"AUTONOMOUS_FACTORY=BLOCKED stages={','.join(failures)}")
        return 1
    print("AUTONOMOUS_FACTORY=READY_FOR_DRAFT_UPLOAD")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
