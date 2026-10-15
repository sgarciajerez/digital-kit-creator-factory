from __future__ import annotations

import re
from pathlib import Path


def _tabs_from_markdown(path: Path) -> list[str]:
    if not path.is_file():
        return []
    content = path.read_text(encoding="utf-8")
    section = re.search(r"## Tabs \(order\)(.*?)(?=\n## |\Z)", content, flags=re.S)
    return re.findall(r"^- `([^`]+)`$", section.group(1), flags=re.M) if section else []


def _native_package(folder: Path, spec: dict) -> dict:
    contract_name = str(spec.get("model_contract_file", "model_contract.json"))
    guide_name = str(spec.get("model_behavior_file", "model_behavior.md"))
    app_guide_name = str(spec.get("app_user_guide_file", "app/README.md"))
    guide_pdf_name = str(spec.get("app_user_guide_pdf_filename", f"{spec.get('slug', 'excel')}-user-guide.pdf"))
    workbook_files = sorted((folder / "app").glob("*.xlsx"))
    return {
        "master_workbook_exists": (folder.parents[2] / str(spec.get("master_workbook", "templates/excel-first-base/master_workbook.xlsx"))).is_file(),
        "model_contract_exists": (folder / contract_name).is_file(),
        "model_behavior_exists": (folder / guide_name).is_file(),
        "app_user_guide_exists": (folder / app_guide_name).is_file(),
        "app_user_guide_pdf_exists": (folder / "app" / guide_pdf_name).is_file(),
        "final_workbook_exists": bool(workbook_files),
        "final_workbook": workbook_files[0].name if workbook_files else None,
    }


def _contract(name: str, checks: dict[str, bool], evidence: list[str]) -> dict:
    passed = all(checks.values())
    score = round(100 * sum(checks.values()) / len(checks)) if checks else 0
    return {"name": name, "status": "passed" if passed else "failed", "score": score, "checks": checks, "evidence": evidence}


def validate_product_rules(folder: Path, spec: dict) -> dict:
    app_spec_path = folder / "app_design_spec.md"
    app_markdown = app_spec_path.read_text(encoding="utf-8") if app_spec_path.is_file() else ""
    tabs = spec.get("tabs", [])
    markdown_tabs = _tabs_from_markdown(app_spec_path)
    native = _native_package(folder, spec)
    chart = spec.get("signature_chart", {})
    lineage = spec.get("source_lineage", [])

    architecture_checks = {
        "app_design_spec_exists": app_spec_path.is_file(),
        "markdown_contains_product_analysis": "## Product analysis" in app_markdown,
        "markdown_tabs_match_json": markdown_tabs == tabs and bool(tabs),
        "native_model_declared": bool(spec.get("model_base")),
        "source_lineage_declared": bool(lineage),
        "native_contract_mode": spec.get("native_reference_mode") == "master_workbook_contract",
        "master_workbook_exists": native["master_workbook_exists"],
        "signature_chart_declared": bool(chart.get("type") and chart.get("purpose")),
    }
    architecture = _contract(
        "architecture",
        architecture_checks,
        ["The Markdown brief defines the application flow.", "The source lineage and master_workbook.xlsx define the design input."],
    )

    feature_count = len(spec.get("functional_scope", []))
    functionality_checks = {
        "functional_scope_is_substantial": feature_count >= 7,
        "onboarding_dashboard_settings_flow": {"Start here", "Dashboard", "Settings"}.issubset(tabs),
        "native_package_exists": native["final_workbook_exists"] and native["model_contract_exists"] and native["model_behavior_exists"] and native["app_user_guide_exists"] and native["app_user_guide_pdf_exists"],
        "native_structure_gate_declared": "master_workbook.xlsx" in app_markdown and "source of truth" in app_markdown.casefold(),
    }
    functionality_evidence = [f"Documented capabilities: {feature_count}"]
    functionality_evidence.append(
        f"Workbook detected: {native['final_workbook']}" if native["final_workbook_exists"] else "No final XLSX yet."
    )
    functionality = _contract("functionality", functionality_checks, functionality_evidence)

    visual_checks = {
        "visual_identity_declared": bool(spec.get("visual_identity", {}).get("name") and spec.get("visual_identity", {}).get("motif")),
        "five_colour_palette_declared": len({str(spec.get("palette", {}).get(key)).casefold() for key in ("background", "ink", "accent", "secondary", "highlight")}) == 5,
        "master_workbook_present": native["master_workbook_exists"],
        "native_geometry_required": bool(spec.get("visual_acceptance", {}).get("native_geometry_required")),
        "source_model_match_required": bool(spec.get("visual_acceptance", {}).get("source_model_match_required")),
    }
    native_contract = _contract(
        "native_structure",
        visual_checks,
        ["The XLSX itself is the visual reference; geometry is checked from cells, tables and native objects."],
    )
    contracts = {"architecture": architecture, "functionality": functionality, "native_structure": native_contract}
    return {
        "all_passed": all(item["status"] == "passed" for item in contracts.values()),
        "contracts": contracts,
        "native_package": native,
        "workbook_level_checks": "scheduled_for_functional_and_structural_qa",
    }
