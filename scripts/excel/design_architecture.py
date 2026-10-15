from __future__ import annotations

"""Shared design-system contract used by generation and validation stages."""

import json
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[2]
SYSTEM_PATH = ROOT / "config" / "FACTORY_DESIGN_SYSTEM.json"


def load_design_system() -> dict[str, Any]:
    return json.loads(SYSTEM_PATH.read_text(encoding="utf-8"))


def classify_screen(tab: str) -> str:
    normalized = tab.casefold().strip()
    if normalized in {"instrucciones", "guía", "guia"}:
        return "guide"
    if normalized in {"vista previa"}:
        return "dashboard"
    if normalized in {"configuración", "configuracion", "ajustes"}:
        return "settings"
    if normalized in {"revisión", "revision"}:
        return "checklist"
    if normalized in {"metas", "deudas"}:
        return "analysis"
    if normalized in {"start here", "inicio", "comenzar"}:
        return "guide"
    if normalized in {"dashboard", "resumen", "overview"}:
        return "dashboard"
    if normalized in {"settings", "configuration", "configuración", "configuracion"}:
        return "settings"
    if any(token in normalized for token in ("checklist", "tasks", "tareas")):
        return "checklist"
    if any(token in normalized for token in ("simulator", "simulador", "forecast", "progress", "progreso", "monthly view", "year end")):
        return "analysis"
    return "table"


def build_screen_contracts(tabs: list[str], system: dict[str, Any] | None = None) -> dict[str, dict[str, Any]]:
    system = system or load_design_system()
    contracts: dict[str, dict[str, Any]] = {}
    for tab in tabs:
        template = classify_screen(tab)
        template_data = system["templates"][template]
        contracts[tab] = {
            "template": template,
            "components": list(template_data["required_components"]),
            "overrides": {},
            "complexity_score": system["complexity_budget"]["target_score"],
            "content": {
                "role": f"{template} screen for {tab}",
                "data_source": "workbook_model",
                "geometry_source": "excel_native_grid",
            },
        }
    return contracts


def apply_design_architecture(spec: dict[str, Any]) -> dict[str, Any]:
    """Add stable contract metadata without changing product-specific content."""
    system = load_design_system()
    tabs = [str(tab) for tab in spec.get("tabs", [])]
    existing = spec.get("screen_contracts") or {}
    generated = build_screen_contracts(tabs, system)
    for tab, contract in existing.items():
        if tab in generated and isinstance(contract, dict):
            generated[tab]["content"].update(contract.get("content") or {})
            generated[tab]["overrides"] = dict(contract.get("overrides") or {})
            declared_components = list(contract.get("components") or [])
            generated[tab]["components"] = list(dict.fromkeys(generated[tab]["components"] + [
                component for component in declared_components if component not in generated[tab]["components"]
            ]))
            if contract.get("chart"):
                if "chart" not in generated[tab]["components"]:
                    generated[tab]["components"].append("chart")
                generated[tab]["chart"] = dict(contract["chart"])
    spec["design_system_contract"] = "config/FACTORY_DESIGN_SYSTEM.json"
    spec["design_system_version"] = system["schema_version"]
    spec["template_library"] = system["template_library"]
    spec["layout_mode"] = "template_inheritance"
    spec["design_order"] = "rules/excel/FACTORY_DESIGN_ORDER_SIMPLE_EXCEL.md"
    # Keep the legacy approach label for downstream compatibility, while the
    # active source of truth is now the native master workbook.
    spec["design_approach"] = "excel_first_simple_table_system"
    spec["native_base"] = "templates/excel-first-base/master_workbook.xlsx"
    spec["native_base_contract"] = "templates/excel-first-base/native_base_contract.json"
    spec.setdefault("model_base", "templates/excel-first-base")
    spec.setdefault("source_lineage", ["templates/sources"])
    spec["native_reference_mode"] = "master_workbook_contract"
    spec["master_workbook"] = "templates/excel-first-base/master_workbook.xlsx"
    # Master-derived products are replicas of the native workbook, not
    # responsive layouts. Content may vary, but the visual slots do not.
    spec["native_geometry_policy"] = "exact_master_grid"
    spec["native_table_slot_policy"] = "same_master_ranges"
    spec["visual_acceptance"] = {
        "human_review_required": True,
        "native_geometry_required": True,
        "source_model_match_required": True,
        "native_shapes_only": True,
        "chart_objects_allowed": True,
    }
    spec["signature_visual"] = {
        "type": "native_summary_or_optional_excel_chart",
        "chart_objects_allowed": True,
        "purpose": "Communicate the main decision with cells, formulas, conditional formatting or one declared Excel chart.",
    }
    spec["visual_complexity"] = {
        "target_score": system["complexity_budget"]["target_score"],
        "maximum_score": system["complexity_budget"]["maximum_score"],
        "scale": system["complexity_budget"]["scale"],
        "reason": "Utility and Excel replication take priority over decorative complexity.",
    }
    spec["screen_contracts"] = generated
    spec["navigation_contract"] = {
        "tabs": tabs,
        "component": system["navigation"]["component"],
        "button_width_mode": system["navigation"]["button_width_mode"],
        "gap_px": system["navigation"]["gap_px"],
        "frame_edges": "inherited_from_outer_padding",
    }
    spec["global_design_system"] = "rules/excel/DESIGN_SYSTEM.md"
    spec["model_integrity_contract"] = {
        "semantic_headers_required": True,
        "generic_headers_forbidden": True,
        "totals_outside_data_ranges": True,
        "semantic_number_formats_required": True,
        "functional_cases_required_for_formula_models": True,
        "validator": "scripts/core/model_quality.py",
    }
    spec["model_behavior_file"] = str(spec.get("model_behavior_file", "model_behavior.md"))
    spec["model_contract_file"] = str(spec.get("model_contract_file", "model_contract.json"))
    spec["app_user_guide_file"] = str(spec.get("app_user_guide_file", "app/README.md"))
    spec["app_user_guide_pdf_filename"] = str(spec.get("app_user_guide_pdf_filename", f"{spec.get('slug', 'excel')}-user-guide.pdf"))
    spec["model_behavior_source"] = "declared_model_contract_before_xlsx"
    spec["user_model_guide"] = {
        "file": spec["model_behavior_file"],
        "purpose": "Guía de uso y de verificación funcional independiente del XLSX.",
        "edit_inputs_only": True,
        "fail_on_contract_mismatch": True,
    }
    return spec


def validate_design_spec(spec: dict[str, Any], system: dict[str, Any] | None = None) -> list[str]:
    system = system or load_design_system()
    errors: list[str] = []
    tabs = list(spec.get("tabs") or [])
    palette = spec.get("palette")
    if palette is not None:
        required_palette = ["background", "ink", "accent", "secondary", "highlight"]
        missing_palette = [token for token in required_palette if not palette.get(token)]
        if missing_palette:
            errors.append(f"palette incompleta; faltan {missing_palette}")
        elif len({str(palette[token]).casefold() for token in required_palette}) != len(required_palette):
            errors.append("palette debe declarar cinco colores distintos")
    if int(spec.get("product_id", 0) or 0) >= 12 and spec.get("source_of_truth"):
        mock_profile = spec.get("mock_profile") or {}
        areas = list(mock_profile.get("areas") or [])
        suppliers = list(mock_profile.get("suppliers") or [])
        if not mock_profile.get("project_name") or len(areas) < 7 or len(suppliers) < 7:
            errors.append("todo producto derivado del master requiere mock_profile coherente con proyecto, areas y proveedores")
        elif len(set(areas)) != len(areas) or len(set(suppliers)) != len(suppliers):
            errors.append("areas y proveedores mock no pueden repetirse")
    contracts = spec.get("screen_contracts")
    if not tabs or len(tabs) != len(set(tabs)):
        errors.append("tabs debe ser una lista no vacía y sin duplicados")
    if spec.get("design_system_contract") != "config/FACTORY_DESIGN_SYSTEM.json":
        errors.append("falta la referencia al contrato config/FACTORY_DESIGN_SYSTEM.json")
    if spec.get("design_system_version") != system.get("schema_version"):
        errors.append("design_system_version no coincide con la versiÃ³n activa")
    if spec.get("template_library") != system.get("template_library"):
        errors.append("falta la referencia a la biblioteca de plantillas Excel")
    if spec.get("layout_mode") != "template_inheritance":
        errors.append("layout_mode debe ser template_inheritance")
    if not isinstance(contracts, dict) or list(contracts) != tabs:
        errors.append("screen_contracts debe contener exactamente las pestañas y en el mismo orden")
        contracts = contracts if isinstance(contracts, dict) else {}
    allowed_components = set(system["allowed_components"])
    allowed_overrides = set(system["allowed_screen_overrides"])
    for tab in tabs:
        contract = contracts.get(tab)
        if not isinstance(contract, dict):
            errors.append(f"{tab}: falta screen_contract")
            continue
        template = contract.get("template")
        template_data = system["templates"].get(template)
        if not template_data:
            errors.append(f"{tab}: plantilla desconocida {template!r}")
            continue
        components = list(contract.get("components") or [])
        missing = set(template_data["required_components"]) - set(components)
        unknown = set(components) - allowed_components
        if missing:
            errors.append(f"{tab}: faltan componentes obligatorios {sorted(missing)}")
        if unknown:
            errors.append(f"{tab}: componentes no permitidos {sorted(unknown)}")
        supported = set(template_data["required_components"]) | set(template_data["optional_components"])
        unsupported = set(components) - supported
        if unsupported:
            errors.append(f"{tab}: componentes no permitidos por la plantilla {sorted(unsupported)}")
        if len(components) > system["complexity_budget"]["maximum_components_per_screen"]:
            errors.append(f"{tab}: supera el maximo de componentes simples")
        chart_components = [item for item in components if item == "chart"]
        if len(chart_components) > system["charts"]["maximum_per_screen"]:
            errors.append(f"{tab}: supera el maximo de grÃ¡ficos por pantalla")
        if "chart" in components:
            chart = contract.get("chart")
            if not isinstance(chart, dict):
                errors.append(f"{tab}: chart requiere una especificaciÃ³n declarada")
            elif chart.get("type") not in system["charts"]["allowed_types"]:
                errors.append(f"{tab}: tipo de grÃ¡fico no permitido")
            elif not chart.get("anchor") or not chart.get("values"):
                errors.append(f"{tab}: el grÃ¡fico debe declarar anchor y values")
        elif contract.get("chart"):
            errors.append(f"{tab}: existe chart pero no estÃ¡ declarado como componente")
        score = contract.get("complexity_score", system["complexity_budget"]["target_score"])
        if not isinstance(score, (int, float)) or not 1 <= score <= system["complexity_budget"]["maximum_score"]:
            errors.append(f"{tab}: complexity_score fuera del rango permitido")
        overrides = contract.get("overrides") or {}
        if not isinstance(overrides, dict) or not set(overrides).issubset(allowed_overrides):
            errors.append(f"{tab}: contiene overrides fuera del contrato permitido")
        if contract.get("content", {}).get("geometry_source") != "excel_native_grid":
            errors.append(f"{tab}: geometry_source debe ser excel_native_grid")
    navigation = spec.get("navigation_contract") or {}
    model_base = spec.get("model_base")
    if not isinstance(model_base, str) or not (ROOT / model_base).exists():
        errors.append("model_base debe apuntar a un modelo nativo existente")
    lineage = spec.get("source_lineage")
    if not isinstance(lineage, list) or not lineage:
        errors.append("source_lineage debe declarar al menos un source")
    else:
        for source in lineage:
            if not isinstance(source, str):
                errors.append("source_lineage contiene una ruta no valida")
                continue
            source_path = ROOT / source
            if not source_path.exists() and not any(ROOT.glob(source)):
                errors.append(f"source_lineage no existe: {source}")
    if spec.get("native_reference_mode") != "master_workbook_contract":
        errors.append("native_reference_mode debe ser master_workbook_contract")
    native_sheet_map = spec.get("native_sheet_map")
    is_new_master_product = int(spec.get("product_id", 0) or 0) >= 12 and bool(spec.get("source_of_truth"))
    # Every new product generated under the master contract must carry an
    # explicit one-to-one source map. Without it, the preflight cannot compare
    # renamed product screens against the real master screens.
    if is_new_master_product and not isinstance(native_sheet_map, dict):
        errors.append("todo producto nuevo derivado del master requiere native_sheet_map completo")
    if is_new_master_product and not spec.get("preserve_native_master_objects"):
        errors.append("todo producto nuevo derivado del master debe conservar preserve_native_master_objects=true")
    if is_new_master_product and spec.get("skip_native_overlay"):
        errors.append("un producto derivado del master no puede omitir el overlay nativo")
    if native_sheet_map is not None:
        if not isinstance(native_sheet_map, dict) or list(native_sheet_map) != tabs:
            errors.append("native_sheet_map debe mapear cada pestaña del producto en el mismo orden")
        elif len(set(native_sheet_map.values())) != len(native_sheet_map):
            errors.append("native_sheet_map no puede reutilizar una hoja maestra en dos pestañas")
    master_workbook = spec.get("master_workbook", "")
    if not isinstance(master_workbook, str) or not (ROOT / master_workbook).is_file():
        errors.append("master_workbook debe apuntar a templates/excel-first-base/master_workbook.xlsx")
    if spec.get("preserve_native_master_objects"):
        if spec.get("native_geometry_policy") != "exact_master_grid":
            errors.append("un producto derivado del master debe usar native_geometry_policy=exact_master_grid")
        if spec.get("native_table_slot_policy") != "same_master_ranges":
            errors.append("las tablas derivadas deben conservar los rangos del master")
        adaptations = spec.get("native_surface_adaptations") or {}
        if any(adaptations.get(key) for key in ("allow_column_width_changes", "merged_slot_resizes", "fit_table_columns_to_content")):
            errors.append("un producto derivado del master no puede adaptar anchos, merges ni ranuras de tablas")
        if not isinstance(native_sheet_map, dict) or list(native_sheet_map) != tabs or len(set(native_sheet_map.values())) != len(native_sheet_map):
            errors.append("un producto derivado del master requiere un native_sheet_map completo, ordenado y sin duplicados")
    if navigation.get("tabs") != tabs:
        errors.append("navigation_contract no coincide con el orden de tabs")
    if navigation.get("button_width_mode") != system["navigation"]["button_width_mode"]:
        errors.append("la navegación debe derivar el ancho de los botones del frame")
    if navigation.get("component") != system["navigation"]["component"]:
        errors.append("la navegación debe usar el componente compartido")
    if spec.get("design_order") != "rules/excel/FACTORY_DESIGN_ORDER_SIMPLE_EXCEL.md":
        errors.append("falta la orden de diseÃ±o Excel-first simplificada")
    complexity = spec.get("visual_complexity") or {}
    if complexity.get("maximum_score", 0) > system["complexity_budget"]["maximum_score"]:
        errors.append("el presupuesto de complejidad del producto supera el mÃ¡ximo")
    declared_charts = sum(
        1
        for contract in contracts.values()
        if isinstance(contract, dict) and "chart" in (contract.get("components") or [])
    )
    if declared_charts > system["charts"]["maximum_per_workbook"]:
        errors.append("el producto supera el mÃ¡ximo de grÃ¡ficos declarados")
    return errors
