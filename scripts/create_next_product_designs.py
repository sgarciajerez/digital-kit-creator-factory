from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from design_architecture import apply_design_architecture
from factory_config import active_product_ids, all_product_ids
from model_contracts import contract_for_product, validate_contract_schema, write_model_behavior_markdown, write_model_contract_json


ROOT = Path(__file__).resolve().parent.parent
OUT_ROOT = ROOT / "outputs" / "listings"


PRODUCTS = [
    {
        "id": 1,
        "slug": "community-garden-planner",
        "title": "Community Garden Budget & Project Planner",
        "subtitle": "A connected Excel system for planning shared garden costs, suppliers, volunteers and seasonal milestones.",
        "target_buyer": "Neighbourhood groups and small community gardens coordinating improvements, purchases and shared responsibilities.",
        "price_test": 7.99,
        "palette": {"background": "#F2F5EC", "ink": "#2F3B32", "accent": "#5C8058", "secondary": "#D17A5A", "highlight": "#C5A34B"},
        "visual_identity": {"name": "Shared harvest", "motif": "A calm leaf-green grid with terracotta markers for shared progress", "usage": "Native workbook accents, table headers and dashboard emphasis"},
        "signature_chart": {"type": "Garden budget remaining and milestone readiness summary", "purpose": "Keep shared project spend, supplier decisions and seasonal tasks visible in one practical view"},
        "tabs": ["Start here", "Dashboard", "Budget", "Purchases", "Quotes", "Timeline", "Settings"],
        "features": ["Budget by garden area", "Actual spend linked to purchase entries", "Supplier quote comparison", "Materials and equipment purchase log", "Seasonal timeline with due dates", "Contingency and date settings", "Dashboard cue for the next garden decision"],
        "mock_profile": {
            "project_name": "Greenway community garden",
            "areas": ["Site preparation", "Soil and compost", "Plants and seeds", "Irrigation", "Tools and storage", "Workshops", "Contingency"],
            "suppliers": ["Groundworks Co", "Harvest Soil", "Seed House", "Waterwise Supply", "Tool Shed", "Neighbourhood Hall", "Garden Reserve"],
            "items": ["Raised bed preparation", "Compost delivery", "Seasonal seed order", "Watering system", "Shared tool set", "Workshop materials", "Reserve allocation"],
            "locations": ["Garden", "Delivery", "Nursery", "Beds", "Tool shed", "Community hall", "Project"],
            "phases": ["Preparation", "Groundwork", "Planting", "Irrigation", "Equipment", "Community", "Finance"],
            "tasks": ["Approve site layout", "Order soil delivery", "Plan seasonal planting", "Install water points", "Label shared tools", "Book volunteer workshop", "Review reserve", "Plant first beds", "Check irrigation", "Review monthly spend", "Archive receipts", "Close season plan"],
            "quote_notes": ["Includes site visit", "Delivery included", "Seasonal stock confirmed", "Installation support included", "Shared-use warranty", "Room hire included", "Held for weather changes"],
        },
        "sources": ["templates/sources/planilla-de-excel-para-presupuesto-de-obra.xlsx", "templates/sources/cronograma-de-actividades-en-excel.xlsx"],
        "model_base": "templates/model-library/family-budget-sources-base",
        "native_sheet_map": {"Start here": "Inicio", "Dashboard": "Resumen", "Budget": "Presupuesto", "Purchases": "Gastos reales", "Quotes": "Movimientos", "Timeline": "Flujo de caja", "Settings": "Listas"},
        "preserve_native_master_objects": True,
        "native_surface_adaptations": {"fit_table_columns_to_content": False, "allow_column_width_changes": [], "merged_slot_resizes": {}},
    },
    {
        "id": 8,
        "slug": "wedding-planner",
        "title": "Wedding Planner & Budget Dashboard",
        "subtitle": "A connected Excel planning system for budget, guests, vendors, payments and checklist progress.",
        "target_buyer": "Engaged couples, DIY wedding planners and families coordinating wedding costs.",
        "price_test": 7.99,
        "palette": {"background": "#FAF7F2", "ink": "#3A3040", "accent": "#B9787A", "secondary": "#7A9B89", "highlight": "#D3A94C"},
        "visual_identity": {"name": "Rings & ribbon", "motif": "Interlocking wedding rings with a small champagne-gold sparkle", "usage": "Header badge, subtle watermark and cover accent"},
        "signature_chart": {"type": "Budget allocation plus countdown timeline", "purpose": "Make wedding progress tangible while keeping spend categories readable"},
        "tabs": ["Start here", "Dashboard", "Budget", "Guests", "Vendors", "Payments", "Checklist", "Settings"],
        "features": ["Budget versus actual by category", "Vendor directory with deposits and balance due", "Guest list and RSVP tracking", "Meal and dietary notes", "Payment timeline", "Wedding checklist with progress", "Countdown dashboard"],
        "sources": ["templates/sources/planilla-de-excel-de-presupuesto-de-fiesta.xlsx", "templates/sources/planilla-de-excel-de-presupuesto-de-casamiento.xlsx"],
        "model_base": "templates/model-library/wedding-planner-base",
    },
    {
        "id": 9,
        "slug": "debt-payoff-planner",
        "title": "Debt Payoff Planner & Strategy Simulator",
        "subtitle": "An Excel system that compares snowball and avalanche strategies, tracks payments and shows a projected debt-free date.",
        "target_buyer": "People organising credit cards, loans and repayment goals who want a visual, practical plan.",
        "price_test": 6.99,
        "palette": {"background": "#F3F7F8", "ink": "#163047", "accent": "#2D8F8B", "secondary": "#EF806C", "highlight": "#D0A83E"},
        "visual_identity": {"name": "Path to zero", "motif": "Descending balance path with coral payment milestones", "usage": "Header badge, payoff chart marker and progress cover accent"},
        "signature_chart": {"type": "Payoff trajectory with milestone bars", "purpose": "Show the balance falling and let extra payments create an obvious visual difference"},
        "tabs": ["Start here", "Dashboard", "Debts", "Payments", "Simulator", "Progress", "Settings"],
        "features": ["Debt setup with balance, APR and minimum payment", "Snowball versus avalanche comparison", "Payment log and principal/interest split", "Extra payment what-if simulator", "Projected payoff date", "Progress dashboard", "Financial disclaimer and assumptions"],
        "sources": ["templates/sources/Programación de la amortización del préstamo.xlsx", "templates/sources/COMPARACIÓN DE HIPOTECAS.xlsx"],
        "model_base": "templates/model-library/family-budget-sources-base",
    },
    {
        "id": 10,
        "slug": "small-business-cashflow",
        "title": "Small Business Cash Flow Dashboard",
        "subtitle": "A clean Excel command centre for freelancers and small businesses to track cash, invoices, recurring costs and runway.",
        "target_buyer": "Freelancers, solopreneurs and small service businesses that need a simple operating view.",
        "price_test": 8.99,
        "palette": {"background": "#F5F7F4", "ink": "#263A35", "accent": "#477A67", "secondary": "#D47B63", "highlight": "#C3A146"},
        "visual_identity": {"name": "Cash-flow pulse", "motif": "Rising bars crossed by a soft coral cash-flow pulse", "usage": "Header badge, dashboard chart cue and runway cover accent"},
        "signature_chart": {"type": "12-month cash-flow bars plus runway indicator", "purpose": "Make seasonality, liquidity and upcoming obligations immediately legible"},
        "tabs": ["Start here", "Dashboard", "Budget", "Purchases", "Quotes", "Timeline", "Settings"],
        "features": ["12-month cash flow and running balance", "Income and expense categories", "Invoice tracker with due dates and status", "Recurring costs and expected income", "30/60/90-day forecast", "Tax reserve and year-end summary", "Multi-currency settings"],
        "sources": ["templates/sources/Presupuesto de gastos empresariales.xlsx", "templates/sources/planilla-de-excel-de-presupuesto-de-ventas.xlsx"],
        "model_base": "templates/model-library/family-budget-sources-base",
    },
    {
        "id": 11,
        "slug": "meal-planner-grocery-budget",
        "title": "Meal Planner, Grocery List & Food Budget",
        "subtitle": "A practical Excel system that connects weekly meals, recipes, grocery lists, pantry items and food spending.",
        "target_buyer": "Families, meal-preppers and budget-conscious households who want less waste and fewer shopping decisions.",
        "price_test": 7.49,
        "palette": {"background": "#FBF8F0", "ink": "#3E4032", "accent": "#7A8B55", "secondary": "#D98369", "highlight": "#D0A542"},
        "visual_identity": {"name": "Plate & leaf", "motif": "Warm plate, fork and leaf mark for practical meal planning", "usage": "Header badge, recipe cover accent and budget callout"},
        "signature_chart": {"type": "Weekly meal coverage grid plus grocery spend ring", "purpose": "Connect meal completeness, shopping effort and food budget in one visual story"},
        "tabs": ["Start here", "Dashboard", "This week", "Monthly view", "Grocery list", "Recipes", "Pantry", "Budget", "Settings"],
        "features": ["Weekly meal planner with dropdown recipes", "Consolidated grocery list by aisle", "Recipe vault with cost per serving", "Pantry and expiry tracking", "Food budget versus actual", "Leftovers and meal-prep notes", "Optional nutrition fields with disclaimer"],
        "sources": ["templates/sources/planilla-de-excel-de-presupuesto-de-navidad.xlsx", "templates/sources/FLUJO DE EFECTIVO PERSONAL.xlsx"],
        "model_base": "templates/model-library/family-budget-sources-base",
    },
    {
        "id": 12,
        "slug": "home-renovation-planner",
        "title": "Home Renovation Budget & Project Planner",
        "subtitle": "A practical Excel system to connect renovation costs, purchases, quotes and the project timeline in one calm view.",
        "target_buyer": "Homeowners and DIY renovators coordinating a room refresh or small renovation without losing control of the budget.",
        "price_test": 7.99,
        "palette": {"background": "#F8F4EF", "ink": "#3E3530", "accent": "#C4775A", "secondary": "#6F8A73", "highlight": "#D5A84F"},
        "visual_identity": {"name": "Measured home", "motif": "Terracotta roofline paired with a sage measuring square", "usage": "Header mark, budget cue and product cover identity"},
        "signature_chart": {"type": "Budget remaining and task readiness summary", "purpose": "Keep the renovation decision focused on remaining spend and the next actionable task"},
        "tabs": ["Start here", "Dashboard", "Budget", "Purchases", "Quotes", "Timeline", "Settings"],
        "features": ["Budget by renovation area", "Actual spend linked to purchase entries", "Available budget after committed spend", "Purchase log with paid and status fields", "Quote comparison by supplier and area", "Timeline with due dates and days left", "Contingency and date settings", "Dashboard cues for the next decision"],
        "sources": ["templates/sources/planilla-de-excel-para-presupuesto-de-obra.xlsx", "templates/sources/planilla-de-excel-de-presupuesto-de-obra-de-vivienda.xlsx", "templates/sources/cronograma-de-actividades-en-excel.xlsx", "templates/sources/planilla-de-excel-de-planeamiento-de-compras.xlsx"],
        "model_base": "templates/model-library/family-budget-sources-base",
        "native_sheet_map": {"Start here": "Inicio", "Dashboard": "Resumen", "Transactions": "Presupuesto", "Invoices": "Gastos reales", "Recurring": "Movimientos", "Forecast": "Flujo de caja", "Year end": "Metas y fondo", "Settings": "Listas"},
        "mock_profile": {
            "project_name": "Oak Street renovation plan",
            "areas": ["Kitchen units", "Flooring", "Lighting", "Plumbing", "Paint", "Permits", "Contingency"],
            "suppliers": ["Oak & Co", "Stone House", "Bright Home", "Flow Trade", "Colour Room", "City office", "Project reserve"],
        },
        "preserve_native_master_objects": True,
        "native_surface_adaptations": {
            "fit_table_columns_to_content": False,
            "allow_column_width_changes": [],
            "merged_slot_resizes": {},
        },
    },
    {
        "id": 13,
        "slug": "freelance-client-invoice-tracker",
        "title": "Freelance Client, Project & Invoice Tracker",
        "subtitle": "A focused Excel system for tracking client work, invoices, recurring costs and cash runway without losing the next action.",
        "target_buyer": "Freelancers and independent consultants managing several clients, project payments and monthly operating costs.",
        "price_test": 7.99,
        "palette": {"background": "#F3F7F5", "ink": "#203A35", "accent": "#2F8275", "secondary": "#D9785F", "highlight": "#D1A64A"},
        "visual_identity": {"name": "Client signal", "motif": "A teal project signal with a coral invoice marker", "usage": "Header identity, invoice cue and dashboard accent"},
        "signature_chart": {"type": "Invoice timing and cash runway summary", "purpose": "Make client cash timing and the next follow-up visible before accepting more work"},
        "tabs": ["Start here", "Dashboard", "Budget", "Purchases", "Quotes", "Timeline", "Settings"],
        "features": ["Client income and expense log", "Invoice due-date tracker", "Recurring cost register", "90-day cash forecast", "Year-end income and expense summary", "Runway decision cue", "Status lists for consistent filtering"],
        "sources": ["templates/sources/Presupuesto de gastos empresariales.xlsx", "templates/sources/planilla-de-excel-de-presupuesto-de-ventas.xlsx"],
        "model_base": "templates/model-library/family-budget-sources-base",
        "native_sheet_map": {"Start here": "Inicio", "Dashboard": "Resumen", "Budget": "Presupuesto", "Purchases": "Gastos reales", "Quotes": "Movimientos", "Timeline": "Flujo de caja", "Settings": "Listas"},
        "mock_profile": {
            "project_name": "Client signal cash plan",
            "areas": ["Client retainers", "Delivery", "Software", "Marketing", "Workspace", "Travel", "Tax reserve"],
            "suppliers": ["Northstar Client", "Studio Partner", "Cloud Tools", "Signal Ads", "Desk House", "Rail & Air", "Tax Office"],
        },
        "preserve_native_master_objects": True,
        "native_surface_adaptations": {"fit_table_columns_to_content": False, "allow_column_width_changes": [], "merged_slot_resizes": {}},
    },
    {
        "id": 14,
        "slug": "creator-content-sponsorship-tracker",
        "title": "Creator Content & Sponsorship Tracker",
        "subtitle": "A focused Excel system for planning content, tracking sponsorship income and keeping creator cash visible.",
        "target_buyer": "Content creators and solo media businesses managing publishing plans, brand deals and recurring creator costs.",
        "price_test": 7.99,
        "palette": {"background": "#F5F3F8", "ink": "#302B42", "accent": "#7257A6", "secondary": "#E27A62", "highlight": "#D5AC4A"},
        "visual_identity": {"name": "Signal studio", "motif": "A violet publishing signal with a coral sponsorship marker", "usage": "Header identity, sponsorship cue and dashboard accent"},
        "signature_chart": {"type": "Content cadence and sponsorship runway summary", "purpose": "Make publishing rhythm and upcoming brand income visible before committing to new work"},
        "tabs": ["Start here", "Dashboard", "Budget", "Purchases", "Quotes", "Timeline", "Settings"],
        "features": ["Content and creator income log", "Sponsorship invoice tracker", "Recurring creator costs", "90-day cash forecast", "Year-end creator income summary", "Runway decision cue", "Status lists for consistent filtering"],
        "mock_profile": {
            "project_name": "Signal Studio creator plan",
            "areas": ["Brand deals", "Video production", "Editing", "Software", "Studio", "Travel", "Tax reserve"],
            "suppliers": ["Brightline Media", "North Studio", "Frame & Co", "Cloud Tools", "Desk House", "Creator Travel", "Tax Office"],
        },
        "sources": ["templates/sources/Presupuesto de gastos empresariales.xlsx", "templates/sources/planilla-de-excel-de-presupuesto-de-ventas.xlsx"],
        "model_base": "templates/model-library/family-budget-sources-base",
        "native_sheet_map": {"Start here": "Inicio", "Dashboard": "Resumen", "Budget": "Presupuesto", "Purchases": "Gastos reales", "Quotes": "Movimientos", "Timeline": "Flujo de caja", "Settings": "Listas"},
        "preserve_native_master_objects": True,
        "native_surface_adaptations": {"fit_table_columns_to_content": False, "allow_column_width_changes": [], "merged_slot_resizes": {}},
    },
]


def write_app_design_spec(product: dict, path: Path, model_contract: dict) -> None:
    tabs = "\n".join(f"- `{label}`" for label in product["tabs"])
    features = "\n".join(f"- {feature}" for feature in product["features"])
    identity = product["visual_identity"]
    chart = product["signature_chart"]
    content = f"""# App design specification · {product['title']}

## Product analysis

- Product: `{product['slug']}`
- Target buyer: {product['target_buyer']}
- Design language: English baseline; Spanish localization preserves the same master grid.
- Visual direction: simple Excel-first workbook with clear tables, compact summaries and restrained accents.
- Design priority: utility, replicability, readability, then visual attractiveness.
- Complexity target: 2/10; hard maximum: 4/10.
- Global source: `rules/FACTORY_DESIGN_ORDER_SIMPLE_EXCEL.md`

## Visual identity

- Identity name: **{identity['name']}**
- Signature motif: {identity['motif']}
- Usage: {identity['usage']}
- The visual identity is implemented through the declared five-colour palette,
  worksheet accents, table surfaces and coherent mock data. No external icon
  asset is required for the Excel product.

## Signature visual

- Requested visual: **{chart['type']}**
- Purpose: {chart['purpose']}
- Prefer a native summary row, formula result, progress bar or conditional formatting.

## Tabs (order)

{tabs}

## Native Excel design contract

- Source of truth: `templates/excel-first-base/master_workbook.xlsx`.
- The master defines the grid, dimensions, table anchors, navigation and shared slots.
- Product changes are limited to declared text, columns, formulas, palette, identity and data slots.
- No image or browser representation is required for approval.

## Functional scope

{features}

## Functional model guide

- Normative guide: `model_behavior.md`.
- Machine-readable contract: `model_contract.json`.
- The guide is written from expected behaviour before the XLSX is generated.
- Edit input columns only, preserve derived columns and run every verification case after changing a source value.
- Contract identifier: `{model_contract['contract_id']}`.

## Acceptance gates

- The master workbook exists and is the only visual and structural base.
- The final XLSX preserves the declared grid, table dimensions, navigation, identity and source-model structure.
- The final XLSX passes direct functional and structural audits.
"""
    path.write_text(content, encoding="utf-8")


def read_tabs_from_spec(path: Path) -> list[str]:
    content = path.read_text(encoding="utf-8")
    section = re.search(r"## Tabs \(order\)(.*?)(?=\n## |\Z)", content, flags=re.S)
    if not section:
        raise ValueError(f"No existe la sección de pestañas en {path}")
    return re.findall(r"^- `([^`]+)`$", section.group(1), flags=re.M)


def validate_product_identity_and_mock(product: dict) -> None:
    required_palette = ("background", "ink", "accent", "secondary", "highlight")
    palette = product.get("palette") or {}
    missing = [key for key in required_palette if not palette.get(key)]
    if missing:
        raise ValueError(f"{product['slug']}: faltan colores de identidad: {', '.join(missing)}")
    if len({str(palette[key]).casefold() for key in required_palette}) != len(required_palette):
        raise ValueError(f"{product['slug']}: la paleta debe tener cinco colores distintos")
    if int(product["id"]) < 12:
        return
    profile = product.get("mock_profile") or {}
    areas = list(profile.get("areas") or [])
    suppliers = list(profile.get("suppliers") or [])
    if not profile.get("project_name") or len(areas) < 7 or len(suppliers) < 7:
        raise ValueError(f"{product['slug']}: el mock_profile debe declarar proyecto, siete areas y siete proveedores")
    if len(set(areas)) != len(areas) or len(set(suppliers)) != len(suppliers):
        raise ValueError(f"{product['slug']}: areas y proveedores mock no pueden repetirse")


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera contratos de producto desde master_workbook.xlsx")
    parser.add_argument("--product", type=int, choices=all_product_ids(), help="Generar solo un producto")
    args = parser.parse_args()
    requested = {args.product} if args.product is not None else set(active_product_ids())
    for product in [item for item in PRODUCTS if item["id"] in requested]:
        validate_product_identity_and_mock(product)
        folder = OUT_ROOT / f"product-{product['id']}"
        folder.mkdir(parents=True, exist_ok=True)
        model_contract = contract_for_product(product["id"])
        errors = validate_contract_schema(model_contract)
        if errors:
            raise RuntimeError(f"product-{product['id']}: contrato inválido: {errors}")
        app_spec_path = folder / "app_design_spec.md"
        write_app_design_spec(product, app_spec_path, model_contract)
        write_model_behavior_markdown(folder / "model_behavior.md", model_contract)
        write_model_contract_json(folder / "model_contract.json", model_contract)
        product["tabs"] = read_tabs_from_spec(app_spec_path)
        spec = {
            "product_id": product["id"],
            "slug": product["slug"],
            "title": product["title"],
            "subtitle": product["subtitle"],
            "target_buyer": product["target_buyer"],
            "price_test_eur": product["price_test"],
            "status": "design_pending_human_review",
            "design_language": "en",
            "localization_plan": "Create Spanish version preserving the master grid.",
            "source_of_truth": "app_design_spec.md + master_workbook.xlsx + model_contract.json",
            "global_design_system": "rules/DESIGN_SYSTEM.md",
            "native_excel_design_contract": {
                "schema_version": "1.0.0",
                "source_of_truth": "master_workbook.xlsx",
                "opening_zoom": 100,
                "grid": {"units": {"column_width": "excel_character_width", "row_height": "points", "shape_position": "points"}, "table_anchor_column": "B", "title_column": "B", "table_outer_padding_columns": ["A", "J"]},
                "header": {"eyebrow_row": 2, "title_row": 3, "subtitle_row": 4, "table_header_row": 7},
                # Master-derived products inherit the master's top navigation
                # geometry. New cell-native products use the shared lower
                # navigation baseline so it cannot overwrite the header row.
                "navigation": {
                    "row": 1 if product.get("preserve_native_master_objects") else 32,
                    "height_points": 26 if product.get("preserve_native_master_objects") else 30,
                    "gap_points": 4,
                    "component": "master_in_sheet_shapes" if product.get("preserve_native_master_objects") else "native_cell_navigation",
                },
            },
            "visual_identity": product["visual_identity"],
            "signature_chart": product["signature_chart"],
            "palette": product["palette"],
            "visual_inheritance": {"pattern_id": "dashboard-card-v1", "manual": "rules/FACTORY_SCREEN_PATTERN_MANUAL.md", "machine_contract": "FACTORY_SCREEN_PATTERN.json", "role": "simple_excel_inheriting_screen_pattern"},
            "tabs": product["tabs"],
            "functional_scope": product["features"],
            "mock_profile": product.get("mock_profile", {}),
            "visual_acceptance": {"human_review_required": True, "native_geometry_required": True, "source_model_match_required": True},
            "research_sources": product["sources"],
            "source_lineage": product["sources"],
            "model_base": product.get("model_base", "templates/excel-first-base"),
            "native_reference_mode": "master_workbook_contract",
            "master_workbook": "templates/excel-first-base/master_workbook.xlsx",
            "template_id": "simple-excel-workbook-v1",
            "template_manifest": "templates/excel-first-base/template_manifest.json",
            "native_sheet_map": product.get("native_sheet_map"),
            "preserve_native_master_objects": bool(product.get("preserve_native_master_objects")),
            "skip_native_overlay": bool(product.get("skip_native_overlay")),
            "native_surface_adaptations": product.get("native_surface_adaptations", {}),
            "model_behavior_file": "model_behavior.md",
            "model_contract_file": "model_contract.json",
            "app_user_guide_file": "app/README.md",
            "model_behavior_source": "declared_model_contract_before_xlsx",
            "model_contract": model_contract,
            "user_model_guide": {"file": "model_behavior.md", "purpose": "Guía de uso y verificación funcional independiente del XLSX.", "edit_inputs_only": True, "fail_on_contract_mismatch": True},
        }
        apply_design_architecture(spec)
        (folder / "design_spec.json").write_text(json.dumps(spec, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        review = {"status": "pending_human_review", "reviewed_by": None, "reviewed_at": None, "notes": "Pendiente de revisión humana del XLSX final; no se requiere evidencia de imagen."}
        (folder / "design_review.json").write_text(json.dumps(review, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"product-{product['id']}: master_workbook.xlsx contract")


if __name__ == "__main__":
    main()
