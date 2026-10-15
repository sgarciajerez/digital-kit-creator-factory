from __future__ import annotations

import json
import argparse
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook

from excel_visual_overlays import apply_simple_native_overlays, apply_visual_overlays
from validate_design_architecture import validate_product
from factory_config import active_product_ids, all_product_ids, finalized_product_ids
from excel_template_factory import ExcelTemplateFactory
from model_quality import validate_model_file
from model_contracts import contract_for_product, write_app_usage_markdown, write_model_behavior_markdown, write_model_contract_json
from artifact_integrity import artifact_identity


ROOT = Path(__file__).resolve().parent.parent
OUT_ROOT = ROOT / "outputs" / "listings"
ALL_PRODUCT_IDS = all_product_ids()
PRODUCT_IDS = active_product_ids()


# Historical renderers live in legacy_workbook_builders.py, outside production.


from native_project_builder import master_workbook_builder


def functional_contract_for_product(product_id: int) -> dict:
    """Backward-compatible view of the authoritative model contract."""
    return {"cases": contract_for_product(product_id).get("cases", [])}


def validate_generated_workbook(path: Path, expected_tabs: list[str], spec: dict | None = None) -> dict:
    checks: list[str] = []
    errors: list[str] = []
    if not path.exists():
        return {"passed": False, "score": 0, "checks": [], "errors": ["XLSX no existe"]}
    guide_path = path.parent / "README.md"
    guide_pdf_name = str((spec or {}).get("app_user_guide_pdf_filename", f"{(spec or {}).get('slug', 'excel')}-user-guide.pdf"))
    guide_pdf_path = path.parent / guide_pdf_name
    expected_language = str((spec or {}).get("design_language") or "en").casefold()
    if not guide_path.is_file():
        errors.append("Falta app/README.md con instrucciones de uso")
    elif f"**Language:** `{expected_language}`" not in guide_path.read_text(encoding="utf-8"):
        errors.append("app/README.md no coincide con el idioma declarado del producto")
    if not guide_pdf_path.is_file():
        errors.append("Falta la guía de uso equivalente en PDF")
    try:
        wb = load_workbook(path, data_only=False)
    except Exception as exc:
        return {"passed": False, "score": 0, "checks": [], "errors": [f"No se puede abrir: {exc}"]}
    if set(expected_tabs).issubset(set(wb.sheetnames)):
        checks.append("Todas las pestañas del diseño existen")
    else:
        errors.append("Faltan pestañas definidas en el diseño")
    formulas = [cell for ws in wb.worksheets for row in ws.iter_rows() for cell in row if isinstance(cell.value, str) and cell.value.startswith("=")]
    if len(formulas) >= 12:
        checks.append("El archivo contiene fórmulas conectadas")
    else:
        errors.append("El archivo contiene pocas fórmulas")
    validations = sum(len(ws.data_validations.dataValidation) for ws in wb.worksheets)
    if validations >= 1:
        checks.append("El archivo contiene validaciones desplegables")
    else:
        errors.append("Faltan validaciones de entrada")
    charts = sum(len(ws._charts) for ws in wb.worksheets)
    cell_native_visuals = any(ws.title == "Dashboard" for ws in wb.worksheets)
    simple_contract = spec and spec.get("design_approach") == "excel_first_simple_table_system"
    declared_charts = sum(
        1
        for contract in (spec or {}).get("screen_contracts", {}).values()
        if isinstance(contract, dict) and "chart" in (contract.get("components") or [])
    )
    preserved_master_charts = None
    if spec and spec.get("preserve_native_master_objects"):
        try:
            master_path = ROOT / str(spec.get("master_workbook", "templates/excel-first-base/master_workbook.xlsx"))
            master = load_workbook(master_path, data_only=False, read_only=False)
            mapping = spec.get("native_sheet_map") or {}
            preserved_master_charts = sum(len(master[source]._charts) for source in mapping.values() if source in master.sheetnames)
            master.close()
        except Exception:
            preserved_master_charts = None
    if (simple_contract and ((preserved_master_charts is not None and charts == preserved_master_charts) or (preserved_master_charts is None and charts == declared_charts))) or (not simple_contract and (charts >= 1 or cell_native_visuals)):
        checks.append("El archivo contiene gráficos de Excel")
    else:
        errors.append("Falta al menos un gráfico")
    title_row = int(((spec or {}).get("native_excel_design_contract", {}).get("header", {}) or {}).get("title_row", 1))
    title_column = str(((spec or {}).get("native_excel_design_contract", {}).get("grid", {}) or {}).get("title_column", "A"))
    for sheet in expected_tabs:
        ws = wb[sheet]
        # The shared visual contract keeps the editorial title in A3 on the
        # two wide landing screens and in B3 on table screens. The old check
        # applied one column globally, so valid Dashboard/Start here screens
        # were incorrectly marked as missing their title.
        sheet_title_column = title_column if spec and spec.get("preserve_native_master_objects") else ("A" if sheet in {"Start here", "Dashboard"} else title_column)
        title_cell = f"{sheet_title_column}{title_row}"
        if ws[title_cell].value:
            checks.append(f"{sheet}: título presente en {title_cell}")
        else:
            errors.append(f"{sheet}: falta título en {title_cell}")
    wb.close()
    model_report = validate_model_file(
        path,
        expected_sheets=expected_tabs,
        functional_contract=(spec or {}).get("functional_contract"),
        run_functional=bool((spec or {}).get("run_functional_qa", False)),
        model_contract=(spec or {}).get("model_contract"),
        model_contract_path=(
            Path(path).parent.parent / str((spec or {}).get("model_behavior_file", "model_behavior.md"))
            if (spec or {}).get("model_contract")
            else None
        ),
    )
    checks.extend(model_report.get("checks", []))
    errors.extend(model_report.get("errors", []))
    score = round(100 * len(checks) / max(1, len(checks) + len(errors)), 2)
    return {
        "passed": not errors,
        "score": score,
        "checks": checks,
        "errors": errors,
        "formula_count": len(formulas),
        "validation_count": validations,
        "chart_count": charts,
        "model_quality": model_report,
    }


def style_product_charts(path: Path, spec: dict) -> None:
    """Apply the product palette and English chart typography globally.

    Charts are native objects inherited from the master, so changing cells or
    table colours alone does not update their series, titles or axis labels.
    This final COM pass makes every chart obey the same product-level visual
    contract.
    """
    import pythoncom
    import win32com.client as win32

    def excel_rgb(value: str) -> int:
        value = str(value).replace("#", "")
        red, green, blue = (int(value[index:index + 2], 16) for index in (0, 2, 4))
        return red + (green << 8) + (blue << 16)

    palette = spec["palette"]
    accent = excel_rgb(palette["accent"])
    secondary = excel_rgb(palette["secondary"])
    background = excel_rgb(palette["background"])
    white = 0xFFFFFF
    black = 0x000000

    def style_text_shape(shape) -> None:
        """Give chart title boxes a soft palette gradient and plain black text."""
        try:
            shape.Format.Fill.Visible = -1
            shape.Format.Fill.TwoColorGradient(1, 1)
            shape.Format.Fill.ForeColor.RGB = background
            shape.Format.Fill.BackColor.RGB = white
            shape.Format.Line.Visible = False
        except Exception:
            pass
        try:
            text_range = shape.Format.TextFrame2.TextRange
            text_range.Font.Fill.ForeColor.RGB = black
            text_range.Font.Line.Visible = False
        except Exception:
            pass

    chart_titles = {
        "Dashboard": "Planned vs actual spend",
        "Timeline": "Days left by task",
    }
    axis_titles = {
        "Dashboard": ("Category", "Amount"),
        "Timeline": ("Task", "Days left"),
    }
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        for sheet in workbook.Worksheets:
            if sheet.ChartObjects().Count < 1:
                continue
            for index in range(1, sheet.ChartObjects().Count + 1):
                chart = sheet.ChartObjects(index).Chart
                sheet_name = str(sheet.Name)
                chart.HasTitle = True
                chart.ChartTitle.Text = chart_titles.get(sheet_name, f"{sheet_name} overview")
                for part in (chart.ChartArea, chart.PlotArea):
                    try:
                        part.Format.Fill.Visible = -1
                        part.Format.Fill.ForeColor.RGB = background
                        part.Format.Line.Visible = False
                    except Exception:
                        pass
                for series_index in range(1, chart.SeriesCollection().Count + 1):
                    series = chart.SeriesCollection(series_index)
                    colour = accent if series_index % 2 else secondary
                    try:
                        series.Format.Fill.ForeColor.RGB = colour
                        series.Format.Line.ForeColor.RGB = colour
                    except Exception:
                        pass
                category_title, value_title = axis_titles.get(sheet_name, ("Category", "Amount"))
                for axis_type, title in ((1, category_title), (2, value_title)):
                    try:
                        axis = chart.Axes(axis_type)
                        axis.HasTitle = True
                        axis.AxisTitle.Text = title
                        style_text_shape(axis.AxisTitle)
                        axis.TickLabels.Font.Color = black
                    except Exception:
                        pass
                try:
                    style_text_shape(chart.ChartTitle)
                except Exception:
                    pass
                try:
                    legend_text = chart.Legend.Format.TextFrame2.TextRange
                    legend_text.Font.Fill.ForeColor.RGB = black
                    legend_text.Font.Line.Visible = False
                except Exception:
                    pass
                try:
                    for series_index in range(1, chart.SeriesCollection().Count + 1):
                        series = chart.SeriesCollection(series_index)
                        if series.HasDataLabels:
                            labels = series.DataLabels().Format.TextFrame2.TextRange
                            labels.Font.Fill.ForeColor.RGB = black
                            labels.Font.Line.Visible = False
                except Exception:
                    pass
        workbook.Save()
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def apply_declared_charts(path: Path, spec: dict) -> None:
    """Compile explicitly declared chart slots after the product builder runs."""
    contracts = spec.get("screen_contracts") or {}
    if not any("chart" in (contract.get("components") or []) for contract in contracts.values() if isinstance(contract, dict)):
        return
    if spec.get("preserve_native_master_objects"):
        from workbook_views import apply_native_charts
        apply_native_charts(path, spec)
        return
    wb = load_workbook(path)
    factory = ExcelTemplateFactory(wb, spec)
    for tab, contract in contracts.items():
        if not isinstance(contract, dict) or "chart" not in (contract.get("components") or []):
            continue
        factory.write_chart(wb[tab], dict(contract["chart"]))
    wb.save(path)
    wb.close()


def repair_native_date_formats(path: Path) -> None:
    """Restore locale-safe date formats after any openpyxl overlay save.

    openpyxl can rewrite Excel's locale-aware COM format when it re-saves a
    workbook produced by the native master builder.  On Spanish Excel that
    can surface the year token as the literal text ``yyyy`` or fall back to
    the machine's ``mm-dd-yy`` format.  A final COM pass is therefore part of
    the factory contract, not a product-specific workaround.
    """
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        date_format = "dd/mm/aaaa"
        ranges_by_sheet = {
            "Start here": ("H10:H11",),
            "Settings": ("C11:C12",),
            "Purchases": ("B8:B105", "M8:M105"),
            "Quotes": ("D8:D105",),
            "Timeline": ("D8:E105",),
        }
        for sheet_name, ranges in ranges_by_sheet.items():
            try:
                sheet = workbook.Worksheets(sheet_name)
            except Exception:
                continue
            for address in ranges:
                sheet.Range(address).NumberFormatLocal = date_format
        workbook.Save()
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main() -> int:
    parser = argparse.ArgumentParser(description="Genera y valida las apps Excel de la fábrica")
    parser.add_argument("--product", type=int, help="Generar un producto con design_spec.json existente, aunque no esté activo en el lote")
    parser.add_argument("--no-open", action="store_true", help="Compatibilidad: la generación no abre ventanas de usuario")
    args = parser.parse_args()
    results = []
    product_ids = (args.product,) if args.product else PRODUCT_IDS
    for product_id in product_ids:
        folder = OUT_ROOT / f"product-{product_id}"
        spec = json.loads((folder / "design_spec.json").read_text(encoding="utf-8"))
        from digital_kit import build_digital_kit, product_kit_policy
        kit_policy = product_kit_policy(spec.get("slug", ""))
        if not kit_policy:
            results.append({
                "product_id": product_id,
                "slug": spec.get("slug"),
                "status": "blocked_by_missing_kit_profile",
            })
            print(f"product-{product_id}: blocked_by_missing_kit_profile")
            continue
        if kit_policy.get("product_readiness", "ready") != "ready":
            manifest_path = folder / "draft_package" / "listing.json"
            if manifest_path.is_file():
                previous_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                previous_manifest["status"] = "blocked_by_product_redesign"
                previous_manifest["next_action"] = kit_policy.get("readiness_note", "Redesign the product before sale.")
                manifest_path.write_text(json.dumps(previous_manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            result = {
                "product_id": product_id,
                "slug": spec.get("slug"),
                "status": "blocked_by_product_redesign",
                "readiness_note": kit_policy.get("readiness_note"),
            }
            results.append(result)
            print(f"product-{product_id}: blocked_by_product_redesign")
            continue
        manifest_path = folder / "draft_package" / "listing.json"
        if manifest_path.is_file():
            previous_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            previous_manifest["status"] = "requires_repackaging"
            previous_manifest["next_action"] = "Regenerate the package after validating the new workbook."
            manifest_path.write_text(json.dumps(previous_manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        # The declarative contract module is the normative source.  Do not
        # reuse a stale embedded copy from an older design spec, otherwise the
        # generated workbook and its model_behavior.md can describe different
        # products after a contract correction.
        model_contract = contract_for_product(product_id)
        spec["model_contract"] = model_contract
        spec["model_behavior_file"] = str(spec.get("model_behavior_file", "model_behavior.md"))
        behavior_path = folder / spec["model_behavior_file"]
        write_model_behavior_markdown(behavior_path, model_contract)
        write_model_contract_json(folder / str(spec.get("model_contract_file", "model_contract.json")), model_contract)
        spec["functional_contract"] = {"cases": model_contract.get("cases", [])}
        spec["run_functional_qa"] = True
        if product_id in finalized_product_ids():
            architecture_ok = True
            architecture_report = {"status": "skipped_finalized_product", "errors": []}
        else:
            architecture_ok, architecture_report = validate_product(product_id)
            if not architecture_ok:
                print(f"product-{product_id}: app=blocked_by_design_architecture")
                for error in architecture_report["errors"]:
                    print(f"  - {error}")
                return 1
        app_dir = folder / "app"
        app_dir.mkdir(parents=True, exist_ok=True)
        app_guide_path = app_dir / "README.md"
        write_app_usage_markdown(app_guide_path, spec, model_contract)
        guide_filename = str(spec.get("app_user_guide_pdf_filename", f"{spec['slug']}-user-guide.pdf"))
        if not guide_filename.lower().endswith(".pdf") or Path(guide_filename).name != guide_filename:
            raise RuntimeError(f"product-{product_id}: app_user_guide_pdf_filename debe ser un nombre PDF sin rutas")
        from product_guide_pdf import generate_user_guide_pdf
        user_guide_pdf = generate_user_guide_pdf(app_guide_path, app_dir / guide_filename, spec)
        user_guide_pdf["file"] = str((app_dir / guide_filename).relative_to(ROOT))
        path = app_dir / f"{spec['slug']}-excel-template.xlsx"
        master_workbook = ROOT / str(spec.get("master_workbook", "templates/excel-first-base/master_workbook.xlsx"))
        spec["_native_source_present"] = master_workbook.is_file() and bool(spec.get("model_base")) and bool(spec.get("source_lineage"))
        # There is one factory renderer: every product starts from the master
        # workbook and is adapted through its declared source-sheet map.
        # Product-specific legacy builders are intentionally outside the
        # production path.
        builder = master_workbook_builder
        if not spec.get("preserve_native_master_objects"):
            raise RuntimeError(f"product-{product_id}: la factoría solo admite productos derivados del master")
        if not isinstance(spec.get("native_sheet_map"), dict) or list(spec["native_sheet_map"]) != list(spec["tabs"]):
            raise RuntimeError(f"product-{product_id}: native_sheet_map completo y ordenado obligatorio")
        if spec.get("skip_native_overlay"):
            raise RuntimeError(f"product-{product_id}: skip_native_overlay no está permitido en productos derivados del master")
        if spec.get("native_geometry_policy") != "exact_master_grid":
            raise RuntimeError(f"product-{product_id}: falta la politica exact_master_grid")
        builder(spec, path)
        apply_declared_charts(path, spec)
        if spec.get("design_approach") == "excel_first_simple_table_system" and not spec.get("skip_native_overlay"):
            apply_simple_native_overlays(path, spec)
        elif not spec.get("skip_native_overlay"):
            apply_visual_overlays(path, spec)
        if spec.get("preserve_native_master_objects"):
            repair_native_date_formats(path)
        style_product_charts(path, spec)
        from workbook_views import apply_workbook_views
        apply_workbook_views(path, spec)
        ebook_spec = spec.get("ebook") or {}
        ebook_filename = str(ebook_spec.get("filename", ""))
        ebook_required = bool(kit_policy.get("ebook_required", False))
        if ebook_filename:
            if not ebook_filename.lower().endswith(".pdf") or Path(ebook_filename).name != ebook_filename:
                raise RuntimeError(f"product-{product_id}: ebook.filename debe ser un nombre PDF descriptivo sin rutas")
            from product_ebook import generate_product_ebook
            ebook_path = app_dir / ebook_filename
            ebook_report = generate_product_ebook(spec, ebook_path)
            ebook_report["file"] = str(ebook_path.relative_to(ROOT))
            ebook_report["included"] = True
        elif ebook_required:
            raise RuntimeError(f"product-{product_id}: el perfil del kit exige un e-book PDF declarado")
        else:
            ebook_report = {"passed": True, "included": False, "reason": "No topic e-book is required by this kit profile."}
        validation = validate_generated_workbook(path, spec["tabs"], spec)
        # Keep the normative model guide and the independent result of its
        # verification as first-class factory artifacts.  The guide is built
        # from the declared design contract before the XLSX exists; this QA
        # report records whether the generated workbook satisfies it.
        model_qa_path = folder / "model_behavior_qa.json"
        model_qa_path.write_text(json.dumps(validation.get("model_quality", {}), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        report = {"product_id": product_id, "slug": spec["slug"], "generated_at": datetime.now(timezone.utc).isoformat(), "file": str(path.relative_to(ROOT)), "user_guide_pdf": user_guide_pdf, "ebook": ebook_report, "validation": validation, "design_sources": {"master_workbook": str(master_workbook.relative_to(ROOT)) if master_workbook.is_relative_to(ROOT) else str(master_workbook), "native_model": spec["_native_source_present"], "source_lineage": spec.get("source_lineage", [])}, "visual_pattern": spec.get("visual_inheritance", {}), "design_architecture": architecture_report, "status": "generated" if validation["passed"] and ebook_report["passed"] and user_guide_pdf["passed"] else "needs_implementation_fix"}
        report["artifact_identity"] = artifact_identity(path, master_workbook, model_contract, spec)
        if validation["passed"]:
            from customer_workbook import create_clean_workbook
            from artifact_integrity import sha256
            clean_path = path.with_name(f"{spec['slug']}-blank.xlsx")
            clean_qa = create_clean_workbook(path, clean_path, model_contract)
            report["customer_copy"] = {"file": str(clean_path.relative_to(ROOT)), "sha256": sha256(clean_path), "validation": clean_qa}
            if not clean_qa["passed"]:
                validation["passed"] = False
                report["status"] = "needs_implementation_fix"
        if validation.get("passed") and report.get("customer_copy", {}).get("validation", {}).get("passed"):
            kit_report = build_digital_kit(spec, app_dir, path, ROOT / report["customer_copy"]["file"], user_guide_pdf, ebook_report)
            if kit_report.get("file"):
                kit_report["file"] = str(Path(kit_report["file"]).relative_to(ROOT))
        else:
            kit_report = {"passed": False, "status": "blocked_by_workbook_qa", "file": None, "included_files": []}
        report["digital_kit"] = kit_report
        if not kit_report.get("passed"):
            report["status"] = "needs_implementation_fix"
        (folder / "app_generation_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        results.append(report)
        print(f"product-{product_id}: app={report['status']} score={validation['score']}")
    batch = {"generated_at": datetime.now(timezone.utc).isoformat(), "products": results, "all_passed": all(item.get("validation", {}).get("passed", False) for item in results)}
    if not args.product:
        (OUT_ROOT / "app_generation_report.json").write_text(json.dumps(batch, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"APP_GENERATION={'PASSED' if batch['all_passed'] else 'FAILED'}")
    return 0 if batch["all_passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
