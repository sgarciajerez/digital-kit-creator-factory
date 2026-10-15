"""Select, render, validate and package product-specific digital kits."""
from __future__ import annotations

import hashlib
import json
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import BaseDocTemplate, Frame, PageBreak, PageTemplate, Paragraph, Spacer

from product_guide_pdf import generate_user_guide_pdf
from html_kit_tool import generate_html_tool, select_html_tool


ROOT = Path(__file__).resolve().parent.parent
CATALOG_PATH = ROOT / "config" / "digital_kit_catalog.json"


def product_kit_policy(slug: str) -> dict:
    catalog = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    return dict(catalog.get("products", {}).get(str(slug), {}))


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _project_path(value: str | Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else ROOT / path


def _safe_hex(value: str, fallback: str) -> colors.Color:
    raw = str(value or fallback).strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", raw):
        raw = fallback.lstrip("#")
    return colors.HexColor(f"#{raw}")


def _render_worksheet_bundle(content_path: Path, output_path: Path, spec: dict) -> dict:
    content = json.loads(content_path.read_text(encoding="utf-8"))
    pages = content.get("pages") or []
    if not pages:
        raise ValueError(f"Companion module has no worksheet pages: {content_path}")
    palette = spec.get("palette") or {}
    pale = _safe_hex(palette.get("background"), "F2F5EC")
    accent = _safe_hex(palette.get("accent"), "5C8058")
    secondary = _safe_hex(palette.get("secondary"), "D17A5A")
    width, height = A4
    margin = 48
    page_numbers: list[int] = []

    def decorate(canvas, doc) -> None:
        page_numbers.append(doc.page)
        canvas.saveState()
        canvas.setStrokeColor(pale)
        canvas.setLineWidth(4)
        canvas.line(margin, height - 31, width - margin, height - 31)
        canvas.setStrokeColor(accent)
        canvas.setLineWidth(2)
        canvas.line(margin, height - 34, margin + 48, height - 34)
        canvas.setFillColor(colors.black)
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(width - margin, 24, str(doc.page))
        canvas.restoreState()

    doc = BaseDocTemplate(
        str(output_path), pagesize=A4, leftMargin=margin, rightMargin=margin,
        topMargin=52, bottomMargin=42, title=content["title"],
        author=str(spec.get("title", "Digital kit")), subject=content.get("subtitle", "Printable worksheets"),
    )
    frame = Frame(margin, 42, width - 2 * margin, height - 94, id="worksheet", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="Worksheet", frames=frame, onPage=decorate)])
    title_style = ParagraphStyle("WorksheetTitle", fontName="Helvetica-Bold", fontSize=16, leading=20, textColor=colors.black, backColor=pale, borderPadding=(8, 8, 8, 8), spaceAfter=13)
    intro_style = ParagraphStyle("WorksheetIntro", fontName="Helvetica", fontSize=12, leading=16, textColor=colors.black, spaceAfter=9)
    label_style = ParagraphStyle("WorksheetLabel", fontName="Helvetica-Bold", fontSize=12, leading=16, textColor=colors.black, spaceBefore=7, spaceAfter=4)
    check_style = ParagraphStyle("WorksheetCheck", fontName="Helvetica", fontSize=12, leading=17, textColor=colors.black, leftIndent=8, spaceAfter=5)
    field_style = ParagraphStyle("WorksheetField", fontName="Helvetica", fontSize=12, leading=17, textColor=colors.black, spaceAfter=6)

    story = []
    for index, page in enumerate(pages):
        if index:
            story.append(PageBreak())
        story.append(Paragraph(escape(str(page.get("title", f"Worksheet {index + 1}"))), title_style))
        if page.get("intro"):
            story.append(Paragraph(escape(str(page["intro"])), intro_style))
        if page.get("checks"):
            story.append(Paragraph("CHECKLIST", label_style))
            story.extend(Paragraph(f"[  ]  {escape(str(item))}", check_style) for item in page["checks"])
        if page.get("fields"):
            story.append(Paragraph("NOTES AND NEXT ACTIONS", label_style))
            for field in page["fields"]:
                story.append(Paragraph(f"<b>{escape(str(field))}</b>", field_style))
                story.append(Paragraph("________________________________________________________________________________", field_style))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.build(story)
    return {
        "passed": bool(page_numbers) and len(page_numbers) == len(pages),
        "file": str(output_path),
        "sha256": _sha(output_path),
        "content_file": str(content_path),
        "content_sha256": _sha(content_path),
        "page_count": max(page_numbers, default=0),
    }


def _kit_index(spec: dict, components: list[dict], selected_modules: list[dict]) -> str:
    lines = [
        f"# Digital kit contents - {spec.get('title', 'Excel product')}",
        "",
        "Start with the PDF user guide. Use the example workbook to learn the flow, then enter your own information in the blank workbook.",
        "",
        "## Included files",
        "",
    ]
    for item in components:
        description = item.get("description", "Product file")
        lines.append(f"- `{item['name']}` — {description}")
    if selected_modules:
        lines += ["", "## Why these companion worksheets are included", ""]
        lines.extend(f"- {module['title']}: {module['reason']}" for module in selected_modules)
    html_component = next((item for item in components if Path(item["name"]).suffix.casefold() == ".html"), None)
    if html_component:
        spanish = str(spec.get("design_language") or spec.get("language") or "en").casefold().startswith("es")
        lines += [
            "",
            ("Abre el archivo HTML en un navegador actualizado. Funciona sin internet y no guarda ni envía tus datos." if spanish else "Open the HTML file in a recent browser. It works offline and does not save or send your data."),
        ]
    lines += [
        "",
        "## Suggested order",
        "",
        "1. Read the PDF user guide and open the example workbook.",
        "2. Make a fresh copy of the blank workbook for your own project.",
        "3. Use the selected companion worksheets alongside the matching Excel workflow.",
        "4. Keep the original example unchanged as a reference.",
        "",
        "The Markdown guide is included as an accessible source format; the PDF is the recommended reading copy.",
        "",
    ]
    return "\n".join(lines)


def build_digital_kit(
    spec: dict,
    app_dir: Path,
    populated_workbook: Path,
    blank_workbook: Path,
    user_guide: dict,
    ebook: dict | None = None,
) -> dict:
    """Choose companion modules from product needs and produce a hash-checked ZIP."""
    catalog = json.loads(CATALOG_PATH.read_text(encoding="utf-8"))
    product = catalog.get("products", {}).get(str(spec.get("slug", "")), {})
    if not product:
        return {
            "passed": False,
            "status": "blocked_by_missing_kit_profile",
            "block_reason": "Declare the product's kit needs, required modules and e-book policy before generation.",
            "profile": None,
            "needs": [],
            "selected_modules": [],
            "file": None,
            "included_files": [],
        }
    needs = set(product.get("needs", []))
    covered_needs = {need for module in catalog.get("modules", []) for need in module.get("triggers", [])}
    uncovered_needs = sorted(needs - covered_needs)
    if uncovered_needs and product.get("product_readiness", "ready") == "ready":
        raise ValueError("No digital-kit module covers declared product needs: " + ", ".join(uncovered_needs))
    selected_modules = []
    module_assets = []
    for module in catalog.get("modules", []):
        matched = sorted(needs.intersection(module.get("triggers", [])))
        if not matched:
            continue
        selected_modules.append({
            "id": module["id"],
            "title": module["title"],
            "matched_needs": matched,
            "reason": "Selected because product needs match: " + ", ".join(matched),
            "asset": module["filename"],
        })

    if product.get("product_readiness", "ready") != "ready":
        return {
            "passed": False,
            "status": product.get("product_readiness", "blocked_by_product_readiness"),
            "block_reason": product.get("readiness_note", "Product readiness has not been approved."),
            "profile": product.get("profile"),
            "needs": sorted(needs),
            "uncovered_needs": uncovered_needs,
            "selected_modules": selected_modules,
            "file": None,
            "included_files": [],
        }

    for module in catalog.get("modules", []):
        selected = next((item for item in selected_modules if item["id"] == module["id"]), None)
        if not selected:
            continue
        content_path = _project_path(module["content_file"])
        output_path = app_dir / str(module["filename"])
        generated = _render_worksheet_bundle(content_path, output_path, spec)
        generated["file"] = str(output_path)
        module_assets.append({**generated, "name": output_path.name, "description": module["title"]})

    selected_html_tool = select_html_tool(product, catalog.get("html_tools", []))
    html_tool_report = None
    html_tool_asset = None
    if selected_html_tool:
        html_tool_path = app_dir / selected_html_tool["filename"]
        html_tool_report = generate_html_tool(spec, html_tool_path, selected_html_tool)
        html_tool_report["id"] = selected_html_tool["id"]
        html_tool_report["file"] = html_tool_path.name
        html_tool_report["sha256"] = _sha(html_tool_path)
        generator_path = Path(__file__).with_name("html_kit_tool.py")
        html_tool_report["generator_file"] = str(generator_path.relative_to(ROOT))
        html_tool_report["generator_sha256"] = _sha(generator_path)
        html_tool_report["matched_needs"] = selected_html_tool["matched_needs"]
        html_tool_report["selection_reason"] = selected_html_tool["reason"]
        html_tool_asset = {
            "file": str(html_tool_path),
            "description": html_tool_report["title"],
        }

    guide_pdf = _project_path(user_guide.get("file", "missing-user-guide.pdf"))
    guide_markdown = app_dir / "README.md"
    required = [
        (populated_workbook, "Pre-filled example workbook"),
        (blank_workbook, "Blank workbook for the buyer's own data"),
        (guide_markdown, "Markdown user guide source"),
        (guide_pdf, "Readable PDF user guide"),
        *[(Path(item["file"]), item["description"]) for item in module_assets],
    ]
    if html_tool_asset:
        required.append((Path(html_tool_asset["file"]), html_tool_asset["description"]))
    ebook_required = bool(product.get("ebook_required", False))
    ebook_path: Path | None = None
    if ebook and ebook.get("file"):
        ebook_path = _project_path(ebook["file"])
        if ebook_required and not ebook.get("passed"):
            raise ValueError("The digital kit requires a valid topic e-book PDF.")
        if ebook.get("passed"):
            required.append((ebook_path, "Topic e-book PDF"))
    elif ebook_required:
        raise ValueError("This product's digital kit requires an e-book, but none was generated.")

    missing = [str(path) for path, _ in required if not path.is_file()]
    if missing:
        raise FileNotFoundError("Digital kit files are missing: " + ", ".join(missing))

    component_records = [{"name": path.name, "description": description, "sha256": _sha(path), "size_bytes": path.stat().st_size} for path, description in required]
    index_text = _kit_index(spec, component_records, selected_modules)
    index_markdown = app_dir / "KIT-INDEX.md"
    index_markdown.write_text(index_text, encoding="utf-8")
    index_pdf = app_dir / "KIT-INDEX.pdf"
    index_pdf_report = generate_user_guide_pdf(index_markdown, index_pdf, spec)
    component_records.extend([
        {"name": index_markdown.name, "description": "Package contents and recommended order", "sha256": _sha(index_markdown), "size_bytes": index_markdown.stat().st_size},
        {"name": index_pdf.name, "description": "PDF package contents and recommended order", "sha256": _sha(index_pdf), "size_bytes": index_pdf.stat().st_size},
    ])
    manifest_path = app_dir / "digital-kit-manifest.json"
    manifest = {
        "schema_version": catalog["schema_version"],
        "catalog_file": str(CATALOG_PATH.relative_to(ROOT)),
        "catalog_sha256": _sha(CATALOG_PATH),
        "product_id": spec.get("product_id"),
        "slug": spec.get("slug"),
        "title": spec.get("title"),
        "profile": product.get("profile", "general"),
        "needs": sorted(needs),
        "uncovered_needs": uncovered_needs,
        "selection_policy": catalog["selection_policy"],
        "selected_modules": selected_modules,
        "interactive_html_tool": html_tool_report,
        "module_sources": [
            {"file": str(Path(item["content_file"]).relative_to(ROOT)), "sha256": item["content_sha256"]}
            for item in module_assets
        ],
        "files": component_records,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    archive = app_dir / f"{spec.get('slug', 'product')}-digital-kit.zip"
    archive_inputs = [path for path, _ in required] + [index_markdown, index_pdf, manifest_path]
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=8) as package:
        for path in archive_inputs:
            package.write(path, arcname=path.name)
    with zipfile.ZipFile(archive, "r") as package:
        bad_member = package.testzip()
        if bad_member:
            archive.unlink(missing_ok=True)
            raise ValueError(f"Digital kit ZIP failed integrity check at {bad_member}")
        packaged_names = package.namelist()
    return {
        "passed": True,
        "status": "ready",
        "file": str(archive),
        "sha256": _sha(archive),
        "size_bytes": archive.stat().st_size,
        "profile": product.get("profile", "general"),
        "needs": sorted(needs),
        "selected_modules": selected_modules,
        "included_files": packaged_names,
        "file_count": len(packaged_names),
        "ebook_included": ebook_path is not None and ebook_path.is_file(),
        "interactive_html_tool": html_tool_report,
        "kit_index_pdf": index_pdf_report,
        "manifest_file": str(manifest_path),
        "manifest_sha256": _sha(manifest_path),
        "catalog_file": str(CATALOG_PATH.relative_to(ROOT)),
        "catalog_sha256": _sha(CATALOG_PATH),
        "module_sources": manifest["module_sources"],
    }
