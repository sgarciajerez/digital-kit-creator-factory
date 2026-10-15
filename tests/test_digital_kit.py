import hashlib
import json
import zipfile
from pathlib import Path

from digital_kit import build_digital_kit
from html_kit_tool import generate_html_tool, select_html_tool
from product_guide_pdf import generate_user_guide_pdf


def _garden_spec() -> dict:
    return {
        "product_id": 1,
        "slug": "community-garden-planner",
        "title": "Community Garden Budget & Project Planner",
        "palette": {"background": "#F2F5EC", "accent": "#5C8058", "secondary": "#D17A5A"},
    }


def test_digital_kit_selects_garden_companions_and_builds_integrity_checked_zip(tmp_path: Path):
    app = tmp_path / "app"
    app.mkdir()
    workbook = app / "community-garden-planner-excel-template.xlsx"
    blank = app / "community-garden-planner-blank.xlsx"
    ebook = app / "garden-ebook.pdf"
    markdown = app / "README.md"
    guide_pdf = app / "community-garden-planner-user-guide.pdf"
    for path in (workbook, blank, ebook):
        path.write_bytes(b"sample artifact")
    markdown.write_text("# User guide\n\nOpen the workbook.", encoding="utf-8")
    guide_report = generate_user_guide_pdf(markdown, guide_pdf, _garden_spec())
    guide_report["file"] = str(guide_pdf)

    report = build_digital_kit(
        _garden_spec(), app, workbook, blank, guide_report,
        {"passed": True, "file": str(ebook)},
    )

    assert report["passed"] is True
    assert report["profile"] == "community_garden"
    assert report["ebook_included"] is True
    assert report["interactive_html_tool"]["id"] == "garden-bed-planting-calculator"
    assert report["interactive_html_tool"]["offline"] is True
    assert report["selected_modules"][0]["id"] == "community-garden-operations"
    archive = Path(report["file"])
    with zipfile.ZipFile(archive) as package:
        assert package.testzip() is None
        names = set(package.namelist())
        manifest = json.loads(package.read("digital-kit-manifest.json"))
        for item in manifest["files"]:
            assert hashlib.sha256(package.read(item["name"])).hexdigest() == item["sha256"]
        html_name = report["interactive_html_tool"]["file"]
        html_text = package.read(html_name).decode("utf-8")
        index_text = package.read("KIT-INDEX.md").decode("utf-8")
        assert html_name in names
        assert '<html lang="en">' in html_text
        assert "Plants per row" in html_text
        assert "script src=" not in html_text.casefold()
        assert any(item["name"] == html_name for item in manifest["files"])
        assert "works offline" in index_text
        assert manifest["profile"] == "community_garden"
    assert "community-garden-operations-worksheets.pdf" in names
    assert "KIT-INDEX.pdf" in names
    assert "README.md" in names
    assert "digital-kit-manifest.json" in names


def test_creator_product_is_blocked_until_its_model_redesign_is_complete(tmp_path: Path):
    report = build_digital_kit(
        {"product_id": 14, "slug": "creator-content-sponsorship-tracker", "title": "Creator tracker"},
        tmp_path,
        tmp_path / "sample.xlsx",
        tmp_path / "blank.xlsx",
        {},
    )

    assert report["passed"] is False
    assert report["status"] == "blocked_by_product_redesign"
    assert report["file"] is None
    assert {item["id"] for item in report["selected_modules"]} == {"creator-content-and-deals"}


def test_html_tool_selection_uses_best_need_match_and_skips_unmatched_topics(tmp_path: Path):
    candidates = [
        {"id": "generic", "title": "Generic", "filename": "generic.html", "operation": "moving_budget", "triggers": ["budget"], "priority": 1},
        {"id": "best", "title": "Best", "filename": "best.html", "operation": "moving_budget", "triggers": ["budget", "quotes"], "priority": 1},
    ]
    selected = select_html_tool({"needs": ["budget", "quotes"]}, candidates)
    assert selected["id"] == "best"
    assert select_html_tool({"needs": ["wedding_seating"]}, candidates) is None

    output = tmp_path / "calculadora.html"
    report = generate_html_tool(
        {"title": "Calculadora de prueba", "design_language": "es", "palette": {"accent": "#123456"}},
        output,
        {"title": "Calculadora de gastos", "filename": output.name, "operation": "moving_budget"},
    )
    content = output.read_text(encoding="utf-8")
    assert report["offline"] is True
    assert '<html lang="es">' in content
    assert "Reserva sugerida" in content
    assert 'max="100"' in content
    assert "input.value.trim()===''?NaN" in content
    assert "http://" not in content and "https://" not in content


def test_factory_has_a_relevant_html_companion_for_each_ready_kit_profile():
    catalog_path = Path(__file__).resolve().parents[1] / "config" / "digital_kit_catalog.json"
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    expected = {
        "community_garden": "garden-bed-planting-calculator",
        "home_move": "moving-budget-buffer-calculator",
        "home_renovation": "renovation-contingency-calculator",
        "freshwater_aquarium_setup": "freshwater-tank-volume-calculator",
        "puppy_first_year": "puppy-first-year-cost-calculator",
    }
    for product in catalog["products"].values():
        if product.get("product_readiness", "ready") != "ready":
            continue
        tool = select_html_tool(product, catalog["html_tools"])
        assert tool is not None
        assert tool["id"] == expected[product["profile"]]


def test_each_factory_calculator_renders_its_declared_formula(tmp_path: Path):
    catalog_path = Path(__file__).resolve().parents[1] / "config" / "digital_kit_catalog.json"
    tools = json.loads(catalog_path.read_text(encoding="utf-8"))["html_tools"]
    expected_terms = {
        "moving_budget": "estimated costs",
        "renovation_contingency": "Suggested contingency",
        "garden_spacing": "Plants per row",
        "aquarium_volume": "Gross litres",
        "puppy_first_year": "Estimated first-year total",
    }
    for tool in tools:
        output = tmp_path / tool["filename"]
        generate_html_tool({"title": "Sample kit", "design_language": "en"}, output, tool)
        content = output.read_text(encoding="utf-8")
        assert expected_terms[tool["operation"]].casefold() in content.casefold()
        assert "Content-Security-Policy" in content
        assert "script src=" not in content.casefold()


def test_aquarium_helper_is_a_volume_estimate_not_a_dosing_tool(tmp_path: Path):
    tool = {
        "id": "freshwater-tank-volume-calculator",
        "title": "Freshwater Tank Volume Calculator",
        "filename": "tank-volume.html",
        "operation": "aquarium_volume",
    }
    output = tmp_path / tool["filename"]
    report = generate_html_tool({"title": "Freshwater setup", "design_language": "en"}, output, tool)
    content = output.read_text(encoding="utf-8")

    assert report["passed"] is True
    assert report["offline"] is True
    assert "Estimated water volume (litres)" in content
    assert "does not account for substrate" in content
    assert "dosage" not in content.casefold()
    assert "fill>100" in content
