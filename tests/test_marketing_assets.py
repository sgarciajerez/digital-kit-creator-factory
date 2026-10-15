from openpyxl import Workbook

import zipfile

from PIL import Image

from marketing_assets import capture_ebook_cover, inventory_copy, inventory_item_label, marketing_identity, marketing_images, package_deliverables, palette_for
from record_excel_demo import choose_demo_input


def test_marketing_palette_follows_product_spec():
    palette = palette_for({"palette": {"background": "#010203", "accent": "#AABBCC"}})
    assert palette["background"] == "#010203"
    assert palette["accent"] == "#AABBCC"
    assert palette["ink"].startswith("#")


def test_deliverables_inventory_uses_the_workbook_language():
    english = inventory_copy({"design_language": "en"})
    spanish = inventory_copy({"design_language": "es"})
    assert english["heading"] == "EVERYTHING INCLUDED"
    assert english["guide_icon"] == "GUIDE"
    assert english["html_plural"] == "HTML TOOLS"
    assert spanish["html_plural"] == "HERRAMIENTAS HTML"
    assert inventory_item_label({"kind": "html", "name": "tool.html"}, spanish) == "Calculadora interactiva offline (HTML)"
    assert spanish["heading"] == "TODO LO QUE INCLUYE"
    assert spanish["guide_icon"] == "GUÍA"
    assert inventory_item_label({"kind": "workbook", "name": "blank.xlsx"}, spanish) == "Libro Excel en blanco"


def test_ebook_cover_capture_renders_the_actual_first_pdf_page(tmp_path):
    from reportlab.pdfgen import canvas

    pdf_path = tmp_path / "ebook.pdf"
    pdf = canvas.Canvas(str(pdf_path), pagesize=(360, 540))
    pdf.setFillColorRGB(0.1, 0.4, 0.7)
    pdf.rect(0, 0, 360, 540, fill=1, stroke=0)
    pdf.drawString(35, 460, "Actual e-book cover")
    pdf.save()

    output = capture_ebook_cover(tmp_path, {"ebook_file": pdf_path.name})

    assert output == tmp_path / "marketing_assets" / "ebook_cover_first_page.png"
    with Image.open(output) as capture:
        assert capture.size[0] > 360
        assert capture.size[1] > 540


def test_product_specific_marketing_identity_changes_signature_and_art(tmp_path):
    garden_art = tmp_path / "garden.png"
    moving_art = tmp_path / "moving.png"
    Image.new("RGB", (80, 120), "#467A4D").save(garden_art)
    Image.new("RGB", (80, 120), "#3B8790").save(moving_art)
    garden = marketing_identity(
        {"title": "Garden Planner", "visual_identity": {"name": "Shared harvest", "motif": "Leaf beds"}},
        tmp_path,
        {"ebook_cover_image": garden_art.name},
    )
    moving = marketing_identity(
        {"title": "Moving Planner", "visual_identity": {"name": "A clear landing", "motif": "Route and milestones"}},
        tmp_path,
        {"ebook_cover_image": moving_art.name},
    )
    assert garden["signature_sha256"] != moving["signature_sha256"]
    assert garden["artwork_sha256"] != moving["artwork_sha256"]


def test_demo_input_is_selected_from_declared_numeric_input(tmp_path):
    path = tmp_path / "sample.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "Budget"
    sheet["C8"] = 1250
    book.save(path)

    spec = {
        "model_contract": {
            "tables": [{
                "sheet": "Budget",
                "ref": "B7:D9",
                "headers": [
                    {"name": "Area", "role": "derived", "kind": "text"},
                    {"name": "Planned", "role": "input", "kind": "amount"},
                    {"name": "Actual", "role": "derived", "kind": "amount"},
                ],
            }]
        }
    }

    assert choose_demo_input(path, spec) == ("Budget", "C8", 1312.5)


def test_demo_input_never_guesses_undeclared_cells(tmp_path):
    path = tmp_path / "sample.xlsx"
    book = Workbook()
    book.active.title = "Budget"
    book.active["C8"] = 1250
    book.save(path)
    assert choose_demo_input(path, {"model_contract": {"tables": []}}) is None


def test_deliverables_are_read_from_files_and_digital_kit_archive(tmp_path):
    names = ["worked.xlsx", "blank.xlsx", "guide.pdf", "ebook.pdf"]
    for name in names:
        (tmp_path / name).write_bytes(b"present")
    archive = tmp_path / "kit.zip"
    with zipfile.ZipFile(archive, "w") as bundle:
        bundle.writestr("Moving worksheets.pdf", b"resource")
        bundle.writestr("moving-budget-calculator.html", "<html></html>")
    manifest = {
        "workbook_file": names[0], "blank_file": names[1],
        "usage_guide_pdf": names[2], "ebook_file": names[3],
        "digital_kit_zip": archive.name,
    }

    found = package_deliverables(manifest, tmp_path)

    assert [item["name"] for item in found] == [*names, "Moving worksheets.pdf", "moving-budget-calculator.html"]
    assert found[-2]["kind"] == "pdf"
    assert found[-1]["kind"] == "html"


def test_marketing_set_creates_product_cover_and_deliverable_mockups(tmp_path):
    screenshots = []
    for index in range(2):
        path = tmp_path / f"sheet-{index}.png"
        Image.new("RGB", (1920, 1080), "white").save(path)
        screenshots.append(path)
    for name in ("worked.xlsx", "blank.xlsx", "guide.pdf", "ebook.pdf"):
        (tmp_path / name).write_bytes(b"available")
    Image.new("RGB", (500, 700), "#AABBCC").save(tmp_path / "ebook-cover.png")
    manifest = {
        "workbook_file": "worked.xlsx", "blank_file": "blank.xlsx",
        "usage_guide_pdf": "guide.pdf", "ebook_file": "ebook.pdf",
        "ebook_cover_image": "ebook-cover.png",
    }

    generated = marketing_images({"title": "Home Move Planner", "palette": {"accent": "#123456"}}, screenshots, tmp_path, manifest)
    names = {path.name for path in generated}

    assert "00-product-cover.png" in names
    assert generated[1].name == "03-whats-included-mockup.png"
    assert "05-excel-templates-mockup.png" in names
    assert "06-pdf-resources-mockup.png" in names
    with Image.open(tmp_path / "marketing_mockups" / "00-product-cover.png") as cover:
        assert cover.size == (2000, 1500)
        expected_ink = tuple(int(palette_for({})["ink"][index:index + 2], 16) for index in (1, 3, 5))
        assert cover.getpixel((1700, 500)) == expected_ink
    assert manifest.get("marketing_identity") is None  # The pure renderer does not mutate its caller's manifest.


def test_marketing_set_highlights_real_html_calculator_with_its_own_mockup(tmp_path):
    screenshot = tmp_path / "workbook.png"
    calculator = tmp_path / "calculator-preview.png"
    Image.new("RGB", (1920, 1080), "#E6E8E2").save(screenshot)
    Image.new("RGB", (1920, 1080), "#D9EEE8").save(calculator)
    (tmp_path / "worked.xlsx").write_bytes(b"workbook")

    generated = marketing_images(
        {"title": "Aquarium Setup Planner", "design_language": "en"},
        [screenshot],
        tmp_path,
        {
            "workbook_file": "worked.xlsx",
            "interactive_html_tool": {
                "passed": True,
                "title": "Tank Volume Calculator",
                "filename": "tank-volume-calculator.html",
                "language": "en",
            },
        },
        calculator,
    )

    names = {path.name for path in generated}
    assert "00-product-cover.png" in names
    assert "07-interactive-html-calculator-mockup.png" in names
    with Image.open(tmp_path / "marketing_mockups" / "00-product-cover.png") as cover:
        assert cover.size == (2000, 1500)
    with Image.open(tmp_path / "marketing_mockups" / "07-interactive-html-calculator-mockup.png") as mockup:
        assert mockup.size == (2000, 1500)


def test_bespoke_campaign_layouts_make_distinct_covers_and_gallery_order(tmp_path):
    screenshot = tmp_path / "dashboard.png"
    Image.new("RGB", (1920, 1080), "#E6F0F0").save(screenshot)
    results = {}
    for layout, color, art_name, identity_name in (
        ("moving_atlas", "#2D8F8B", "atlas.png", "A Clear Landing"),
        ("interiors_editorial", "#C4775A", "interiors.png", "Measured Home"),
    ):
        art = tmp_path / art_name
        Image.new("RGB", (1536, 1024), color).save(art)
        spec = {
            "title": identity_name,
            "subtitle": "A practical product-specific planning kit.",
            "functional_scope": ["Budget", "Quotes", "Project timeline"],
            "palette": {"accent": color},
            "visual_identity": {
                "name": identity_name,
                "motif": f"Distinctive {layout} motif",
                "marketing_layout": layout,
                "marketing_artwork_file": art.name,
                "marketing_art_direction": f"Original {layout} campaign art",
            },
        }
        generated = marketing_images(spec, [screenshot], tmp_path, {})
        results[layout] = [path.name for path in generated]
        with Image.open(tmp_path / "marketing_mockups" / "00-product-cover.png") as cover:
            assert cover.size == (2000, 1500)
            results[f"{layout}_pixel"] = cover.getpixel((100, 100))

    assert results["moving_atlas"][1] == "03-whats-included-mockup.png"
    assert results["moving_atlas"][2] == "01-route-ledger-mockup.png"
    assert results["interiors_editorial"][1] == "03-whats-included-mockup.png"
    assert results["interiors_editorial"][2] == "01-editorial-spread-mockup.png"
    assert results["moving_atlas_pixel"] != results["interiors_editorial_pixel"]


def test_aquarium_campaign_pairs_excel_and_html_tool_on_cover(tmp_path):
    screenshot = tmp_path / "dashboard.png"
    calculator = tmp_path / "calculator.png"
    artwork = tmp_path / "aquarium-art.png"
    Image.new("RGB", (1920, 1080), "#E6F0F0").save(screenshot)
    Image.new("RGB", (1920, 1080), "#D9EEE8").save(calculator)
    Image.new("RGB", (1536, 1024), "#3B8176").save(artwork)
    spec = {
        "title": "Freshwater Aquarium Setup Planner",
        "subtitle": "Plan the habitat before the first purchase.",
        "functional_scope": ["Budget", "Equipment", "Readiness"],
        "design_language": "en",
        "visual_identity": {
            "name": "Stillwater field notes",
            "motif": "Aquarium specimen window",
            "marketing_layout": "aquarium_fieldguide",
            "marketing_artwork_file": artwork.name,
        },
    }
    manifest = {
        "interactive_html_tool": {
            "passed": True,
            "title": "Freshwater Tank Volume Calculator",
            "filename": "freshwater-tank-volume-calculator.html",
            "language": "en",
        }
    }

    generated = marketing_images(spec, [screenshot], tmp_path, manifest, calculator)

    assert "07-interactive-html-calculator-mockup.png" in {path.name for path in generated}
    with Image.open(tmp_path / "marketing_mockups" / "00-product-cover.png") as cover:
        assert cover.size == (2000, 1500)
