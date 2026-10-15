from pathlib import Path

from product_guide_pdf import generate_user_guide_pdf


def test_user_guide_pdf_is_a_readable_markdown_companion(tmp_path: Path):
    markdown = tmp_path / "README.md"
    markdown.write_text(
        "# How to use the workbook\n\n"
        "**Language:** `en`\n\n"
        "## Quick start\n\n"
        "1. Open the `Start here` sheet.\n"
        "2. Enter your inputs and review the dashboard.\n\n"
        "> Keep formulas intact.\n",
        encoding="utf-8",
    )
    output = tmp_path / "user-guide.pdf"

    report = generate_user_guide_pdf(
        markdown,
        output,
        {
            "title": "Garden Planner",
            "palette": {"background": "#F2F5EC", "accent": "#5C8058", "secondary": "#D17A5A"},
        },
    )

    assert report["passed"] is True
    assert report["page_count"] >= 1
    assert report["body_font_size_pt"] == 12
    assert report["markdown_sha256"]
    assert output.read_bytes().startswith(b"%PDF-")


def test_user_guide_pdf_rejects_missing_markdown(tmp_path: Path):
    try:
        generate_user_guide_pdf(tmp_path / "missing.md", tmp_path / "guide.pdf", {"title": "Guide"})
    except FileNotFoundError as error:
        assert "Missing Markdown user guide" in str(error)
    else:
        raise AssertionError("Expected the missing Markdown file to be rejected")
