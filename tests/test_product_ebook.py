import json
from pathlib import Path

import pytest

from product_ebook import generate_product_ebook


ROOT = Path(__file__).resolve().parents[1]


def garden_spec() -> dict:
    return {
        "slug": "community-garden-planner",
        "title": "Community Garden Budget & Project Planner",
        "palette": {
            "background": "#F2F5EC",
            "ink": "#2F3B32",
            "accent": "#5C8058",
            "secondary": "#D17A5A",
            "highlight": "#C5A34B",
        },
        "ebook": {
            "title": "A Greener Community Garden",
            "content_file": "config/ebooks/community-garden-planner.json",
            "cover_image": "assets/ebook-covers/community-garden-planner-cover.png",
        },
    }


def test_product_ebook_has_cover_blank_leaf_contents_and_ten_full_sections(tmp_path):
    result = generate_product_ebook(garden_spec(), tmp_path / "garden-guide.pdf")

    assert result["passed"] is True
    assert result["page_count"] >= 14
    assert result["blank_page"] == 2
    assert result["contents_page"] == 3
    assert len(result["section_pages"]) == 10
    assert any(section["end"] > section["start"] for section in result["section_pages"])
    assert result["sources_page"] == result["page_count"]
    assert result["section_count"] == 10
    assert min(result["section_word_counts"]) >= 180
    assert max(result["section_word_counts"]) - min(result["section_word_counts"]) >= 50
    assert result["body_font_size_pt"] == 12
    assert (tmp_path / "garden-guide.pdf").read_bytes().startswith(b"%PDF-")


def test_ebook_rejects_short_advice_sections(tmp_path):
    content_path = tmp_path / "short.json"
    content = json.loads((ROOT / "config/ebooks/community-garden-planner.json").read_text(encoding="utf-8"))
    content["sections"][0]["paragraphs"] = ["Only a few short words."]
    content_path.write_text(json.dumps(content), encoding="utf-8")
    spec = garden_spec()
    spec["ebook"]["content_file"] = str(content_path)

    with pytest.raises(ValueError, match="mínimo editorial"):
        generate_product_ebook(spec, tmp_path / "short.pdf")


def test_ebook_rejects_missing_ai_cover(tmp_path):
    spec = garden_spec()
    spec["ebook"]["cover_image"] = str(tmp_path / "missing-cover.png")

    with pytest.raises(FileNotFoundError, match="portada generada con IA"):
        generate_product_ebook(spec, tmp_path / "no-cover.pdf")
