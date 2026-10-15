"""Render a customer-facing README.md as a readable, branded PDF guide."""
from __future__ import annotations

import hashlib
import re
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import BaseDocTemplate, Frame, PageBreak, PageTemplate, Paragraph, Spacer


def _color(value: str, fallback: str) -> colors.Color:
    raw = str(value or fallback).strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", raw):
        raw = fallback.lstrip("#")
    return colors.HexColor(f"#{raw}")


def _inline_markup(text: str) -> str:
    """Convert the small inline-Markdown subset used by generated guides."""
    text = re.sub(r"[\u2190-\u21ff]", "->", text)
    result = escape(text)
    result = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", result)
    result = re.sub(r"`([^`]+)`", r'<font name="Courier">\1</font>', result)
    result = re.sub(r"\[([^\]]+)\]\((https?://[^ )]+)\)", r'<link href="\2" color="#000000">\1</link>', result)
    return result


def _markdown_story(markdown: str, styles: dict[str, ParagraphStyle]) -> list:
    story = []
    paragraph_lines: list[str] = []

    def flush() -> None:
        if paragraph_lines:
            story.append(Paragraph(_inline_markup(" ".join(line.strip() for line in paragraph_lines)), styles["body"]))
            paragraph_lines.clear()

    for raw_line in markdown.splitlines():
        line = raw_line.strip()
        if not line:
            flush()
            continue
        heading = re.match(r"^(#{1,3})\s+(.+)$", line)
        if heading:
            flush()
            level = len(heading.group(1))
            story.append(Paragraph(_inline_markup(heading.group(2)), styles[f"h{level}"]))
            continue
        if line.startswith("> "):
            flush()
            story.append(Paragraph(_inline_markup(line[2:]), styles["quote"]))
            continue
        bullet = re.match(r"^[-*]\s+(.+)$", line)
        if bullet:
            flush()
            story.append(Paragraph("• " + _inline_markup(bullet.group(1)), styles["list"]))
            continue
        ordered = re.match(r"^(\d+)[.)]\s+(.+)$", line)
        if ordered:
            flush()
            story.append(Paragraph(f'{ordered.group(1)}. {_inline_markup(ordered.group(2))}', styles["list"]))
            continue
        if line in {"---", "***", "___"}:
            flush()
            story.append(Spacer(1, 8))
            continue
        paragraph_lines.append(line)
    flush()
    return story


def generate_user_guide_pdf(markdown_path: Path, output_path: Path, spec: dict) -> dict:
    """Create a PDF companion of the user-facing Markdown guide."""
    if not markdown_path.is_file():
        raise FileNotFoundError(f"Missing Markdown user guide: {markdown_path}")

    palette = spec.get("palette") or {}
    pale = _color(palette.get("background"), "F2F5EC")
    accent = _color(palette.get("accent"), "5C8058")
    secondary = _color(palette.get("secondary"), "D17A5A")
    page_width, page_height = A4
    margin = 54
    page_numbers: list[int] = []

    def decorate(canvas, doc) -> None:
        page_numbers.append(doc.page)
        canvas.saveState()
        canvas.setStrokeColor(pale)
        canvas.setLineWidth(3)
        canvas.line(margin, page_height - 34, page_width - margin, page_height - 34)
        canvas.setStrokeColor(accent)
        canvas.setLineWidth(2)
        canvas.line(margin, page_height - 37, margin + 48, page_height - 37)
        canvas.setFillColor(colors.black)
        canvas.setFont("Helvetica", 9)
        canvas.drawString(margin, 27, str(spec.get("title", "Excel user guide")))
        canvas.drawRightString(page_width - margin, 27, str(doc.page))
        canvas.restoreState()

    styles = {
        "body": ParagraphStyle("GuideBody", fontName="Helvetica", fontSize=12, leading=17, textColor=colors.black, alignment=TA_LEFT, spaceAfter=9),
        "h1": ParagraphStyle("GuideH1", fontName="Helvetica-Bold", fontSize=18, leading=23, textColor=colors.black, backColor=pale, borderPadding=(8, 9, 8, 9), spaceBefore=5, spaceAfter=18, keepWithNext=True),
        "h2": ParagraphStyle("GuideH2", fontName="Helvetica-Bold", fontSize=15, leading=20, textColor=colors.black, backColor=pale, borderPadding=(6, 8, 6, 8), spaceBefore=13, spaceAfter=10, keepWithNext=True),
        "h3": ParagraphStyle("GuideH3", fontName="Helvetica-Bold", fontSize=12, leading=17, textColor=colors.black, borderColor=secondary, borderWidth=0, borderPadding=3, spaceBefore=10, spaceAfter=7, keepWithNext=True),
        "list": ParagraphStyle("GuideList", fontName="Helvetica", fontSize=12, leading=17, textColor=colors.black, leftIndent=16, firstLineIndent=-11, spaceAfter=5),
        "quote": ParagraphStyle("GuideQuote", fontName="Helvetica-Oblique", fontSize=12, leading=17, textColor=colors.black, leftIndent=12, borderColor=accent, borderWidth=2, borderPadding=7, spaceAfter=10),
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = BaseDocTemplate(
        str(output_path), pagesize=A4, leftMargin=margin, rightMargin=margin,
        topMargin=58, bottomMargin=48, title=f"How to use - {spec.get('title', 'Excel workbook')}",
        author="Excel product guide", subject="Customer instructions for the accompanying Excel workbook",
    )
    frame = Frame(margin, 48, page_width - 2 * margin, page_height - 58 - 48, id="guide", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="Guide", frames=frame, onPage=decorate)])
    doc.build(_markdown_story(markdown_path.read_text(encoding="utf-8"), styles))
    return {
        "passed": bool(page_numbers),
        "file": str(output_path),
        "sha256": hashlib.sha256(output_path.read_bytes()).hexdigest(),
        "markdown_file": str(markdown_path),
        "markdown_sha256": hashlib.sha256(markdown_path.read_bytes()).hexdigest(),
        "page_count": max(page_numbers, default=0),
        "body_font_size_pt": 12,
    }
