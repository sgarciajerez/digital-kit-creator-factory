"""Build a branded, paginated product e-book from its declared manuscript."""
from __future__ import annotations

import json
import re
import argparse
import hashlib
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Flowable,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
)


ROOT = Path(__file__).resolve().parents[2]
SECTION_COUNT = 10
FONT_SIZE = 12
MIN_SECTION_WORDS = 180
MIN_CONTINUATION_HEIGHT = 150


class SectionStart(Flowable):
    def __init__(self, index: int, starts: dict[int, int]) -> None:
        super().__init__()
        self.index = index
        self.starts = starts
        self.width = 0
        self.height = 0

    def draw(self) -> None:
        self.starts[self.index] = self.canv.getPageNumber()


def _resolve_project_path(value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else ROOT / path


def _register_fonts() -> tuple[str, str]:
    regular = Path("C:/Windows/Fonts/arial.ttf")
    bold = Path("C:/Windows/Fonts/arialbd.ttf")
    if regular.is_file() and bold.is_file():
        if "EbookArial" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("EbookArial", str(regular)))
            pdfmetrics.registerFont(TTFont("EbookArial-Bold", str(bold)))
        return "EbookArial", "EbookArial-Bold"
    return "Helvetica", "Helvetica-Bold"


def _hex_color(value: str) -> colors.Color:
    normalized = str(value).strip().removeprefix("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", normalized):
        raise ValueError(f"Color de paleta no válido para el e-book: {value!r}")
    return colors.HexColor(f"#{normalized}")


def _validate_content(content: dict) -> list[dict]:
    sections = content.get("sections")
    if not isinstance(sections, list) or len(sections) != SECTION_COUNT:
        raise ValueError(f"El e-book debe declarar exactamente {SECTION_COUNT} puntos")
    titles: set[str] = set()
    for index, section in enumerate(sections, start=1):
        title = str(section.get("title", "")).strip()
        paragraphs = section.get("paragraphs")
        if not title or title.casefold() in titles:
            raise ValueError(f"El punto {index} debe tener un título no vacío y único")
        titles.add(title.casefold())
        if not isinstance(paragraphs, list) or not paragraphs or any(not str(item).strip() for item in paragraphs):
            raise ValueError(f"El punto {index} debe incluir párrafos con contenido")
        word_count = len(re.findall(r"\b[\w’'-]+\b", " ".join(map(str, paragraphs)), flags=re.UNICODE))
        if word_count < MIN_SECTION_WORDS:
            raise ValueError(
                f"El punto {index} solo tiene {word_count} palabras; el mínimo editorial es {MIN_SECTION_WORDS}"
            )
        section["word_count"] = word_count
    return sections


def generate_product_ebook(spec: dict, output_path: Path) -> dict:
    """Create the required cover + blank leaf + contents + ten chapter PDF."""
    ebook = spec.get("ebook") or {}
    for key in ("title", "content_file", "cover_image"):
        if not ebook.get(key):
            raise ValueError(f"{spec.get('slug', 'producto')}: falta ebook.{key}")

    content_path = _resolve_project_path(ebook["content_file"])
    cover_path = _resolve_project_path(ebook["cover_image"])
    if not content_path.is_file():
        raise FileNotFoundError(f"Falta el manuscrito del e-book: {content_path}")
    if not cover_path.is_file():
        raise FileNotFoundError(
            f"Falta la portada generada con IA: {cover_path}. Genérala para este producto antes de compilarlo."
        )
    content = json.loads(content_path.read_text(encoding="utf-8"))
    sections = _validate_content(content)
    if len(content.get("sections", [])) != SECTION_COUNT:
        raise ValueError("El índice debe coincidir con los diez puntos del manuscrito")

    palette = spec.get("palette") or {}
    required_colors = ("background", "accent", "secondary", "highlight")
    missing = [key for key in required_colors if not palette.get(key)]
    if missing:
        raise ValueError(f"Faltan colores para diseñar el e-book: {', '.join(missing)}")
    pale = _hex_color(palette["background"])
    accent = _hex_color(palette["accent"])
    secondary = _hex_color(palette["secondary"])
    highlight = _hex_color(palette["highlight"])
    body_font, bold_font = _register_fonts()

    page_width, page_height = A4
    left_margin = right_margin = 54
    top_margin = bottom_margin = 52
    output_path.parent.mkdir(parents=True, exist_ok=True)
    page_numbers: list[int] = []
    section_starts: dict[int, int] = {}

    def cover_page(canvas, doc) -> None:
        page_numbers.append(doc.page)
        canvas.saveState()
        canvas.setFillColor(pale)
        canvas.rect(0, 0, page_width, page_height, fill=1, stroke=0)
        from PIL import Image, ImageOps

        with Image.open(cover_path) as source:
            fitted = ImageOps.fit(source.convert("RGB"), (int(page_width * 2), int(page_height * 2)), centering=(0.5, 0.48))
            canvas.drawImage(ImageReader(fitted), 0, 0, width=page_width, height=page_height, mask="auto")
        band_height = 148
        canvas.setFillColor(pale)
        canvas.rect(0, 0, page_width, band_height, fill=1, stroke=0)
        canvas.setFillColor(accent)
        canvas.rect(0, band_height - 5, page_width, 5, fill=1, stroke=0)
        canvas.setFillColor(colors.black)
        canvas.setFont(bold_font, FONT_SIZE)
        canvas.drawString(left_margin, 105, str(ebook["title"]))
        canvas.setFont(body_font, FONT_SIZE)
        canvas.drawString(left_margin, 79, str(content.get("subtitle", "Ten practical ideas for your next step")))
        canvas.setFillColor(secondary)
        canvas.rect(left_margin, 56, 70, 3, fill=1, stroke=0)
        canvas.setFillColor(colors.black)
        canvas.setFont(body_font, FONT_SIZE)
        canvas.drawString(left_margin, 32, "10 practical tips and ideas")
        canvas.restoreState()

    def blank_page(canvas, doc) -> None:
        page_numbers.append(doc.page)

    def body_page(canvas, doc) -> None:
        page_numbers.append(doc.page)
        canvas.saveState()
        canvas.setStrokeColor(highlight)
        canvas.setLineWidth(2)
        canvas.line(left_margin, page_height - 30, page_width - right_margin, page_height - 30)
        canvas.setFillColor(secondary)
        canvas.rect(left_margin, page_height - 33, 45, 3, fill=1, stroke=0)
        canvas.setFillColor(colors.black)
        canvas.setFont(body_font, FONT_SIZE)
        canvas.drawRightString(page_width - right_margin, 25, str(doc.page))
        canvas.restoreState()

    doc = BaseDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
        title=str(ebook["title"]),
        author=str(spec.get("title", "Product guide")),
        subject=str(content.get("subtitle", "Ten practical tips")),
    )
    frame = Frame(left_margin, bottom_margin, page_width - left_margin - right_margin, page_height - top_margin - bottom_margin, id="body", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([
        PageTemplate(id="Cover", frames=frame, onPage=cover_page),
        PageTemplate(id="Blank", frames=frame, onPage=blank_page),
        PageTemplate(id="Body", frames=frame, onPage=body_page),
    ])

    body_style = ParagraphStyle(
        "EbookBody", fontName=body_font, fontSize=FONT_SIZE, leading=17,
        textColor=colors.black, alignment=TA_LEFT, spaceAfter=13,
    )
    heading_style = ParagraphStyle(
        "EbookHeading", fontName=bold_font, fontSize=FONT_SIZE, leading=17,
        textColor=colors.black, backColor=pale, borderPadding=(8, 8, 8, 8),
        spaceAfter=18, keepWithNext=True,
    )
    toc_style = ParagraphStyle(
        "EbookContents", fontName=body_font, fontSize=FONT_SIZE, leading=22,
        textColor=colors.black, leftIndent=8,
    )

    story = [Spacer(1, 1), NextPageTemplate("Blank"), PageBreak(), NextPageTemplate("Body"), PageBreak()]
    story.extend([
        Paragraph(escape("Contents"), ParagraphStyle("ContentsHeading", parent=heading_style, spaceAfter=20)),
        Paragraph(escape(str(content.get("contents_intro", "Ten practical ideas for your next project"))), body_style),
        Spacer(1, 8),
    ])
    story.extend(Paragraph(escape(section["title"]), toc_style) for section in sections)
    story.append(PageBreak())

    available_width = page_width - left_margin - right_margin
    available_height = page_height - top_margin - bottom_margin
    for index, section in enumerate(sections):
        blocks = [Paragraph(escape(section["title"]), heading_style)]
        blocks.extend(Paragraph(escape(str(paragraph)), body_style) for paragraph in section["paragraphs"])
        blocks.append(Spacer(1, 4))
        measured = sum(block.wrap(available_width, available_height)[1] for block in blocks)
        if measured > available_height and measured - available_height < MIN_CONTINUATION_HEIGHT:
            raise ValueError(
                f"El punto {index + 1} supera una página ({measured:.0f} pt); edítalo para evitar una continuación mínima."
            )
        story.append(SectionStart(index + 1, section_starts))
        story.extend(blocks)
        if index < len(sections) - 1:
            story.append(PageBreak())

    story.extend([
        PageBreak(),
        Paragraph("Sources and further reading", ParagraphStyle("SourcesHeading", parent=heading_style, spaceAfter=18)),
        Paragraph(escape(str(content.get("sources_note", "The following sources provide further reading; follow qualified local guidance for individual decisions."))), body_style),
    ])
    source_style = ParagraphStyle("EbookSource", parent=body_style, fontSize=FONT_SIZE, leading=17, spaceAfter=10)
    for source in content.get("sources", []):
        author = escape(str(source.get("author", "")))
        title = escape(str(source.get("title", "")))
        url = escape(str(source.get("url", "")), {'"': "&quot;"})
        story.append(Paragraph(f'<b>{author}</b>. {title}. <link href="{url}" color="#000000">{url}</link>', source_style))
    doc.build(story)
    page_count = max(page_numbers, default=0)
    expected_pages = SECTION_COUNT + 4
    if page_count < SECTION_COUNT + 4:
        output_path.unlink(missing_ok=True)
        raise ValueError(f"El e-book tiene {page_count} páginas; se esperaban {expected_pages}")
    return {
        "passed": True,
        "file": str(output_path),
        "sha256": hashlib.sha256(output_path.read_bytes()).hexdigest(),
        "page_count": page_count,
        "section_count": len(sections),
        "section_word_counts": [section["word_count"] for section in sections],
        "body_font_size_pt": FONT_SIZE,
        "cover_image": str(cover_path),
        "cover_sha256": hashlib.sha256(cover_path.read_bytes()).hexdigest(),
        "content_file": str(content_path),
        "content_sha256": hashlib.sha256(content_path.read_bytes()).hexdigest(),
        "blank_page": 2,
        "contents_page": 3,
        "section_pages": [
            {"section": index, "start": section_starts[index],
             "end": max(section_starts[index], section_starts.get(index + 1, page_count - 1) - 1)}
            for index in range(1, SECTION_COUNT + 1)
        ],
        "sources_page": page_count,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Compila y valida el e-book de un producto")
    parser.add_argument("--spec", required=True, help="Ruta al design_spec.json del producto")
    parser.add_argument("--output", help="Ruta PDF; por defecto usa app/ y el nombre declarado")
    args = parser.parse_args()
    spec_path = _resolve_project_path(args.spec)
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    filename = str((spec.get("ebook") or {}).get("filename", ""))
    if not filename.lower().endswith(".pdf") or Path(filename).name != filename:
        raise ValueError("ebook.filename debe ser un nombre PDF descriptivo sin rutas")
    output = _resolve_project_path(args.output) if args.output else spec_path.parent / "app" / filename
    report = generate_product_ebook(spec, output)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
