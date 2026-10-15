from __future__ import annotations

import json
import shutil
from pathlib import Path

import imageio.v2 as imageio
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent.parent
PRODUCT = ROOT / "outputs" / "listings" / "product-12"
WORKBOOK = PRODUCT / "app" / "home-renovation-planner-excel-template.xlsx"
OUT = PRODUCT / "marketing_assets"
SIZE = (2000, 1500)
BG = "#F8F4EF"
INK = "#3E3530"
ACCENT = "#C4775A"
SAGE = "#6F8A73"
GOLD = "#D5A84F"
MUTED = "#756D67"
GRID = "#E4D9D0"
WHITE = "#FFFFFF"


def font(size: int, bold: bool = False):
    path = Path(f"C:/Windows/Fonts/{'segoeuib' if bold else 'segoeui'}.ttf")
    return ImageFont.truetype(path, size) if path.exists() else ImageFont.load_default()


def text(draw, xy, value, size=28, fill=INK, bold=False):
    draw.text(xy, str(value), font=font(size, bold), fill=fill)


def base(eyebrow: str, title: str, subtitle: str):
    image = Image.new("RGB", SIZE, BG)
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, SIZE[0], 96), fill=INK)
    text(draw, (80, 28), "MAQUETAS EXCEL  ·  DIGITAL DOWNLOAD", 25, "#F0D9A7", True)
    text(draw, (105, 160), eyebrow.upper(), 24, ACCENT, True)
    text(draw, (105, 215), title, 62, INK, True)
    text(draw, (108, 300), subtitle, 28, MUTED)
    return image, draw


def card(draw, box, title, value, color=ACCENT):
    x1, y1, x2, y2 = box
    draw.rounded_rectangle(box, radius=22, fill=WHITE, outline=GRID, width=3)
    draw.rectangle((x1, y1, x1 + 14, y2), fill=color)
    text(draw, (x1 + 38, y1 + 32), title, 23, MUTED, True)
    text(draw, (x1 + 38, y1 + 88), value, 48, INK, True)


def cover(wb):
    image = Image.new("RGB", SIZE, BG)
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, 780, SIZE[1]), fill=INK)
    text(draw, (105, 150), "MAQUETAS EXCEL", 34, "#F0D9A7", True)
    text(draw, (105, 370), "HOME", 88, WHITE, True)
    text(draw, (105, 470), "RENOVATION", 88, WHITE, True)
    text(draw, (105, 570), "PLANNER", 88, "#E5B99D", True)
    text(draw, (110, 770), "Budget · purchases · quotes\nTimeline · dashboard", 34, "#EFE8E1")
    draw.rounded_rectangle((105, 1045, 555, 1128), radius=34, fill=ACCENT)
    text(draw, (150, 1068), "EDITABLE XLSX", 29, WHITE, True)
    text(draw, (110, 1280), "Instant digital download", 30, WHITE, True)
    # Workbook preview built from values read from the real source workbook.
    draw.rounded_rectangle((940, 260, 1840, 1120), radius=34, fill="#D8D0C8")
    draw.rounded_rectangle((990, 305, 1790, 1020), radius=24, fill=WHITE)
    text(draw, (1045, 365), str(wb["Dashboard"]["B3"].value), 37, INK, True)
    text(draw, (1045, 430), "Connected renovation control panel", 23, MUTED)
    card(draw, (1045, 525, 1260, 700), "BUDGET CAP", "€ 24,000", ACCENT)
    card(draw, (1290, 525, 1505, 700), "ACTUAL SPEND", "€ 7,400", SAGE)
    card(draw, (1535, 525, 1750, 700), "AVAILABLE", "€ 16,600", GOLD)
    draw.rounded_rectangle((1045, 770, 1750, 940), radius=18, fill="#F7F2ED", outline=GRID, width=2)
    text(draw, (1080, 802), "Budget by renovation area", 25, INK, True)
    for idx, (label, amount, color) in enumerate((("Kitchen", 8500, ACCENT), ("Bathroom", 5100, SAGE), ("Living room", 3200, GOLD))):
        y = 855 + idx * 36
        text(draw, (1080, y), label, 19, MUTED)
        draw.rounded_rectangle((1280, y + 4, 1685, y + 22), radius=9, fill="#EDE4DC")
        draw.rounded_rectangle((1280, y + 4, 1280 + int(amount / 8500 * 405), y + 22), radius=9, fill=color)
    return image


def dashboard(wb):
    image, draw = base("Workbook preview", "See the next renovation decision", "A compact control panel connects the budget, purchases and timeline.")
    card(draw, (110, 470, 520, 720), "BUDGET CAP", "€ 24,000", ACCENT)
    card(draw, (590, 470, 1000, 720), "ACTUAL SPEND", "€ 7,400", SAGE)
    card(draw, (1070, 470, 1480, 720), "AVAILABLE", "€ 16,600", GOLD)
    card(draw, (1550, 470, 1890, 720), "OPEN TASKS", "4", INK)
    draw.rounded_rectangle((110, 820, 1890, 1250), radius=24, fill=WHITE, outline=GRID, width=3)
    text(draw, (160, 870), "What the workbook keeps connected", 32, INK, True)
    items = [("01", "Budget by area", "Planned, actual and remaining spend", ACCENT), ("02", "Purchase log", "One source for material and supplier costs", SAGE), ("03", "Timeline", "Due dates with days-left formulas", GOLD)]
    for idx, (num, heading, caption, color) in enumerate(items):
        x = 165 + idx * 560
        draw.ellipse((x, 980, x + 78, 1058), fill=color)
        text(draw, (x + 20, 998), num, 21, WHITE, True)
        text(draw, (x, 1100), heading, 26, INK, True)
        text(draw, (x, 1145), caption, 20, MUTED)
    return image


def table_view(wb):
    image, draw = base("Workbook preview", "Track costs without double entry", "Purchases flow into the budget by area so your remaining amount stays current.")
    x, y = 110, 500
    widths = [280, 260, 420, 270, 220, 230]
    headers = ["Purchase date", "Area", "Item", "Supplier", "Amount", "Status"]
    rows = []
    for row in range(6, 10):
        rows.append([wb["Purchases"].cell(row, col).value or "—" for col in (1, 2, 3, 4, 5, 7)])
    draw.rounded_rectangle((x, y, x + sum(widths), y + 620), radius=22, fill=WHITE, outline=GRID, width=3)
    cursor = x
    for width, header in zip(widths, headers):
        draw.rectangle((cursor, y, cursor + width, y + 72), fill="#F0E1D8")
        text(draw, (cursor + 16, y + 24), header, 19, INK, True)
        cursor += width
    for ridx, row in enumerate(rows):
        top = y + 72 + ridx * 86
        if ridx % 2:
            draw.rectangle((x, top, x + sum(widths), top + 86), fill="#FCFAF8")
        cursor = x
        for width, value in zip(widths, row):
            draw.line((cursor, top, cursor, top + 86), fill=GRID, width=2)
            shown = value.strftime("%d/%m/%Y") if hasattr(value, "strftime") else str(value)
            text(draw, (cursor + 16, top + 28), shown, 20, INK)
            cursor += width
        draw.line((x, top + 86, x + sum(widths), top + 86), fill=GRID, width=2)
    draw.rounded_rectangle((110, 1225, 1890, 1310), radius=22, fill="#E7EFE8")
    text(draw, (155, 1250), "Edit the input columns; leave the derived columns to Excel.", 25, SAGE, True)
    return image


def workflow(wb):
    image, draw = base("How it works", "From quote to next task", "Set the project rules once, then move through the workbook in a simple sequence.")
    steps = [("1", "Set up", "Name the project, budget cap and dates.", ACCENT), ("2", "Log", "Add purchases and compare supplier quotes.", SAGE), ("3", "Decide", "Use the dashboard to see spend and time left.", GOLD)]
    for idx, (number, title, caption, color) in enumerate(steps):
        x = 130 + idx * 610
        draw.rounded_rectangle((x, 520, x + 500, 1130), radius=26, fill=WHITE, outline=GRID, width=3)
        draw.ellipse((x + 48, 590, x + 148, 690), fill=color)
        text(draw, (x + 84, 616), number, 32, WHITE, True)
        text(draw, (x + 48, 790), title, 38, INK, True)
        text(draw, (x + 48, 880), caption, 25, MUTED)
    return image


def make_video(images: list[Path], path: Path):
    with imageio.get_writer(path, fps=24, codec="libx264", quality=7, pixelformat="yuv420p", macro_block_size=1) as writer:
        for image_path in images:
            frame = np.asarray(Image.open(image_path).convert("RGB").resize((1920, 960), Image.Resampling.LANCZOS))
            for _ in range(42):
                writer.append_data(frame)


def main() -> None:
    wb = load_workbook(WORKBOOK, data_only=False)
    OUT.mkdir(parents=True, exist_ok=True)
    renders = [("01-cover.png", cover(wb)), ("02-dashboard.png", dashboard(wb)), ("03-purchases.png", table_view(wb)), ("04-workflow.png", workflow(wb))]
    files = []
    for name, image in renders:
        path = OUT / name
        image.save(path, format="PNG", optimize=True)
        files.append(path)
    video = OUT / "home-renovation-planner-demo.mp4"
    make_video(files[1:], video)
    manifest = {"source_workbook": str(WORKBOOK.relative_to(ROOT)), "images": [p.name for p in files], "video": video.name, "image_size": list(SIZE), "video_size": [1920, 960], "video_seconds": round(42 * 3 / 24, 2)}
    (OUT / "marketing_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    draft_dir = PRODUCT / "draft_package"
    draft_dir.mkdir(parents=True, exist_ok=True)
    for asset in [*files, video]:
        shutil.copy2(asset, draft_dir / asset.name)
    draft = PRODUCT / "draft_package" / "listing.json"
    data = json.loads(draft.read_text(encoding="utf-8"))
    data["image_files"] = [p.name for p in files]
    data["video_file"] = video.name
    data["marketing_assets"] = str(OUT.relative_to(ROOT))
    draft.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    wb.close()
    print(f"MARKETING_ASSETS={OUT}")
    print(f"MARKETING_IMAGES={len(files)}")
    print(f"MARKETING_VIDEO={video}")


if __name__ == "__main__":
    main()
