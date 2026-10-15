from __future__ import annotations

import textwrap
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageFilter


ROOT = Path(__file__).resolve().parent.parent
PRODUCT = ROOT / "outputs" / "listings" / "product-7"
CAPTURES = PRODUCT / "screenshots_recortadas"
OUT = PRODUCT / "marketing_mockups"
SIZE = (2000, 1500)

CREAM = "#F5F1E8"
NAVY = "#142B45"
INK = "#203449"
TEAL = "#2E9C9A"
MINT = "#D8EFEB"
CORAL = "#F18A72"
MUTED = "#6D7D89"
WHITE = "#FFFFFF"
SHADOW = (20, 43, 69, 42)


def font(size: int, bold: bool = False):
    candidates = [
        Path("C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf"),
        Path("C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def shadow_box(image: Image.Image, box, radius=28, fill=WHITE, outline=None):
    x1, y1, x2, y2 = box
    shadow = Image.new("RGBA", image.size, (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.rounded_rectangle((x1 + 12, y1 + 18, x2 + 12, y2 + 18), radius=radius, fill=SHADOW)
    shadow = shadow.filter(ImageFilter.GaussianBlur(18))
    image.alpha_composite(shadow)
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=2)


def fit_image(path: Path, size: tuple[int, int]) -> Image.Image:
    image = Image.open(path).convert("RGB")
    image.thumbnail(size, Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", size, "#EEF3F0")
    canvas.paste(image, ((size[0] - image.width) // 2, (size[1] - image.height) // 2))
    return canvas


def device_frame(image: Image.Image, screen_size: tuple[int, int], label: str = "") -> Image.Image:
    screen = image.copy().convert("RGB")
    screen.thumbnail(screen_size, Image.Resampling.LANCZOS)
    sw, sh = screen_size
    frame = Image.new("RGBA", (sw + 64, sh + 112), (0, 0, 0, 0))
    d = ImageDraw.Draw(frame)
    d.rounded_rectangle((0, 0, sw + 64, sh + 80), radius=34, fill=NAVY)
    d.ellipse((28, 28, 42, 42), fill=CORAL)
    d.ellipse((52, 28, 66, 42), fill="#D5B45B")
    d.ellipse((76, 28, 90, 42), fill=TEAL)
    frame.paste(screen, (32 + (sw - screen.width) // 2, 80 + (sh - screen.height) // 2))
    d.rounded_rectangle((sw // 2 - 150, sh + 80, sw // 2 + 150, sh + 108), radius=14, fill="#B9C6C3")
    if label:
        d.text((sw + 108, 27), label, font=font(22, True), fill="#DDEBED")
    return frame


def header(image: Image.Image, eyebrow: str, title: str, subtitle: str) -> None:
    d = ImageDraw.Draw(image)
    d.text((120, 92), eyebrow.upper(), font=font(24, True), fill=TEAL)
    d.text((120, 140), title, font=font(68, True), fill=NAVY)
    d.multiline_text((124, 235), "\n".join(textwrap.wrap(subtitle, width=68)), font=font(28), fill=MUTED, spacing=8)


def hero() -> Image.Image:
    image = Image.new("RGBA", SIZE, CREAM)
    header(image, "Planificador de presupuesto familiar", "Organiza tu dinero con claridad", "Una plantilla Excel visual y sencilla para planificar, comparar gastos y repartir costes en casa.")
    frame = device_frame(Image.open(CAPTURES / "03-vista-previa.png"), (1180, 690), "EXCEL TEMPLATE")
    image.alpha_composite(frame, (710, 520))
    shadow_box(image, (120, 600, 650, 1010), 28, WHITE, "#E0E9E5")
    d = ImageDraw.Draw(image)
    d.ellipse((170, 660, 252, 742), fill=TEAL)
    d.text((186, 674), "OK", font=font(25, True), fill=WHITE)
    d.text((170, 795), "Todo en un solo lugar", font=font(32, True), fill=INK)
    d.multiline_text((170, 855), "Presupuesto · resumen\nConfiguración · gastos compartidos", font=font(24), fill=MUTED, spacing=10)
    d.rounded_rectangle((120, 1120, 630, 1190), radius=20, fill=NAVY)
    d.text((170, 1137), "Diseñada para empezar hoy", font=font(24, True), fill=WHITE)
    return image.convert("RGB")


def feature_grid() -> Image.Image:
    image = Image.new("RGBA", SIZE, CREAM)
    header(image, "Una plantilla, varias vistas", "Todo lo que necesitas para verlo rápido", "Las pantallas reales del archivo se presentan en una composición limpia para explicar el producto sin perder fidelidad.")
    items = [("02-presupuesto.png", "Planifica y compara", "Presupuesto mensual", TEAL), ("03-vista-previa.png", "Entiende el resumen", "Vista previa", CORAL), ("05-gastos-compartidos.png", "Reparte los gastos", "Gastos compartidos", "#D5B45B")]
    for x, (name, eyebrow, title, color) in zip((90, 690, 1290), items):
        shadow_box(image, (x, 570, x + 520, 1270), 28, WHITE, "#E0E9E5")
        d = ImageDraw.Draw(image)
        d.rounded_rectangle((x + 28, 610, x + 492, 950), radius=18, fill="#EFF4F1")
        image.paste(fit_image(CAPTURES / name, (430, 315)), (x + 45, 625))
        d.text((x + 38, 1000), eyebrow, font=font(21, True), fill=color)
        d.text((x + 38, 1045), title, font=font(31, True), fill=INK)
        d.text((x + 38, 1120), "Datos de ejemplo editables", font=font(22), fill=MUTED)
    return image.convert("RGB")


def what_includes() -> Image.Image:
    image = Image.new("RGBA", SIZE, CREAM)
    header(image, "Qué incluye", "Empieza con una base lista para usar", "Un paquete digital claro, visual y pensado para que el comprador entienda qué recibe antes de descargarlo.")
    image.alpha_composite(device_frame(Image.open(CAPTURES / "02-presupuesto.png"), (1080, 650), "EDITABLE XLSX"), (95, 570))
    shadow_box(image, (1260, 570, 1880, 1210), 28, NAVY, None)
    d = ImageDraw.Draw(image)
    d.text((1320, 645), "Incluye", font=font(38, True), fill="#8FE0D4")
    for index, line in enumerate(["Archivo Excel editable", "5 hojas organizadas", "Fórmulas y resumen", "Gastos compartidos", "Guía de inicio"]):
        y = 735 + index * 78
        d.ellipse((1320, y + 8, 1338, y + 26), fill=TEAL)
        d.text((1360, y), line, font=font(26, True), fill=WHITE)
    d.rounded_rectangle((1320, 1100, 1820, 1160), radius=18, fill=CORAL)
    d.text((1390, 1117), "Descarga digital", font=font(24, True), fill=WHITE)
    return image.convert("RGB")


def how_it_works() -> Image.Image:
    image = Image.new("RGBA", SIZE, CREAM)
    header(image, "Cómo funciona", "De los datos a la decisión", "Una secuencia sencilla para que el comprador vea el valor del producto en pocos segundos.")
    d = ImageDraw.Draw(image)
    steps = [("01", "Configura", "Añade nombres, moneda y regla de reparto.", TEAL), ("02", "Introduce", "Escribe presupuesto y gasto real en las celdas.", CORAL), ("03", "Comprueba", "Consulta el resumen y los saldos calculados.", "#D5B45B")]
    for index, (number, title, text, color) in enumerate(steps):
        x = 150 + index * 610
        shadow_box(image, (x, 610, x + 500, 1080), 30, WHITE, "#E0E9E5")
        d = ImageDraw.Draw(image)
        d.ellipse((x + 42, 670, x + 142, 770), fill=color)
        d.text((x + 68, 691), number, font=font(27, True), fill=WHITE)
        d.text((x + 42, 840), title, font=font(34, True), fill=INK)
        d.multiline_text((x + 42, 910), "\n".join(textwrap.wrap(text, width=27)), font=font(24), fill=MUTED, spacing=9)
    d.rounded_rectangle((150, 1190, 1850, 1270), radius=22, fill=MINT)
    d.text((205, 1213), "Editable · visual · pensado para el día a día", font=font(29, True), fill=NAVY)
    return image.convert("RGB")


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    outputs = {"01-hero-mockup.png": hero(), "02-feature-grid-mockup.png": feature_grid(), "03-que-incluye-mockup.png": what_includes(), "04-como-funciona-mockup.png": how_it_works()}
    for name, image in outputs.items():
        path = OUT / name
        image.save(path, format="PNG", optimize=True)
        print(path)


if __name__ == "__main__":
    main()
