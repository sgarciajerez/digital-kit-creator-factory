from __future__ import annotations

from pathlib import Path
import textwrap

from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont


PROJECT_ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "autopilot_3_Planificador_de_presupuesto_familiar_0.1.0.xlsx"
OUT = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "screenshots"
WIDTH, HEIGHT = 2000, 1500

NAVY = "#142B45"
INK = "#203449"
TEAL = "#2E9C9A"
MINT = "#D8EFEB"
CORAL = "#F18A72"
CREAM = "#F5F1E8"
GRID = "#D9E5E1"
MUTED = "#6D7D89"
WHITE = "#FFFFFF"


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
    ]
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def base(active: str, title: str, subtitle: str) -> tuple[Image.Image, ImageDraw.ImageDraw]:
    image = Image.new("RGB", (WIDTH, HEIGHT), "#F7F9F8")
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, WIDTH, 86), fill=NAVY)
    draw.text((42, 24), "Planificador de presupuesto familiar · Excel", font=font(32, True), fill=WHITE)
    draw.text((42, 126), title, font=font(58, True), fill=INK)
    draw.text((44, 198), subtitle, font=font(28), fill=MUTED)
    draw.rectangle((0, 1385, WIDTH, HEIGHT), fill="#E8EFEC")
    tabs = ["Instrucciones", "Presupuesto", "Vista previa", "Configuración", "Gastos compartidos"]
    x = 42
    for tab in tabs:
        width = max(185, int(draw.textlength(tab, font=font(25)) + 56))
        fill = WHITE if tab == active else "#D7E4E0"
        draw.rounded_rectangle((x, 1403, x + width, 1478), radius=10, fill=fill, outline=GRID, width=2)
        draw.text((x + 28, 1424), tab, font=font(25, tab == active), fill=INK if tab == active else MUTED)
        x += width + 12
    return image, draw


def table(draw: ImageDraw.ImageDraw, x: int, y: int, widths: list[int], headers: list[str], rows: list[list[str]], row_h: int = 76) -> None:
    total = sum(widths)
    draw.rounded_rectangle((x, y, x + total, y + row_h * (len(rows) + 1)), radius=14, fill=WHITE, outline=GRID, width=2)
    cursor = x
    for width, header in zip(widths, headers):
        draw.rectangle((cursor, y, cursor + width, y + row_h), fill=MINT)
        draw.text((cursor + 20, y + 22), header, font=font(24, True), fill=NAVY)
        cursor += width
    for row_index, row in enumerate(rows):
        top = y + row_h * (row_index + 1)
        if row_index % 2:
            draw.rectangle((x, top, x + total, top + row_h), fill="#FBFDFC")
        cursor = x
        for width, value in zip(widths, row):
            draw.line((cursor, top, cursor, top + row_h), fill=GRID, width=2)
            draw.text((cursor + 20, top + 22), str(value), font=font(24), fill=INK)
            cursor += width
        draw.line((x, top, x + total, top), fill=GRID, width=2)
    draw.line((x, y + row_h * (len(rows) + 1), x + total, y + row_h * (len(rows) + 1)), fill=GRID, width=2)


def money(value: int) -> str:
    return f"€ {value:,.0f}".replace(",", ".")


def instructions(wb) -> Image.Image:
    image, draw = base("Instrucciones", "Empieza en tres pasos", "Una guía breve para configurar tu plantilla y comenzar a organizar el presupuesto del hogar.")
    values = [str(wb["Instrucciones"][f"A{i}"].value or "") for i in range(1, 8)]
    draw.rounded_rectangle((130, 310, 1870, 560), radius=26, fill=NAVY)
    draw.text((185, 355), values[0], font=font(52, True), fill=WHITE)
    draw.text((185, 440), values[1], font=font(27), fill="#DDEBED")
    steps = [
        ("1", "Introduce tus cifras", values[4]),
        ("2", "Revisa la desviación", values[5]),
        ("3", "Consulta el resumen", values[6]),
    ]
    for index, (number, heading, text) in enumerate(steps):
        x = 130 + index * 585
        draw.rounded_rectangle((x, 660, x + 520, 1110), radius=24, fill=WHITE, outline=GRID, width=3)
        draw.ellipse((x + 34, 705, x + 124, 795), fill=TEAL if index != 1 else CORAL)
        draw.text((x + 66, 723), number, font=font(40, True), fill=WHITE)
        draw.text((x + 34, 855), heading, font=font(31, True), fill=INK)
        wrapped = "\n".join(textwrap.wrap(text, width=31))
        draw.multiline_text((x + 34, 920), wrapped, font=font(24), fill=MUTED, spacing=8)
    draw.text((130, 1210), "Vista previa · datos de ejemplo para mostrar el funcionamiento de la plantilla", font=font(26, True), fill=TEAL)
    return image


def budget(wb) -> Image.Image:
    image, draw = base("Presupuesto", "Presupuesto mensual", "Compara lo que habías previsto con el gasto real y detecta desviaciones rápidamente.")
    categories = [str(wb["Presupuesto"][f"A{i}"].value or "") for i in range(2, 9)]
    planned = [950, 420, 180, 90, 75, 260, 135]
    actual = [930, 398, 164, 110, 62, 245, 128]
    rows = [[cat, money(p), money(a), money(p-a), "Dentro del plan"] for cat, p, a in zip(categories, planned, actual)]
    rows.append(["TOTAL", money(sum(planned)), money(sum(actual)), money(sum(planned)-sum(actual)), "Resumen"])
    table(draw, 120, 340, [430, 270, 270, 270, 470], ["Categoría", "Presupuesto", "Gasto real", "Desviación", "Notas"], rows, 74)
    draw.rounded_rectangle((120, 1260, 1880, 1340), radius=20, fill=MINT)
    draw.text((155, 1282), "Ejemplo visual: sustituye las cifras por tus propios datos en las celdas de entrada.", font=font(26, True), fill=NAVY)
    return image


def preview() -> Image.Image:
    image, draw = base("Vista previa", "Resumen de un vistazo", "Una pantalla sencilla para comprobar el total presupuestado, el gasto real y la diferencia.")
    cards = [(150, "PRESUPUESTO TOTAL", money(2110), TEAL), (635, "GASTO REAL TOTAL", money(2037), CORAL), (1120, "DIFERENCIA", money(73), "#D5B45B")]
    for x, label, value, color in cards:
        draw.rounded_rectangle((x, 350, x + 390, 590), radius=24, fill=WHITE, outline=GRID, width=3)
        draw.rectangle((x, 350, x + 16, 590), fill=color)
        draw.text((x + 42, 405), label, font=font(24, True), fill=MUTED)
        draw.text((x + 42, 478), value, font=font(56, True), fill=INK)
    draw.rounded_rectangle((150, 710, 1750, 1190), radius=26, fill=WHITE, outline=GRID, width=3)
    draw.text((205, 765), "Distribución del gasto", font=font(34, True), fill=INK)
    bars = [("Vivienda", 930, TEAL), ("Alimentación", 398, CORAL), ("Transporte", 164, "#D5B45B"), ("Otros", 545, "#8BA9B0")]
    max_value = max(value for _, value, _ in bars)
    for index, (label, value, color) in enumerate(bars):
        y = 850 + index * 76
        draw.text((210, y), label, font=font(25), fill=INK)
        draw.rounded_rectangle((500, y + 5, 1500, y + 43), radius=16, fill="#E8EFEC")
        draw.rounded_rectangle((500, y + 5, 500 + int(1000 * value / max_value), y + 43), radius=16, fill=color)
        draw.text((1540, y), money(value), font=font(25, True), fill=INK)
    return image


def settings(wb) -> Image.Image:
    image, draw = base("Configuración", "Configuración del hogar", "Personaliza los nombres y la regla de reparto según tu situación.")
    rows = [
        ["Miembro A", "Ana"],
        ["Miembro B", "Carlos"],
        ["Regla de reparto", "50/50 o personalizada"],
        ["Moneda", "EUR"],
        ["Nota", "No constituye asesoramiento financiero"],
    ]
    table(draw, 220, 360, [560, 940], ["Campo", "Valor de ejemplo"], rows, 92)
    draw.rounded_rectangle((220, 1040, 1720, 1220), radius=24, fill="#FFF4EC", outline="#F4C8B7", width=3)
    draw.text((270, 1090), "Configura primero esta hoja y después completa Presupuesto.", font=font(34, True), fill=INK)
    draw.text((270, 1150), "Los nombres y cifras del ejemplo se pueden reemplazar por los tuyos.", font=font(26), fill=MUTED)
    return image


def shared_expenses() -> Image.Image:
    image, draw = base("Gastos compartidos", "Gastos compartidos", "Registra quién ha pagado cada gasto y consulta el reparto calculado.")
    rows = [
        ["05/01/2026", "Compra semanal", "Alimentación", "Ana", "€ 86", "€ 43", "€ 43", "€ 0"],
        ["08/01/2026", "Electricidad", "Vivienda", "Carlos", "€ 62", "€ 31", "€ 31", "€ 0"],
        ["12/01/2026", "Abono transporte", "Transporte", "Ana", "€ 48", "€ 24", "€ 24", "€ 0"],
        ["19/01/2026", "Farmacia", "Salud", "Carlos", "€ 29", "€ 14,50", "€ 14,50", "€ 0"],
        ["25/01/2026", "Internet", "Vivienda", "Ana", "€ 40", "€ 20", "€ 20", "€ 0"],
    ]
    table(draw, 80, 330, [210, 320, 250, 210, 160, 185, 185, 180], ["Fecha", "Concepto", "Categoría", "Pagado por", "Importe", "Parte A", "Parte B", "Saldo A/B"], rows, 84)
    draw.rounded_rectangle((80, 920, 1920, 1110), radius=24, fill=NAVY)
    draw.text((140, 970), "Reparto automático", font=font(36, True), fill="#8FE0D4")
    draw.text((140, 1035), "Las partes y el saldo se calculan a partir del importe y de quién ha pagado.", font=font(27), fill=WHITE)
    draw.text((80, 1220), "Datos de ejemplo · no contienen información personal real", font=font(26, True), fill=TEAL)
    return image


def main() -> None:
    wb = load_workbook(WORKBOOK, data_only=False)
    OUT.mkdir(parents=True, exist_ok=True)
    outputs = {
        "01-instrucciones.png": instructions(wb),
        "02-presupuesto.png": budget(wb),
        "03-vista-previa.png": preview(),
        "04-configuracion.png": settings(wb),
        "05-gastos-compartidos.png": shared_expenses(),
    }
    for name, image in outputs.items():
        image.save(OUT / name, format="PNG", optimize=True)
        print(OUT / name)


if __name__ == "__main__":
    main()
