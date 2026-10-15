from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


OUT = Path("outputs/listings/product-7/portada.png")
WIDTH, HEIGHT = 2000, 1500


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arialbd.ttf" if bold else "C:/Windows/Fonts/arial.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def rounded(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], fill: str, radius: int) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill)


def main() -> None:
    image = Image.new("RGB", (WIDTH, HEIGHT), "#F5F1E8")
    draw = ImageDraw.Draw(image)
    navy = "#142B45"
    teal = "#2E9C9A"
    coral = "#F18A72"
    cream = "#F5F1E8"
    ink = "#203449"
    muted = "#6D7D89"

    draw.rectangle((0, 0, WIDTH, HEIGHT), fill=cream)
    draw.rectangle((0, 0, 850, HEIGHT), fill=navy)
    draw.ellipse((1430, -320, 2180, 430), fill="#E6D9C7")
    draw.ellipse((1320, 1160, 2140, 1900), fill="#D7ECE8")

    # Left text panel.
    draw.text((115, 170), "MAQUETAS EXCEL", font=font(36, True), fill="#8FE0D4")
    draw.text((115, 360), "PLANIFICADOR", font=font(84, True), fill="white")
    draw.text((115, 455), "DE PRESUPUESTO", font=font(84, True), fill="white")
    draw.text((115, 550), "FAMILIAR", font=font(84, True), fill="#8FE0D4")
    draw.text(
        (120, 735),
        "Organiza tus gastos\ny revisa tu mes de un vistazo.",
        font=font(38),
        fill="#E4EEF0",
        spacing=12,
    )
    rounded(draw, (115, 990, 555, 1070), coral, 36)
    draw.text((153, 1010), "PLANTILLA XLSX", font=font(30, True), fill=navy)
    draw.text((120, 1220), "Descarga digital", font=font(34, True), fill="white")
    draw.text((120, 1270), "Presupuesto · gastos · vista mensual", font=font(25), fill="#C9DBDF")

    # Laptop shadow and body.
    rounded(draw, (930, 285, 1845, 1035), "#203449", 42)
    rounded(draw, (985, 340, 1790, 970), "#EAF4F1", 20)
    draw.rectangle((1015, 370, 1760, 530), fill="#D8EFEB")
    draw.text((1055, 405), "PRESUPUESTO MENSUAL", font=font(34, True), fill=navy)
    draw.text((1055, 458), "Vista general", font=font(25), fill=muted)

    # Dashboard cards.
    cards = [
        (1055, 580, 1250, 710, "INGRESOS", "4.200", teal),
        (1280, 580, 1475, 710, "GASTOS", "2.180", coral),
        (1505, 580, 1700, 710, "AHORRO", "2.020", "#D5B45B"),
    ]
    for x1, y1, x2, y2, label, value, color in cards:
        rounded(draw, (x1, y1, x2, y2), "white", 18)
        draw.rectangle((x1, y1, x1 + 10, y2), fill=color)
        draw.text((x1 + 28, y1 + 20), label, font=font(19, True), fill=muted)
        draw.text((x1 + 28, y1 + 57), value, font=font(36, True), fill=ink)

    # Spreadsheet grid.
    rounded(draw, (1055, 760, 1700, 900), "white", 16)
    draw.text((1080, 782), "Categoría", font=font(20, True), fill=navy)
    draw.text((1350, 782), "Plan", font=font(20, True), fill=navy)
    draw.text((1510, 782), "Real", font=font(20, True), fill=navy)
    for row in range(4):
        y = 820 + row * 20
        draw.line((1080, y, 1675, y), fill="#D9E5E1", width=2)
    for x in (1305, 1470):
        draw.line((x, 770, x, 890), fill="#D9E5E1", width=2)
    draw.text((1080, 830), "Hogar", font=font(17), fill=ink)
    draw.text((1350, 830), "900", font=font(17), fill=muted)
    draw.text((1510, 830), "860", font=font(17), fill=muted)
    draw.text((1080, 860), "Alimentación", font=font(17), fill=ink)
    draw.text((1350, 860), "450", font=font(17), fill=muted)
    draw.text((1510, 860), "420", font=font(17), fill=muted)

    # Laptop base.
    draw.polygon([(820, 1050), (1900, 1050), (1815, 1150), (900, 1150)], fill="#263D54")
    draw.polygon([(930, 1060), (1790, 1060), (1745, 1100), (975, 1100)], fill="#6D8991")
    rounded(draw, (1310, 1115, 1440, 1130), "#A6C0C2", 8)
    draw.text((1165, 1250), "Una forma sencilla de ver tu mes", font=font(29, True), fill=navy)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    image.save(OUT, format="PNG", optimize=True)
    print(OUT)


if __name__ == "__main__":
    main()
