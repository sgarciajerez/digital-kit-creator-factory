from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets" / "brand"


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


def logo() -> None:
    image = Image.new("RGB", (1000, 1000), "#F5F1E8")
    draw = ImageDraw.Draw(image)
    navy = "#142B45"
    teal = "#2E9C9A"
    coral = "#F18A72"
    draw.ellipse((100, 35, 900, 835), fill=navy)
    draw.ellipse((175, 110, 825, 760), outline=teal, width=18)
    draw.text((338, 205), "M", font=font(330, True), fill="white")
    draw.rounded_rectangle((345, 595, 655, 655), radius=30, fill=coral)
    draw.text((220, 875), "MAQUETAS EXCEL", font=font(52, True), fill=navy)
    image.save(OUT / "logo-maquetasexcel.png", format="PNG", optimize=True)


def banner() -> None:
    width, height = 3360, 840
    image = Image.new("RGB", (width, height), "#F5F1E8")
    draw = ImageDraw.Draw(image)
    navy = "#142B45"
    teal = "#2E9C9A"
    coral = "#F18A72"
    ink = "#203449"
    draw.rectangle((0, 0, width, height), fill="#F5F1E8")
    draw.rectangle((0, 0, 1230, height), fill=navy)
    draw.ellipse((2600, -380, 3600, 650), fill="#E6D9C7")
    draw.ellipse((2450, 480, 3500, 1300), fill="#D7ECE8")
    draw.text((150, 145), "MAQUETAS EXCEL", font=font(78, True), fill="#8FE0D4")
    draw.text((150, 300), "Plantillas que ponen orden", font=font(92, True), fill="white")
    draw.text((150, 425), "a tus ideas y proyectos", font=font(92, True), fill="white")
    draw.text((155, 610), "Excel · e-books · herramientas digitales", font=font(45), fill="#DDEBED")
    draw.rounded_rectangle((2020, 140, 3000, 700), radius=48, fill="#203449")
    draw.rounded_rectangle((2090, 205, 2930, 620), radius=28, fill="#EAF4F1")
    draw.rectangle((2125, 245, 2895, 325), fill="#D8EFEB")
    draw.text((2170, 263), "PANEL DE CONTROL", font=font(29, True), fill=navy)
    for row, color in enumerate((teal, coral, "#D5B45B")):
        y = 380 + row * 72
        draw.rounded_rectangle((2170, y, 2290, y + 40), radius=12, fill=color)
        draw.rounded_rectangle((2330, y, 2800, y + 40), radius=12, fill="white")
    draw.text((2185, 710), "Recursos prácticos para empezar hoy", font=font(34, True), fill=ink)
    image.save(OUT / "banner-maquetasexcel.png", format="PNG", optimize=True)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    logo()
    banner()
    print(OUT / "logo-maquetasexcel.png")
    print(OUT / "banner-maquetasexcel.png")


if __name__ == "__main__":
    main()
