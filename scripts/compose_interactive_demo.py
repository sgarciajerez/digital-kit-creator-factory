from __future__ import annotations

from pathlib import Path

import imageio.v2 as imageio
import numpy as np
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent.parent
PRODUCT = ROOT / "outputs" / "listings" / "product-7"
SOURCE = PRODUCT / "video" / "planificador-demo-interactivo.mp4"
OUT = PRODUCT / "video" / "planificador-demo-interactivo-etsy.mp4"
HERO = PRODUCT / "marketing_mockups" / "01-hero-mockup.png"
END = PRODUCT / "marketing_mockups" / "04-como-funciona-mockup.png"
SIZE = (1920, 960)
FPS = 24
CREAM = "#F5F1E8"


def fit(path: Path) -> np.ndarray:
    source = Image.open(path).convert("RGB")
    source.thumbnail((1740, 960), Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", SIZE, CREAM)
    canvas.paste(source, ((SIZE[0] - source.width) // 2, (SIZE[1] - source.height) // 2))
    return np.asarray(canvas)


def blend(a: np.ndarray, b: np.ndarray, amount: float) -> np.ndarray:
    return np.asarray((1.0 - amount) * a.astype(np.float32) + amount * b.astype(np.float32), dtype=np.uint8)


def main() -> None:
    intro = fit(HERO)
    outro = fit(END)
    reader = imageio.get_reader(SOURCE)
    frames = [reader.get_data(i) for i in range(reader.count_frames())]
    with imageio.get_writer(OUT, fps=FPS, codec="libx264", quality=7, pixelformat="yuv420p", macro_block_size=1) as writer:
        for _ in range(int(1.2 * FPS)):
            writer.append_data(intro)
        for index in range(int(0.6 * FPS)):
            writer.append_data(blend(intro, frames[0], (index + 1) / (0.6 * FPS)))
        for frame in frames:
            writer.append_data(frame)
        for index in range(int(0.6 * FPS)):
            writer.append_data(blend(frames[-1], outro, (index + 1) / (0.6 * FPS)))
        for _ in range(int(1.2 * FPS)):
            writer.append_data(outro)
    print(OUT)


if __name__ == "__main__":
    main()
