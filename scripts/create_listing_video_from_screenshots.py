from __future__ import annotations

from pathlib import Path

import imageio.v2 as imageio
import numpy as np
from PIL import Image, ImageOps


PROJECT_ROOT = Path(__file__).resolve().parent.parent
PRODUCT_DIR = PROJECT_ROOT / "outputs" / "listings" / "product-7"
SCREENSHOTS = PRODUCT_DIR / "screenshots"
OUTPUT = PRODUCT_DIR / "video" / "planificador-demo-etsy.mp4"
TARGET = (1920, 960)  # 2:1, recomendado por Etsy para listing videos.
FPS = 24


def frame(path: Path) -> np.ndarray:
    image = Image.open(path).convert("RGB")
    image = ImageOps.fit(image, TARGET, method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))
    return np.asarray(image)


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    cover = frame(PRODUCT_DIR / "portada.png")
    screenshots = [frame(path) for path in sorted(SCREENSHOTS.glob("*.png"))]
    with imageio.get_writer(
        OUTPUT,
        fps=FPS,
        codec="libx264",
        quality=7,
        pixelformat="yuv420p",
        macro_block_size=1,
    ) as writer:
        for image, seconds in [(cover, 1.2), *[(shot, 1.8) for shot in screenshots]]:
            for _ in range(round(seconds * FPS)):
                writer.append_data(image)
    print(OUTPUT)


if __name__ == "__main__":
    main()
