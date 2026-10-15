from __future__ import annotations

from pathlib import Path

import imageio.v2 as imageio
import numpy as np
from PIL import Image, ImageDraw, ImageFilter


ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "outputs" / "listings" / "product-7" / "screenshots_recortadas"
OUTPUT = ROOT / "outputs" / "listings" / "product-7" / "video" / "planificador-demo-etsy.mp4"
TARGET = (1920, 960)
FPS = 24
SLIDE_FRAMES = 52  # about 2.2 s of stable reading time per slide
FADE_FRAMES = 24   # 1.0 s crossfade with eased timing
BACKGROUND = "#F5F1E8"
NAVY = "#142B45"


def slide(image_path: Path, progress: float) -> np.ndarray:
    source = Image.open(image_path).convert("RGB")
    canvas = Image.new("RGB", TARGET, BACKGROUND)
    margin_x, margin_y = 80, 55
    available = (TARGET[0] - 2 * margin_x, TARGET[1] - 2 * margin_y)
    scale = min(available[0] / source.width, available[1] / source.height)
    scale *= 0.97 + 0.03 * progress
    size = (max(1, int(source.width * scale)), max(1, int(source.height * scale)))
    image = source.resize(size, Image.Resampling.LANCZOS)
    x = (TARGET[0] - size[0]) // 2
    y = (TARGET[1] - size[1]) // 2
    shadow = Image.new("RGBA", TARGET, (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle((x + 8, y + 10, x + size[0] + 8, y + size[1] + 10), radius=12, fill=(20, 43, 69, 45))
    shadow = shadow.filter(ImageFilter.GaussianBlur(8))
    canvas = Image.alpha_composite(canvas.convert("RGBA"), shadow)
    frame = Image.new("RGBA", (size[0] + 6, size[1] + 6), (255, 255, 255, 255))
    frame.paste(image, (3, 3))
    canvas.alpha_composite(frame, (x - 3, y - 3))
    return np.asarray(canvas.convert("RGB"))


def eased(progress: float) -> float:
    """Smoothstep easing: slower at the start and end of each transition."""
    return progress * progress * (3.0 - 2.0 * progress)


def main() -> int:
    files = sorted(SOURCE.glob("*.png"))
    if not files:
        raise SystemExit(f"No hay capturas recortadas en {SOURCE}")
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with imageio.get_writer(
        OUTPUT,
        fps=FPS,
        codec="libx264",
        quality=7,
        pixelformat="yuv420p",
        macro_block_size=1,
    ) as writer:
        for index, image_path in enumerate(files):
            start = slide(image_path, 0.0)
            end = slide(image_path, 1.0)
            for frame_index in range(SLIDE_FRAMES):
                progress = frame_index / max(1, SLIDE_FRAMES - 1)
                blend = Image.blend(Image.fromarray(start), Image.fromarray(end), progress)
                writer.append_data(np.asarray(blend))
            if index < len(files) - 1:
                next_start = slide(files[index + 1], 0.0)
                for frame_index in range(FADE_FRAMES):
                    progress = eased((frame_index + 1) / FADE_FRAMES)
                    blend = Image.blend(Image.fromarray(end), Image.fromarray(next_start), progress)
                    writer.append_data(np.asarray(blend))
    print(OUTPUT)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
