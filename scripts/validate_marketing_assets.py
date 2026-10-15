from __future__ import annotations

import json
from pathlib import Path

import imageio.v2 as imageio
from PIL import Image


ROOT = Path(__file__).resolve().parent.parent
PRODUCT = ROOT / "outputs" / "listings" / "product-7"
SPEC = PRODUCT / "marketing_spec.json"
MOCKUPS = PRODUCT / "marketing_mockups"
VIDEO = PRODUCT / "video" / "planificador-demo-interactivo-etsy.mp4"
RAW_VIDEO = PRODUCT / "video" / "planificador-demo-interactivo.mp4"


def main() -> int:
    spec = json.loads(SPEC.read_text(encoding="utf-8"))
    errors: list[str] = []
    for scene in spec["scenes"]:
        path = MOCKUPS / scene["file"]
        if not path.is_file():
            errors.append(f"Falta mockup: {path}")
            continue
        with Image.open(path) as image:
            if image.size != (2000, 1500):
                errors.append(f"Dimensión incorrecta en {path.name}: {image.size}")
    for path in (RAW_VIDEO, VIDEO):
        if not path.is_file() or path.stat().st_size < 100_000:
            errors.append(f"Vídeo ausente o demasiado pequeño: {path}")
    if VIDEO.is_file():
        reader = imageio.get_reader(VIDEO)
        metadata = reader.get_meta_data()
        duration = float(metadata.get("duration", 0))
        size = tuple(metadata.get("size", ()))
        if size != (1920, 960):
            errors.append(f"Dimensión incorrecta en vídeo: {size}")
        if not (5 <= duration <= 15):
            errors.append(f"Duración fuera de Etsy: {duration:.2f}s")
        reader.close()
    if errors:
        print("MARKETING_QA=FAILED")
        for error in errors:
            print(error)
        return 1
    print("MARKETING_QA=APPROVED")
    print(f"Mockups={len(spec['scenes'])}")
    print(f"Video={VIDEO.name}")
    print("VideoContract=PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
