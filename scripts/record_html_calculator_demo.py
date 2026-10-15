"""Capture and record the shipped offline HTML calculator in a real browser."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

import imageio.v2 as imageio
import numpy as np
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent.parent
FPS = 20
SIZE = (1920, 1080)
DEMO_VALUES = {
    "moving_budget": [1800, 1120, 12],
    "renovation_contingency": [14500, 15, 1000],
    "garden_spacing": [2.4, 1.2, 30],
    "aquarium_volume": [90, 45, 36, 80],
    "puppy_first_year": [425, 620, 74, 58, 31],
}
LABELS = {
    "en": ("1 / OPEN THE OFFLINE TOOL", "2 / ENTER YOUR VALUES", "3 / REVIEW THE LIVE RESULT"),
    "es": ("1 / ABRE LA HERRAMIENTA OFFLINE", "2 / INTRODUCE TUS DATOS", "3 / REVISA EL RESULTADO"),
}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def browser_path() -> Path:
    available = shutil.which("msedge") or shutil.which("msedge.exe")
    candidates = [Path(available)] if available else []
    candidates.extend(
        Path(os.environ.get(variable, "")) / "Microsoft" / "Edge" / "Application" / "msedge.exe"
        for variable in ("PROGRAMFILES(X86)", "PROGRAMFILES")
        if os.environ.get(variable)
    )
    candidates.extend((
        Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    ))
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise FileNotFoundError("Microsoft Edge is required to capture the HTML calculator demo.")


def _font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = (
        Path("C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    )
    for candidate in candidates:
        try:
            return ImageFont.truetype(str(candidate), size)
        except OSError:
            continue
    return ImageFont.load_default()


def _labelled_frame(path: Path, label: str) -> np.ndarray:
    image = Image.open(path).convert("RGB")
    if image.size != SIZE:
        image = image.resize(SIZE, Image.Resampling.LANCZOS)
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, SIZE[1] - 82, SIZE[0], SIZE[1]), fill="#203A3B")
    draw.rectangle((0, SIZE[1] - 82, 14, SIZE[1]), fill="#CF765B")
    draw.text((45, SIZE[1] - 59), label, font=_font(25, True), fill="white")
    return np.asarray(image)


def record_html_calculator_demo(
    html_path: Path,
    preview_path: Path,
    video_path: Path,
    *,
    operation: str,
    language: str = "en",
) -> dict:
    """Load the delivered HTML in Edge, enter a sample, capture it and record MP4."""
    html_path = html_path.resolve()
    if not html_path.is_file():
        raise FileNotFoundError(f"HTML calculator is missing: {html_path}")
    if operation not in DEMO_VALUES:
        raise ValueError(f"No safe demo values declared for HTML operation: {operation}")
    values = DEMO_VALUES[operation]
    browser = browser_path()
    preview_path.parent.mkdir(parents=True, exist_ok=True)
    video_path.parent.mkdir(parents=True, exist_ok=True)
    language_key = "es" if str(language).casefold().startswith("es") else "en"
    labels = LABELS[language_key]

    with tempfile.TemporaryDirectory(prefix="html-calculator-demo-", dir=preview_path.parent) as temporary:
        temp = Path(temporary)
        shots: list[Path] = []
        for index, stage_values in enumerate((None, values[:-1] + [values[-1] / 2], values)):
            state_page = temp / f"stage-{index + 1}.html"
            source = html_path.read_text(encoding="utf-8")
            if stage_values is not None:
                payload = json.dumps(stage_values, separators=(",", ":"))
                script = (
                    "<script>const demoValues=" + payload + ";"
                    "const demoInputs=[...document.querySelectorAll('#tool-form input')];"
                    "demoInputs.forEach((input,index)=>{input.value=String(demoValues[index]);"
                    "input.dispatchEvent(new Event('input',{bubbles:true}));});"
                    "document.getElementById('tool-form').dispatchEvent(new Event('submit',{bubbles:true,cancelable:true}));"
                    "</script>"
                )
                source = source.replace("</body>", script + "</body>", 1)
            state_page.write_text(source, encoding="utf-8")
            shot = temp / f"stage-{index + 1}.png"
            profile = temp / f"edge-profile-{index + 1}"
            command = [
                str(browser), "--headless=new", "--disable-gpu", "--hide-scrollbars",
                "--no-first-run", "--disable-extensions", "--disable-background-networking",
                f"--user-data-dir={profile}", f"--window-size={SIZE[0]},{SIZE[1]}",
                "--force-device-scale-factor=1", "--virtual-time-budget=1800",
                f"--screenshot={shot}", state_page.as_uri(),
            ]
            result = subprocess.run(command, capture_output=True, text=True, timeout=40, check=False)
            if result.returncode or not shot.is_file():
                details = (result.stderr or result.stdout)[-1800:]
                raise RuntimeError(f"Edge could not capture HTML demo state {index + 1}: {details}")
            shots.append(shot)

        preview_path.write_bytes(shots[-1].read_bytes())
        if Image.open(preview_path).size != SIZE:
            raise RuntimeError("HTML calculator screenshot has an unexpected viewport size")
        with imageio.get_writer(video_path, fps=FPS, codec="libx264", quality=7, pixelformat="yuv420p", macro_block_size=1) as writer:
            for shot, label in zip(shots, labels):
                frame = _labelled_frame(shot, label)
                for _ in range(2 * FPS):
                    writer.append_data(frame)
    if not video_path.is_file() or video_path.stat().st_size < 100_000:
        raise RuntimeError("HTML calculator demo video was not produced correctly")
    return {
        "passed": True,
        "operation": operation,
        "language": language_key,
        "browser": "Microsoft Edge headless",
        "preview_file": preview_path.name,
        "preview_sha256": sha256(preview_path),
        "video_file": video_path.name,
        "video_sha256": sha256(video_path),
        "video_size_bytes": video_path.stat().st_size,
        "video_duration_seconds": 6,
        "scenes": list(labels),
        "offline_page_loaded": True,
        "demo_input_changed": True,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Render a browser preview and video of the shipped HTML calculator.")
    parser.add_argument("--html", type=Path, required=True)
    parser.add_argument("--preview", type=Path, required=True)
    parser.add_argument("--video", type=Path, required=True)
    parser.add_argument("--operation", required=True, choices=sorted(DEMO_VALUES))
    parser.add_argument("--language", default="en")
    args = parser.parse_args()
    report = record_html_calculator_demo(args.html, args.preview, args.video, operation=args.operation, language=args.language)
    print("HTML_CALCULATOR_DEMO=" + json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
