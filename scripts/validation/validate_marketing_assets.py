from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path

import imageio.v2 as imageio
from PIL import Image


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def validate(manifest_path: Path) -> list[str]:
    manifest_path = manifest_path.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    package = manifest_path.parent
    errors: list[str] = []
    record = manifest.get("marketing_assets") or {}
    if record.get("status") != "ready":
        errors.append("Manifest does not report marketing_assets=ready")
        return errors
    workbook = package / str(record.get("source_workbook") or manifest.get("workbook_file") or "")
    if not workbook.is_file() or file_hash(workbook) != record.get("source_sha256"):
        errors.append("Marketing source workbook is missing or its SHA-256 changed")
    images = [str(path) for path in manifest.get("image_files", [])]
    if not images or len(images) > 10:
        errors.append(f"Etsy image count must be 1..10, found {len(images)}")
    if images and manifest.get("cover_image") != images[0]:
        errors.append("The kit cover must be the first Etsy gallery image")
    if int(record.get("image_count", -1)) != len(images):
        errors.append("Marketing manifest image_count does not match image_files")
    identity = manifest.get("marketing_identity") or {}
    if not str(identity.get("name", "")).strip() or not str(identity.get("motif", "")).strip():
        errors.append("Marketing package has no named, product-specific visual identity")
    artwork_ref = str(identity.get("artwork_file") or "").strip()
    if artwork_ref:
        artwork = package / artwork_ref
        if not artwork.is_file():
            errors.append("Product-specific marketing artwork is missing")
        elif file_hash(artwork) != identity.get("artwork_sha256"):
            errors.append("Product-specific marketing artwork SHA-256 does not match its identity record")
    for relative in images:
        path = package / relative
        if not path.is_file():
            errors.append(f"Missing listing image: {relative}")
    if manifest.get("cover_image") not in images:
        errors.append("Cover image must be listed in image_files")
    screenshot_files = [str(path) for path in manifest.get("screenshot_files", [])]
    mockup_files = [str(path) for path in manifest.get("mockup_files", [])]
    if not screenshot_files:
        errors.append("No workbook screenshots were declared")
    if len(mockup_files) < 3:
        errors.append("At least three marketing mockups are required")
    mockup_names = {Path(name).name for name in mockup_files}
    if "00-product-cover.png" not in mockup_names or manifest.get("cover_image") not in mockup_files:
        errors.append("The listing cover must be the product-led marketing cover")
    if "05-excel-templates-mockup.png" not in mockup_names:
        errors.append("The marketing set must include a dedicated workbook-editions mockup")
    if manifest.get("ebook_file") and "06-pdf-resources-mockup.png" not in mockup_names:
        errors.append("The marketing set must show the included e-book and PDF resources")
    if manifest.get("ebook_file"):
        capture_ref = str(manifest.get("ebook_cover_capture_file") or "").strip()
        capture = package / capture_ref if capture_ref else None
        if not capture or not capture.is_file():
            errors.append("The actual first page of the e-book must be captured for its marketing mockup")
        elif file_hash(capture) != manifest.get("ebook_cover_capture_sha256"):
            errors.append("E-book first-page capture SHA-256 does not match its manifest")
        elif record.get("ebook_cover_capture_sha256") != manifest.get("ebook_cover_capture_sha256"):
            errors.append("Marketing record does not match the e-book first-page capture")
    html_tool = manifest.get("interactive_html_tool") or {}
    html_demo = record.get("html_tool_demo") or {}
    if html_tool.get("passed"):
        preview_ref = str(manifest.get("html_tool_preview_file") or "").strip()
        preview = package / preview_ref if preview_ref else None
        if not preview or not preview.is_file():
            errors.append("An included HTML calculator must have a captured browser preview")
        elif file_hash(preview) != html_demo.get("preview_sha256"):
            errors.append("HTML calculator preview SHA-256 does not match the marketing record")
        if "07-interactive-html-calculator-mockup.png" not in mockup_names:
            errors.append("An included HTML calculator must have its own prominent marketing mockup")
        html_mockup = next((name for name in manifest.get("mockup_files", []) if Path(name).name == "07-interactive-html-calculator-mockup.png"), None)
        inventory_mockup = next((name for name in manifest.get("mockup_files", []) if Path(name).name == "03-whats-included-mockup.png"), None)
        if len(images) < 3 or images[1:3] != [html_mockup, inventory_mockup]:
            errors.append("When HTML is included, the gallery order must be cover, calculator mockup, inventory mockup")
        html_video_ref = str(manifest.get("html_tool_video_file") or "").strip()
        html_video = package / html_video_ref if html_video_ref else None
        if html_video_ref not in manifest.get("video_files", []):
            errors.append("The HTML calculator demo must be declared in video_files for Etsy upload")
        if not html_video or not html_video.is_file() or html_video.stat().st_size < 100_000:
            errors.append("An included HTML calculator must have a usable demonstration video")
        else:
            if file_hash(html_video) != html_demo.get("video_sha256"):
                errors.append("HTML calculator video SHA-256 does not match the marketing record")
            reader = imageio.get_reader(html_video)
            try:
                metadata = reader.get_meta_data()
                duration = float(metadata.get("duration", 0))
                size = tuple(metadata.get("size", ()))
                if size != (1920, 1080):
                    errors.append(f"HTML calculator video dimensions are not 1920x1080: {size}")
                if not (5 <= duration <= 15):
                    errors.append(f"HTML calculator video must be between 5 and 15 seconds: {duration:.2f}s")
            finally:
                reader.close()
    for relative in screenshot_files:
        path = package / relative
        if not path.is_file():
            errors.append(f"Missing workbook screenshot: {relative}")
            continue
        with Image.open(path) as image:
            if image.size != (1920, 1080):
                errors.append(f"Workbook screenshot has unexpected dimensions: {relative} {image.size}")
    for relative in mockup_files:
        path = package / relative
        if not path.is_file():
            errors.append(f"Missing marketing mockup: {relative}")
            continue
        with Image.open(path) as image:
            if image.size != (2000, 1500):
                errors.append(f"Mockup has unexpected dimensions: {relative} {image.size}")
    video = package / str(manifest.get("video_file") or "")
    if not video.is_file() or video.stat().st_size < 100_000:
        errors.append("Marketing demo video is missing or too small")
    else:
        if record.get("video_sha256") != file_hash(video):
            errors.append("Marketing demo video SHA-256 does not match its manifest")
        reader = imageio.get_reader(video)
        try:
            metadata = reader.get_meta_data()
            duration = float(metadata.get("duration", 0))
            size = tuple(metadata.get("size", ()))
            if size != (1920, 1080):
                errors.append(f"Demo video dimensions are not 1920x1080: {size}")
            if not (5 <= duration <= 15):
                errors.append(f"Demo video must be between 5 and 15 seconds: {duration:.2f}s")
        finally:
            reader.close()
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="Valida imágenes, mockups y vídeo del paquete de un producto.")
    parser.add_argument("--manifest", type=Path, required=True)
    args = parser.parse_args()
    errors = validate(args.manifest)
    if errors:
        print("MARKETING_QA=FAILED")
        for error in errors:
            print(f"- {error}")
        return 1
    print("MARKETING_QA=PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
