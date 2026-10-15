from __future__ import annotations

import argparse
import json
import mimetypes
import secrets
from pathlib import Path
from typing import Any

from etsy_client import (
    EtsyApiError,
    api_key,
    load_env,
    multipart_body,
    refresh_access_token,
    request_bytes,
    request_json,
    update_env_value,
)


def env_flag(values: dict[str, str], key: str) -> bool:
    return values.get(key, "").strip().casefold() in {"1", "true", "yes", "on"}


def upload_image(
    values: dict[str, str], shop_id: str, listing_id: str, image_path: Path,
    token: str, rank: int, alt_text: str, *, overwrite: bool = False,
) -> dict[str, Any]:
    boundary = f"----CodexEtsyImage{secrets.token_hex(16)}"
    content_type = mimetypes.guess_type(image_path.name)[0] or "image/png"
    fields = [("rank", str(rank)), ("alt_text", alt_text[:250])]
    if overwrite:
        fields.append(("overwrite", "true"))
    body = multipart_body(
        fields,
        "image", image_path, boundary, content_type=content_type,
    )
    url = f"https://api.etsy.com/v3/application/shops/{shop_id}/listings/{listing_id}/images"
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
        "x-api-key": api_key(values),
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    }
    _, raw = request_bytes("POST", url, headers=headers, body=body)
    return json.loads(raw.decode("utf-8")) if raw else {}


def main() -> int:
    parser = argparse.ArgumentParser(description="Sube las imágenes de marketing de un paquete de listing Etsy.")
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--execute", action="store_true", help="Permite escrituras reales en Etsy")
    parser.add_argument("--replace", action="store_true", help="Elimina el carrusel actual, excepto la portada, antes de subirlo")
    parser.add_argument("--replace-all", action="store_true", help="Reemplaza portada y carrusel completos por las imágenes del manifiesto")
    args = parser.parse_args()
    if args.replace and args.replace_all:
        parser.error("Usa --replace o --replace-all, no ambos")

    values = load_env()
    manifest_path = args.manifest.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    listing_id = str(manifest.get("etsy_listing_id", ""))
    shop_id = values.get("ETSY_SHOP_ID", "").strip()
    cover = str(manifest.get("cover_image") or "")
    package_dir = manifest_path.parent
    image_paths = [
        package_dir / str(item)
        for item in manifest.get("image_files", [])
        if args.replace_all or str(item) != cover
    ]
    if not listing_id or not shop_id:
        raise SystemExit("Falta ETSY_SHOP_ID o etsy_listing_id en el manifiesto")
    missing = [path for path in image_paths if not path.is_file()]
    if missing:
        raise SystemExit("Faltan imágenes del manifiesto: " + ", ".join(str(path) for path in missing))
    if not image_paths:
        raise SystemExit("El manifiesto no contiene imágenes de marketing adicionales a la portada")
    print(f"Imágenes de marketing preparadas: {len(image_paths)}")
    if not args.execute:
        print("Modo seguro: no se modifica Etsy.")
        for path in image_paths:
            print(f"- {path.name}")
        return 0
    if not env_flag(values, "ETSY_PUBLISH_ENABLED") or env_flag(values, "ETSY_DRY_RUN"):
        raise SystemExit("Carga bloqueada por configuración de seguridad de Etsy")

    token_data = refresh_access_token(values)
    token = str(token_data.get("access_token", ""))
    if "listings_w" not in set(str(token_data.get("scope", "")).split()):
        raise SystemExit("El token OAuth no incluye listings_w")
    rotated_refresh = str(token_data.get("refresh_token", ""))
    if rotated_refresh:
        update_env_value("ETSY_REFRESH_TOKEN", rotated_refresh)

    existing = request_json("GET", f"/application/listings/{listing_id}/images", values, oauth_token=token)
    existing_ids: set[int] = set()
    current_images = list(existing.get("results", []))
    if args.replace_all:
        cover_path = package_dir / cover
        # Replace rank 1 first. Etsy documents `overwrite=true` for this endpoint,
        # so a failed cover upload leaves the old primary image intact.
        title = str((manifest.get("draft") or {}).get("title", "Excel workbook"))
        cover_result = upload_image(values, shop_id, listing_id, cover_path, token, 1, f"{title}: product cover", overwrite=True)
        cover_id = cover_result.get("listing_image_id")
        if cover_id is None:
            raise SystemExit("Etsy no confirmó la nueva portada; no se eliminó ninguna imagen anterior.")
        cover_id = int(cover_id)
        print(f"Subida: {cover_path.name}")
        after_cover = request_json("GET", f"/application/listings/{listing_id}/images", values, oauth_token=token)
        for item in after_cover.get("results", []):
            image_id = int(item["listing_image_id"])
            if image_id != cover_id:
                request_json("DELETE", f"/application/shops/{shop_id}/listings/{listing_id}/images/{image_id}", values, oauth_token=token)
        existing_ids.add(cover_id)
        current_images = [{"listing_image_id": cover_id, "rank": 1, "alt_text": f"{title}: product cover"}]
        image_paths = [path for path in image_paths if path.resolve() != cover_path.resolve()]
    elif args.replace:
        for item in current_images:
            if int(item.get("rank", 1)) <= 1:
                continue
            image_id = int(item["listing_image_id"])
            request_json("DELETE", f"/application/shops/{shop_id}/listings/{listing_id}/images/{image_id}", values, oauth_token=token)
            existing_ids.discard(image_id)
        current_images = [item for item in current_images if int(item.get("rank", 1)) <= 1]

    known_alt_texts = {str(item.get("alt_text", "")) for item in current_images}
    current_count = len(current_images)
    title = str((manifest.get("draft") or {}).get("title", "Excel workbook"))
    uploaded: list[int] = []
    uploaded_by_path: dict[str, int] = {}
    for rank, path in enumerate(image_paths, start=current_count + 1):
        if current_count >= 10:
            print("Se alcanzó el límite de diez imágenes de Etsy.")
            break
        is_cover = str(path.relative_to(package_dir).as_posix()) == cover
        alt_text = f"{title}: {'product cover' if is_cover else path.stem.replace('-', ' ').replace('_', ' ')}"[:250]
        if alt_text in known_alt_texts:
            continue
        result = upload_image(values, shop_id, listing_id, path, token, rank, alt_text)
        image_id = result.get("listing_image_id")
        if image_id is not None:
            uploaded.append(int(image_id))
            existing_ids.add(int(image_id))
            uploaded_by_path[str(path.relative_to(package_dir).as_posix())] = int(image_id)
        known_alt_texts.add(alt_text)
        current_count += 1
        print(f"Subida: {path.name}")
    manifest["screenshot_image_ids"] = sorted(existing_ids)
    manifest["screenshots_uploaded"] = len(uploaded)
    if args.replace_all:
        manifest["etsy_image_id"] = existing_ids and int(current_images[0]["listing_image_id"])
        manifest["cover_image_uploaded"] = bool(existing_ids)
        manifest["marketing_image_ids"] = {cover: int(current_images[0]["listing_image_id"]), **uploaded_by_path}
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Imágenes nuevas subidas: {len(uploaded)}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except EtsyApiError as error:
        raise SystemExit(str(error)) from error
