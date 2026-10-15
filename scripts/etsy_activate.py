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
from listing_package import ListingDraft, validate_listing_draft
from etsy_upload_video import upload as upload_listing_video


def env_flag(values: dict[str, str], key: str) -> bool:
    return values.get(key, "").strip().casefold() in {"1", "true", "yes", "on"}


def upload_image(values: dict[str, str], shop_id: str, listing_id: str, image_path: Path, token: str, rank: int, alt_text: str) -> dict[str, Any]:
    boundary = f"----CodexEtsyImage{secrets.token_hex(16)}"
    content_type = mimetypes.guess_type(image_path.name)[0] or "image/png"
    body = multipart_body(
        [("rank", str(rank)), ("alt_text", alt_text[:250])],
        "image",
        image_path,
        boundary,
        content_type=content_type,
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
    parser = argparse.ArgumentParser(description="Sube los medios del paquete y activa un listing Etsy")
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--execute", action="store_true", help="Permite escrituras reales en Etsy")
    args = parser.parse_args()

    values = load_env()
    manifest_path = args.manifest.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    listing_id = str(manifest.get("etsy_listing_id", ""))
    shop_id = values.get("ETSY_SHOP_ID", "").strip()
    image_path = manifest_path.parent / str(manifest.get("cover_image", "portada.png"))
    image_paths = [manifest_path.parent / str(item) for item in manifest.get("image_files", [])]
    video_refs = manifest.get("video_files") or [manifest.get("video_file")]
    video_paths = [manifest_path.parent / str(item or "") for item in video_refs if item]
    draft = ListingDraft.model_validate(manifest["draft"])
    validation = validate_listing_draft(draft)
    if not validation.passed:
        raise SystemExit("El listing no pasa la validación: " + "; ".join(validation.errors))
    if not listing_id:
        raise SystemExit("El manifiesto no contiene un Etsy listing_id")
    if manifest.get("etsy_status") == "active":
        raise SystemExit("El listing ya está activo; usa los cargadores de medios específicos para evitar duplicados.")
    if not shop_id:
        raise SystemExit("Falta ETSY_SHOP_ID")
    if not image_path.exists():
        raise SystemExit(f"No existe la portada: {image_path}")
    if not video_paths or any(not path.is_file() for path in video_paths):
        raise SystemExit("No se activa: el paquete no incluye el vídeo demo declarado en el manifiesto")
    if len(manifest.get("image_files", [])) > 10:
        raise SystemExit("No se activa: el manifiesto supera el límite de diez imágenes de Etsy")
    missing_images = [path for path in image_paths if not path.is_file()]
    if missing_images:
        raise SystemExit("Faltan imágenes de marketing: " + ", ".join(str(path) for path in missing_images))
    if not args.execute:
        print("Modo seguro: no se modifica Etsy.")
        print(f"Portada preparada: {image_path}")
        print(f"Imágenes de marketing: {len(image_paths)}; vídeos: {', '.join(path.name for path in video_paths)}")
        print("Usa --execute para subir el paquete visual y activar el listing.")
        return 0
    if not env_flag(values, "ETSY_PUBLISH_ENABLED") or env_flag(values, "ETSY_DRY_RUN"):
        raise SystemExit(
            "Activación bloqueada por configuración: usa ETSY_DRY_RUN=false y "
            "ETSY_PUBLISH_ENABLED=true solo durante esta ejecución."
        )

    token_data = refresh_access_token(values)
    token = str(token_data.get("access_token", ""))
    scopes = set(str(token_data.get("scope", "")).split())
    if "listings_w" not in scopes:
        raise SystemExit(f"El token no tiene listings_w; scopes concedidos: {' '.join(sorted(scopes))}")
    rotated_refresh = str(token_data.get("refresh_token", ""))
    if rotated_refresh:
        update_env_value("ETSY_REFRESH_TOKEN", rotated_refresh)

    draft_title = str(manifest["draft"]["title"])
    image_id_map: dict[str, int] = {}
    uploaded_paths: set[str] = set()
    cover_result = upload_image(values, shop_id, listing_id, image_path, token, 1, f"{draft_title}: cover")
    if cover_result.get("listing_image_id") is not None:
        image_id_map[image_path.name] = int(cover_result["listing_image_id"])
    uploaded_paths.add(image_path.name)
    rank = 2
    for relative in manifest.get("image_files", []):
        path = manifest_path.parent / str(relative)
        if path.name in uploaded_paths:
            continue
        if rank > 10:
            break
        alt_text = f"{draft_title}: {path.stem.replace('-', ' ').replace('_', ' ')}"
        result = upload_image(values, shop_id, listing_id, path, token, rank, alt_text)
        if result.get("listing_image_id") is not None:
            image_id_map[path.relative_to(manifest_path.parent).as_posix()] = int(result["listing_image_id"])
        uploaded_paths.add(path.name)
        rank += 1
    video_results = [upload_listing_video(values, shop_id, listing_id, token, path) for path in video_paths]
    request_json(
        "PATCH",
        f"/application/shops/{shop_id}/listings/{listing_id}",
        values,
        oauth_token=token,
        form=[("state", "active"), ("type", "download")],
    )
    manifest["etsy_status"] = "active"
    manifest["cover_image_uploaded"] = True
    manifest["etsy_image_id"] = cover_result.get("listing_image_id")
    manifest["marketing_image_ids"] = image_id_map
    video_ids = [int(result["video_id"]) for result in video_results if result.get("video_id")]
    manifest["etsy_video_id"] = video_ids[0] if video_ids else None
    manifest["etsy_html_tool_video_id"] = video_ids[1] if len(video_ids) > 1 else None
    manifest["video_uploaded"] = len(video_ids) == len(video_paths)
    manifest["html_tool_video_uploaded"] = len(video_paths) > 1 and len(video_ids) == len(video_paths)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("Portada subida y listing activado en Etsy.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except EtsyApiError as error:
        raise SystemExit(str(error)) from error

