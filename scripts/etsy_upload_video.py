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


PROJECT_ROOT = Path(__file__).resolve().parent.parent
VIDEO_DIR = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "video"
VIDEO_CANDIDATES = (
    VIDEO_DIR / "planificador-demo-interactivo-etsy.mp4",
    VIDEO_DIR / "planificador-demo-etsy.mp4",
)
VIDEO_PATH = next((path for path in VIDEO_CANDIDATES if path.exists()), VIDEO_CANDIDATES[0])


def env_flag(values: dict[str, str], key: str) -> bool:
    return values.get(key, "").strip().casefold() in {"1", "true", "yes", "on"}


def upload(values: dict[str, str], shop_id: str, listing_id: str, token: str) -> dict[str, Any]:
    boundary = f"----CodexEtsyVideo{secrets.token_hex(16)}"
    content_type = mimetypes.guess_type(VIDEO_PATH.name)[0] or "video/mp4"
    body = multipart_body(
        [("name", VIDEO_PATH.name)],
        "video",
        VIDEO_PATH,
        boundary,
        content_type=content_type,
    )
    url = f"https://api.etsy.com/v3/application/shops/{shop_id}/listings/{listing_id}/videos?is_multi_video=true"
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
        "x-api-key": api_key(values),
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    }
    _, raw = request_bytes("POST", url, headers=headers, body=body)
    return json.loads(raw.decode("utf-8")) if raw else {}


def main() -> int:
    parser = argparse.ArgumentParser(description="Sube el vídeo de demostración a un listing Etsy")
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--execute", action="store_true", help="Permite escrituras reales en Etsy")
    parser.add_argument("--allow-additional", action="store_true", help="Sube el vídeo nuevo aunque exista uno activo")
    parser.add_argument("--replace-active", action="store_true", help="Elimina los vídeos activos anteriores después de cargar el nuevo")
    args = parser.parse_args()
    values = load_env()
    manifest_path = args.manifest.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    shop_id = values.get("ETSY_SHOP_ID", "").strip()
    listing_id = str(manifest.get("etsy_listing_id", ""))
    if not shop_id or not listing_id:
        raise SystemExit("Falta ETSY_SHOP_ID o etsy_listing_id")
    if not VIDEO_PATH.exists():
        raise SystemExit(f"No existe el vídeo: {VIDEO_PATH}")
    print(f"Vídeo preparado: {VIDEO_PATH.name} ({VIDEO_PATH.stat().st_size} bytes)")
    if not args.execute:
        print("Modo seguro: no se modifica Etsy.")
        return 0
    if not env_flag(values, "ETSY_PUBLISH_ENABLED") or env_flag(values, "ETSY_DRY_RUN"):
        raise SystemExit(
            "Carga bloqueada por configuración: usa ETSY_DRY_RUN=false y "
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
    existing = request_json("GET", f"/application/listings/{listing_id}/videos", values, oauth_token=token)
    active = [item for item in existing.get("results", []) if item.get("video_state") == "active"]
    if active and not (args.allow_additional or args.replace_active):
        raise SystemExit("El anuncio ya tiene un vídeo activo; no se sustituye automáticamente.")
    if active:
        print(f"Vídeo activo existente detectado: {len(active)}; se añadirá el nuevo vídeo solicitado.")
    result = upload(values, shop_id, listing_id, token)
    new_video_id = int(result.get("video_id")) if result.get("video_id") else None
    if args.replace_active and new_video_id:
        for old_video in active:
            old_video_id = int(old_video["video_id"])
            if old_video_id == new_video_id:
                continue
            request_json(
                "DELETE",
                f"/application/shops/{shop_id}/listings/{listing_id}/videos/{old_video_id}",
                values,
                oauth_token=token,
            )
            print(f"Vídeo anterior eliminado: {old_video_id}")
    manifest["etsy_video_id"] = new_video_id
    manifest["video_uploaded"] = True
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Vídeo subido a Etsy. ID: {result.get('video_id', '<no indicado>')}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except EtsyApiError as error:
        raise SystemExit(str(error)) from error
