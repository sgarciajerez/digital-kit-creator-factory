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


def env_flag(values: dict[str, str], key: str) -> bool:
    return values.get(key, "").strip().casefold() in {"1", "true", "yes", "on"}


def upload_image(values: dict[str, str], shop_id: str, listing_id: str, image_path: Path, token: str) -> dict[str, Any]:
    boundary = f"----CodexEtsyImage{secrets.token_hex(16)}"
    content_type = mimetypes.guess_type(image_path.name)[0] or "image/png"
    body = multipart_body(
        [("rank", "1"), ("alt_text", "Portada del planificador de presupuesto familiar")],
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
    parser = argparse.ArgumentParser(description="Añade portada y activa un listing Etsy existente")
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--execute", action="store_true", help="Permite escrituras reales en Etsy")
    args = parser.parse_args()

    values = load_env()
    manifest_path = args.manifest.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    listing_id = str(manifest.get("etsy_listing_id", ""))
    shop_id = values.get("ETSY_SHOP_ID", "").strip()
    image_path = manifest_path.parent / str(manifest.get("cover_image", "portada.png"))
    draft = ListingDraft.model_validate(manifest["draft"])
    validation = validate_listing_draft(draft)
    if not validation.passed:
        raise SystemExit("El listing no pasa la validación: " + "; ".join(validation.errors))
    if not listing_id:
        raise SystemExit("El manifiesto no contiene un Etsy listing_id")
    if not shop_id:
        raise SystemExit("Falta ETSY_SHOP_ID")
    if not image_path.exists():
        raise SystemExit(f"No existe la portada: {image_path}")
    if not args.execute:
        print("Modo seguro: no se modifica Etsy.")
        print(f"Portada preparada: {image_path}")
        print("Usa --execute para subirla y activar el listing existente.")
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

    image_result = upload_image(values, shop_id, listing_id, image_path, token)
    request_json(
        "PATCH",
        f"/application/shops/{shop_id}/listings/{listing_id}",
        values,
        oauth_token=token,
        form=[("state", "active"), ("type", "download")],
    )
    manifest["etsy_status"] = "active"
    manifest["cover_image_uploaded"] = True
    manifest["etsy_image_id"] = image_result.get("listing_image_id")
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("Portada subida y listing activado en Etsy.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except EtsyApiError as error:
        raise SystemExit(str(error)) from error

