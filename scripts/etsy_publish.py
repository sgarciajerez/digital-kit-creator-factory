from __future__ import annotations

import argparse
import json
import mimetypes
import secrets
from pathlib import Path
from typing import Any

from etsy_client import (
    ENV_PATH,
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


def bool_env(values: dict[str, str], key: str) -> bool:
    return values.get(key, "").strip().casefold() in {"1", "true", "yes", "on"}


def form_fields(manifest: dict[str, Any], taxonomy_id: str) -> list[tuple[str, str]]:
    draft = manifest["draft"]
    fields = [
        ("quantity", "999"),
        ("title", str(draft["title"])),
        ("description", str(draft["description"])),
        ("price", f"{float(draft['price']):.2f}"),
        ("who_made", "i_did"),
        ("when_made", "2020_2026"),
        ("taxonomy_id", taxonomy_id),
        ("type", "download"),
    ]
    fields.extend(("tags", str(tag)) for tag in draft.get("tags", []))
    return fields


def upload_file(
    values: dict[str, str], shop_id: str, listing_id: str, file_path: Path, token: str
) -> dict[str, Any]:
    boundary = f"----CodexEtsy{secrets.token_hex(16)}"
    body = multipart_body(
        [("name", file_path.name), ("rank", "1")], "file", file_path, boundary
    )
    url = f"https://api.etsy.com/v3/application/shops/{shop_id}/listings/{listing_id}/files"
    content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
        "x-api-key": api_key(values),
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    }
    _, raw = request_bytes("POST", url, headers=headers, body=body)
    return json.loads(raw.decode("utf-8")) if raw else {"content_type": content_type}


def update_manifest(path: Path, **changes: Any) -> None:
    manifest = json.loads(path.read_text(encoding="utf-8"))
    manifest.update(changes)
    path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Crea y carga un listing digital en Etsy")
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--execute", action="store_true", help="Permite escrituras reales en Etsy")
    parser.add_argument("--activate", action="store_true", help="Deprecated; activa con etsy_activate.py para cargar todos los medios")
    args = parser.parse_args()

    values = load_env()
    if args.activate:
        raise SystemExit("La activación directa está deshabilitada; usa etsy_activate.py para subir portada, imágenes y vídeo antes de publicar.")
    manifest_path = args.manifest.resolve()
    if not manifest_path.exists():
        raise SystemExit(f"No existe el manifiesto: {manifest_path}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    draft = ListingDraft.model_validate(manifest["draft"])
    validation = validate_listing_draft(draft)
    if not validation.passed:
        raise SystemExit("El listing no pasa la validación: " + "; ".join(validation.errors))

    shop_id = values.get("ETSY_SHOP_ID", "").strip()
    taxonomy_id = values.get("ETSY_TAXONOMY_ID", "").strip()
    local_file = manifest_path.parent / str(manifest["local_file"])
    if not shop_id:
        raise SystemExit("Falta ETSY_SHOP_ID")
    if not taxonomy_id:
        raise SystemExit("Falta ETSY_TAXONOMY_ID; ejecuta etsy_taxonomy.py primero")
    if not local_file.exists():
        raise SystemExit(f"No existe el archivo del producto: {local_file}")
    if local_file.stat().st_size > 20 * 1024 * 1024:
        raise SystemExit("El ZIP supera el límite de 20 MB por archivo digital de Etsy")

    if args.activate:
        images = [
            child
            for child in manifest_path.parent.iterdir()
            if child.suffix.casefold() in {".png", ".jpg", ".jpeg", ".webp"}
        ]
        if not images:
            raise SystemExit("No se activa: Etsy requiere una imagen de listing y el paquete no contiene ninguna")

    marketing = manifest.get("marketing_assets") or {}
    if marketing.get("status") != "ready":
        raise SystemExit("No se crea el listing: falta completar MARKETING_ASSETS")
    image_files = [manifest_path.parent / str(item) for item in manifest.get("image_files", [])]
    video_refs = manifest.get("video_files") or [manifest.get("video_file")]
    video_paths = [manifest_path.parent / str(item or "") for item in video_refs if item]
    missing_media = [path for path in [*image_files, *video_paths] if not path.is_file()]
    if missing_media:
        raise SystemExit("No se crea el listing: faltan recursos de marketing: " + ", ".join(str(path) for path in missing_media))
    if len(image_files) > 10:
        raise SystemExit("El listing declara mÃ¡s de diez imÃ¡genes, por encima del lÃ­mite de Etsy")

    dry_run = bool_env(values, "ETSY_DRY_RUN")
    publish_enabled = bool_env(values, "ETSY_PUBLISH_ENABLED")
    if not args.execute:
        print("Modo seguro: no se hará ninguna escritura en Etsy.")
        print(f"Producto: {local_file.name}")
        print(f"Título: {draft.title}")
        print(f"Taxonomía configurada: {taxonomy_id}")
        print("Para crear el borrador tras OAuth: añade --execute y activa publicación en .env.")
        return 0
    if dry_run or not publish_enabled:
        raise SystemExit(
            "Publicación bloqueada por configuración: establece ETSY_DRY_RUN=false y "
            "ETSY_PUBLISH_ENABLED=true después de completar OAuth y la revisión humana."
        )
    if args.activate and not publish_enabled:
        raise SystemExit("La activación requiere ETSY_PUBLISH_ENABLED=true")

    token_data = refresh_access_token(values)
    token = str(token_data.get("access_token", ""))
    scopes = set(str(token_data.get("scope", "")).split())
    if "listings_w" not in scopes:
        raise SystemExit(f"El token no tiene listings_w; scopes concedidos: {' '.join(sorted(scopes))}")
    rotated_refresh = str(token_data.get("refresh_token", ""))
    if rotated_refresh:
        update_env_value("ETSY_REFRESH_TOKEN", rotated_refresh)

    listing = request_json(
        "POST",
        f"/application/shops/{shop_id}/listings",
        values,
        oauth_token=token,
        form=form_fields(manifest, taxonomy_id),
    )
    listing_id = str(listing.get("listing_id", ""))
    if not listing_id:
        raise SystemExit("Etsy no devolvió listing_id")
    upload_file(values, shop_id, listing_id, local_file, token)
    status = "draft_uploaded"
    if args.activate:
        request_json(
            "PATCH",
            f"/application/shops/{shop_id}/listings/{listing_id}",
            values,
            oauth_token=token,
            form=[("state", "active"), ("type", "download")],
        )
        status = "active"
    update_manifest(manifest_path, etsy_listing_id=int(listing_id), etsy_status=status)
    print(f"Listing Etsy creado: {listing_id} (estado: {status})")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except EtsyApiError as error:
        raise SystemExit(str(error)) from error

