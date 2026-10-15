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


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_FILE = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "autopilot_3_Planificador_de_presupuesto_familiar_0.1.0.xlsx"
FINAL_REVIEW = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "final_visual_review.json"


def env_flag(values: dict[str, str], key: str) -> bool:
    return values.get(key, "").strip().casefold() in {"1", "true", "yes", "on"}


def upload_file(values: dict[str, str], shop_id: str, listing_id: str, file_path: Path, token: str) -> dict[str, Any]:
    boundary = f"----CodexEtsyFile{secrets.token_hex(16)}"
    content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    body = multipart_body(
        [("name", "Planificador_de_presupuesto_familiar.xlsx"), ("rank", "1")],
        "file",
        file_path,
        boundary,
        content_type=content_type,
    )
    url = f"https://api.etsy.com/v3/application/shops/{shop_id}/listings/{listing_id}/files"
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
        "x-api-key": api_key(values),
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    }
    _, raw = request_bytes("POST", url, headers=headers, body=body)
    return json.loads(raw.decode("utf-8")) if raw else {}


def delete_file(values: dict[str, str], shop_id: str, listing_id: str, file_id: int, token: str) -> None:
    url = f"https://api.etsy.com/v3/application/shops/{shop_id}/listings/{listing_id}/files/{file_id}"
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
        "x-api-key": api_key(values),
    }
    request_bytes("DELETE", url, headers=headers)


def main() -> int:
    parser = argparse.ArgumentParser(description="Sustituye el archivo digital de un anuncio Etsy")
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--file", type=Path, default=DEFAULT_FILE)
    parser.add_argument("--execute", action="store_true")
    args = parser.parse_args()

    values = load_env()
    manifest_path = args.manifest.resolve()
    file_path = args.file.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    listing_id = str(manifest.get("etsy_listing_id", "")).strip()
    shop_id = values.get("ETSY_SHOP_ID", "").strip()
    if not listing_id:
        raise SystemExit("El manifiesto no contiene etsy_listing_id")
    if not shop_id:
        raise SystemExit("Falta ETSY_SHOP_ID")
    if not file_path.exists():
        raise SystemExit(f"No existe el archivo: {file_path}")

    print(f"Archivo preparado: {file_path.name} ({file_path.stat().st_size} bytes)")
    if not args.execute:
        print("Modo seguro: no se modifica Etsy.")
        return 0
    if not env_flag(values, "ETSY_PUBLISH_ENABLED") or env_flag(values, "ETSY_DRY_RUN"):
        raise SystemExit(
            "Carga bloqueada por configuración: usa ETSY_DRY_RUN=false y "
            "ETSY_PUBLISH_ENABLED=true solo durante esta ejecución."
        )
    if FINAL_REVIEW.exists():
        final_review = json.loads(FINAL_REVIEW.read_text(encoding="utf-8"))
        if final_review.get("status") != "approved":
            raise SystemExit(
                "Publicación bloqueada: final_visual_review.json no está aprobado. "
                "Las capturas finales deben superar la revisión visual mínima del 99%."
            )

    token_data = refresh_access_token(values)
    token = str(token_data.get("access_token", ""))
    scopes = set(str(token_data.get("scope", "")).split())
    if "listings_w" not in scopes:
        raise SystemExit(f"El token no tiene listings_w; scopes concedidos: {' '.join(sorted(scopes))}")
    rotated_refresh = str(token_data.get("refresh_token", ""))
    if rotated_refresh:
        update_env_value("ETSY_REFRESH_TOKEN", rotated_refresh)

    files_before = request_json(
        "GET",
        f"/application/shops/{shop_id}/listings/{listing_id}/files",
        values,
        oauth_token=token,
    )
    old_ids = [int(item["listing_file_id"]) for item in files_before.get("results", []) if item.get("listing_file_id")]
    print(f"Archivos actuales detectados: {len(old_ids)}")

    uploaded = upload_file(values, shop_id, listing_id, file_path, token)
    new_id = uploaded.get("listing_file_id")
    if not new_id:
        raise SystemExit(f"Etsy no devolvió listing_file_id tras la carga: {uploaded}")
    new_id = int(new_id)
    print("Archivo nuevo cargado y confirmado.")

    deleted = 0
    for old_id in old_ids:
        if old_id == new_id:
            continue
        delete_file(values, shop_id, listing_id, old_id, token)
        deleted += 1
    print(f"Archivos antiguos eliminados: {deleted}")

    files_after = request_json(
        "GET",
        f"/application/shops/{shop_id}/listings/{listing_id}/files",
        values,
        oauth_token=token,
    )
    final_ids = [int(item["listing_file_id"]) for item in files_after.get("results", []) if item.get("listing_file_id")]
    if new_id not in final_ids:
        raise SystemExit(f"La verificación final no encuentra el archivo nuevo. IDs: {final_ids}")
    manifest["etsy_file_id"] = new_id
    manifest["etsy_file_ids"] = final_ids
    manifest["sha256"] = __import__("hashlib").sha256(file_path.read_bytes()).hexdigest()
    manifest["file_replaced_at"] = __import__("datetime").datetime.now().astimezone().isoformat(timespec="seconds")
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Verificación correcta: {len(final_ids)} archivo(s) digital(es) activo(s).")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except EtsyApiError as error:
        raise SystemExit(str(error)) from error

