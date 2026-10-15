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
SCREENSHOT_DIR = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "screenshots_recortadas"
MOCKUP_DIR = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "marketing_mockups"


def env_flag(values: dict[str, str], key: str) -> bool:
    return values.get(key, "").strip().casefold() in {"1", "true", "yes", "on"}


def upload_image(
    values: dict[str, str],
    shop_id: str,
    listing_id: str,
    image_path: Path,
    token: str,
    rank: int,
    alt_text: str,
) -> dict[str, Any]:
    boundary = f"----CodexEtsyScreenshot{secrets.token_hex(16)}"
    content_type = mimetypes.guess_type(image_path.name)[0] or "image/png"
    body = multipart_body(
        [("rank", str(rank)), ("alt_text", alt_text)],
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
    parser = argparse.ArgumentParser(description="Añade capturas de una plantilla Excel a un listing Etsy")
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--execute", action="store_true", help="Permite escrituras reales en Etsy")
    parser.add_argument("--replace", action="store_true", help="Elimina las capturas existentes y sube el carrusel actual")
    args = parser.parse_args()

    values = load_env()
    manifest_path = args.manifest.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    listing_id = str(manifest.get("etsy_listing_id", ""))
    shop_id = values.get("ETSY_SHOP_ID", "").strip()
    screenshot_paths = sorted(MOCKUP_DIR.glob("*.png")) + sorted(SCREENSHOT_DIR.glob("*.png"))
    if not listing_id:
        raise SystemExit("El manifiesto no contiene un Etsy listing_id")
    if not shop_id:
        raise SystemExit("Falta ETSY_SHOP_ID")
    if not screenshot_paths:
        raise SystemExit(f"No hay capturas en {SCREENSHOT_DIR}")
    print(f"Assets visuales preparados: {len(screenshot_paths)}")
    if not args.execute:
        print("Modo seguro: no se modifica Etsy.")
        for path in screenshot_paths:
            print(f"- {path.name}")
        print("Usa --execute con publicación habilitada solo para subirlas.")
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

    existing = request_json(
        "GET",
        f"/application/listings/{listing_id}/images",
        values,
        oauth_token=token,
    )
    existing_ids = set(manifest.get("screenshot_image_ids", []))
    existing_alt_texts = {
        str(item.get("alt_text", "")) for item in existing.get("results", [])
    }
    current_count = len(existing.get("results", []))
    results: list[int] = []
    alt_texts = {
        "01-hero-mockup.png": "Mockup del planificador de presupuesto familiar editable en Excel",
        "02-feature-grid-mockup.png": "Funciones principales del planificador de presupuesto familiar",
        "03-que-incluye-mockup.png": "Contenido incluido en la plantilla Excel descargable",
        "04-como-funciona-mockup.png": "Pasos para usar el planificador de presupuesto familiar",
        "01-instrucciones.png": "Vista de la hoja de instrucciones de la plantilla Excel",
        "02-presupuesto.png": "Vista de la hoja de presupuesto mensual en Excel",
        "03-vista-previa.png": "Vista resumen del planificador de presupuesto en Excel",
        "04-configuracion.png": "Vista de la hoja de configuración del hogar en Excel",
        "05-gastos-compartidos.png": "Vista de la hoja de gastos compartidos en Excel",
    }
    if args.replace:
        # Rank 1 is the cover. Replace only the current screenshot carousel.
        old_screenshots = [
            item for item in existing.get("results", []) if int(item.get("rank", 1)) > 1
        ]
        for item in old_screenshots:
            image_id = int(item["listing_image_id"])
            request_json(
                "DELETE",
                f"/application/shops/{shop_id}/listings/{listing_id}/images/{image_id}",
                values,
                oauth_token=token,
            )
            existing_ids.discard(image_id)
            print(f"Eliminada captura anterior: {image_id}")
        current_count = len(existing.get("results", [])) - len(old_screenshots)
        existing_alt_texts = set()

    for path in screenshot_paths:
        if current_count >= 10:
            print("Etsy permite como máximo 10 imágenes; se detiene la carga.")
            break
        alt_text = alt_texts.get(path.name, "Captura de la plantilla Excel")
        if alt_text in existing_alt_texts:
            print(f"Ya existente: {path.name}")
            continue
        result = upload_image(
            values,
            shop_id,
            listing_id,
            path,
            token,
            current_count + 1,
            alt_text,
        )
        image_id = result.get("listing_image_id")
        if image_id is not None:
            results.append(int(image_id))
            existing_ids.add(int(image_id))
        existing_alt_texts.add(alt_text)
        current_count += 1
        print(f"Subida: {path.name}")

    manifest["screenshot_image_ids"] = sorted(existing_ids)
    manifest["screenshots_uploaded"] = len(results)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Capturas nuevas subidas: {len(results)}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except EtsyApiError as error:
        raise SystemExit(str(error)) from error
