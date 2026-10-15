from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from etsy_client import (
    EtsyApiError,
    load_env,
    refresh_access_token,
    request_json,
    update_env_value,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config" / "etsy_shop_config.json"


def flag(values: dict[str, str], key: str) -> bool:
    return values.get(key, "").strip().casefold() in {"1", "true", "yes", "on"}


def main() -> int:
    parser = argparse.ArgumentParser(description="Configura la identidad y secciones de una tienda Etsy")
    parser.add_argument("--execute", action="store_true", help="Aplica cambios reales en Etsy")
    args = parser.parse_args()
    values = load_env()
    config: dict[str, Any] = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    shop_id = values.get("ETSY_SHOP_ID", "").strip()
    if not shop_id:
        raise SystemExit("Falta ETSY_SHOP_ID")

    print(f"Título propuesto: {config['title']}")
    print(f"Secciones propuestas: {', '.join(config['sections'])}")
    if not args.execute:
        print("Modo seguro: no se modifica Etsy.")
        print("Para aplicar: reautoriza con shops_r shops_w y ejecuta con --execute.")
        return 0
    if flag(values, "ETSY_DRY_RUN") or not flag(values, "ETSY_PUBLISH_ENABLED"):
        raise SystemExit("La aplicación está en modo seguro; habilita publicación solo para esta ejecución.")

    token_data = refresh_access_token(values)
    token = str(token_data.get("access_token", ""))
    scopes = set(str(token_data.get("scope", "")).split())
    missing = {"shops_r", "shops_w"} - scopes
    if missing:
        raise SystemExit("El token no tiene los permisos requeridos: " + ", ".join(sorted(missing)))
    rotated_refresh = str(token_data.get("refresh_token", ""))
    if rotated_refresh:
        update_env_value("ETSY_REFRESH_TOKEN", rotated_refresh)

    request_json(
        "PUT",
        f"/application/shops/{shop_id}",
        values,
        oauth_token=token,
        form=[(key, str(config[key])) for key in ("title", "announcement", "sale_message", "digital_sale_message", "policy_additional")],
    )
    existing = request_json("GET", f"/application/shops/{shop_id}/sections", values, oauth_token=token)
    section_ids = {
        str(item.get("title", "")).casefold(): item.get("shop_section_id")
        for item in existing.get("results", [])
    }
    created = 0
    for title in config["sections"]:
        if title.casefold() in section_ids:
            continue
        created_section = request_json(
            "POST",
            f"/application/shops/{shop_id}/sections",
            values,
            oauth_token=token,
            form=[("title", title)],
        )
        section_ids[title.casefold()] = created_section.get("shop_section_id")
        created += 1

    # Clasifica la publicación ya activa en la primera sección de la tienda.
    manifest_path = PROJECT_ROOT / "outputs" / "listings" / "product-7" / "listing.json"
    if manifest_path.exists():
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        listing_id = manifest.get("etsy_listing_id")
        section_id = section_ids.get("plantillas excel")
        if listing_id and section_id:
            request_json(
                "PATCH",
                f"/application/shops/{shop_id}/listings/{listing_id}",
                values,
                oauth_token=token,
                form=[("section_id", str(section_id))],
            )
            print(f"Publicación {listing_id} asignada a Plantillas Excel.")

    print(f"Tienda actualizada. Secciones nuevas: {created}.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except EtsyApiError as error:
        raise SystemExit(str(error)) from error

