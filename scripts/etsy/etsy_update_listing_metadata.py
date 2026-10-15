from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from etsy_client import (
    EtsyApiError,
    load_env,
    refresh_access_token,
    request_json,
    update_env_value,
)
from listing_package import ListingDraft, validate_listing_draft


def env_flag(values: dict[str, str], key: str) -> bool:
    return values.get(key, "").strip().casefold() in {"1", "true", "yes", "on"}


def main() -> int:
    parser = argparse.ArgumentParser(description="Actualiza título y etiquetas de un listing Etsy existente.")
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--execute", action="store_true", help="Aplica cambios reales en Etsy")
    args = parser.parse_args()

    manifest_path = args.manifest.resolve()
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    listing_id = str(manifest.get("etsy_listing_id") or "")
    if not listing_id or manifest.get("etsy_status") != "active":
        raise SystemExit("Se requiere un listing_id de Etsy activo en el manifiesto.")

    research_name = str(manifest.get("seo_keyword_research_file") or "seo_keyword_research.json")
    research_path = manifest_path.parent / research_name
    if not research_path.is_file():
        raise SystemExit(f"Falta el informe SEO: {research_path}")
    research = json.loads(research_path.read_text(encoding="utf-8"))
    draft = ListingDraft.model_validate(manifest["draft"])
    validation = validate_listing_draft(draft)
    if not validation.passed:
        raise SystemExit("El borrador no pasa la validación: " + "; ".join(validation.errors))
    if draft.title != research.get("proposed_title"):
        raise SystemExit("El título del borrador no coincide con el título investigado.")
    if [tag for tag in draft.tags] != [item.get("phrase") for item in research.get("tags", [])]:
        raise SystemExit("Las etiquetas del borrador no coinciden con el informe SEO.")
    if len(draft.tags) != 13 or any(len(tag) > 20 for tag in draft.tags):
        raise SystemExit("Etsy requiere trece etiquetas de hasta veinte caracteres para esta actualización.")

    print(f"Listing: {listing_id}")
    print(f"Título propuesto: {draft.title}")
    print("Etiquetas: " + " | ".join(draft.tags))
    if not args.execute:
        print("Modo seguro: no se modifica Etsy. Usa --execute para aplicar el cambio.")
        return 0

    values = load_env()
    shop_id = values.get("ETSY_SHOP_ID", "").strip()
    if not shop_id:
        raise SystemExit("Falta ETSY_SHOP_ID.")
    if not env_flag(values, "ETSY_PUBLISH_ENABLED") or env_flag(values, "ETSY_DRY_RUN"):
        raise SystemExit("Actualización bloqueada por ETSY_PUBLISH_ENABLED / ETSY_DRY_RUN.")

    token_data = refresh_access_token(values)
    token = str(token_data.get("access_token", ""))
    if "listings_w" not in set(str(token_data.get("scope", "")).split()):
        raise SystemExit("El token OAuth no incluye listings_w.")
    rotated_refresh = str(token_data.get("refresh_token", ""))
    if rotated_refresh:
        update_env_value("ETSY_REFRESH_TOKEN", rotated_refresh)

    current = request_json("GET", f"/application/listings/{listing_id}", values, oauth_token=token)
    if str(current.get("shop_id", "")) != shop_id:
        raise SystemExit("El listing no pertenece a la tienda configurada.")
    if str(current.get("state", "")).casefold() != "active":
        raise SystemExit(f"El listing ya no está activo (state={current.get('state')!r}); no se modificó.")

    old_title = str(current.get("title", ""))
    old_tags = [str(tag) for tag in current.get("tags", [])]
    # Etsy's x-www-form-urlencoded updateListing expects an array as one
    # comma-separated value; repeated `tags` keys retain only the last tag.
    fields = [("title", draft.title), ("tags", ",".join(draft.tags))]
    request_json(
        "PATCH",
        f"/application/shops/{shop_id}/listings/{listing_id}",
        values,
        oauth_token=token,
        form=fields,
    )
    updated = request_json("GET", f"/application/listings/{listing_id}", values, oauth_token=token)
    remote_tags = [str(tag) for tag in updated.get("tags", [])]
    if updated.get("title") != draft.title or remote_tags != draft.tags:
        raise SystemExit("Etsy aceptó la solicitud, pero la verificación no coincide; revisar el listing antes de reintentar.")

    manifest["seo_keyword_research_file"] = research_name
    manifest["draft"]["title"] = draft.title
    manifest["draft"]["tags"] = draft.tags
    manifest["validation"] = validation.model_dump()
    history: list[dict[str, Any]] = list(manifest.get("seo_update_history", []))
    history.append({
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "previous_title": old_title,
        "previous_tags": old_tags,
        "title": draft.title,
        "tags": draft.tags,
        "research_file": research_name,
        "etsy_state": updated.get("state"),
    })
    manifest["seo_update_history"] = history
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("SEO_METADATA=UPDATED_AND_VERIFIED")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except EtsyApiError as error:
        raise SystemExit(str(error)) from error
