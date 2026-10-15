from __future__ import annotations

import json
from pathlib import Path

from pydantic import BaseModel, Field

from listing_package import ListingDraft


class EtsyDraftRequest(BaseModel):
    method: str = "POST"
    endpoint: str
    form_data: dict[str, str]
    file_name: str
    activation_requested: bool = False
    required_credentials: tuple[str, ...] = ("x-api-key", "oauth_listings_w")


def build_etsy_draft_request(
    draft: ListingDraft,
    shop_id: int,
    taxonomy_id: int,
    file_name: str,
) -> EtsyDraftRequest:
    if shop_id <= 0:
        raise ValueError("shop_id debe ser positivo")
    if taxonomy_id <= 0:
        raise ValueError("taxonomy_id debe ser positivo y debe confirmarse en Etsy")
    if not file_name:
        raise ValueError("file_name es obligatorio")
    return EtsyDraftRequest(
        endpoint=f"https://api.etsy.com/v3/application/shops/{shop_id}/listings",
        form_data={
            "quantity": "999",
            "title": draft.title,
            "description": draft.description,
            "price": f"{draft.price:.2f}",
            "who_made": "i_did",
            "when_made": "2020_2026",
            "taxonomy_id": str(taxonomy_id),
            "type": "download",
        },
        file_name=file_name,
    )


def write_etsy_dry_run(request: EtsyDraftRequest, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(request.model_dump(), ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return output_path
