from __future__ import annotations

import hashlib
import json
import re
import shutil
from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field

from factory_core import (
    DATABASE_PATH,
    PROJECT_ROOT,
    get_product,
    initialize_database,
    managed_connection,
    now_utc,
    record_audit_event,
)


FORBIDDEN_PHRASES = (
    "resultados garantizados",
    "gana dinero rápido",
    "100% garantizado",
    "testimonio verificado",
    "número uno",
)


class ListingDraft(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = Field(min_length=5, max_length=140)
    description: str = Field(min_length=20, max_length=5000)
    tags: list[str] = Field(min_length=1, max_length=13)
    price: float = Field(gt=0)
    currency: str = Field(default="EUR", min_length=3, max_length=3)
    policy_reviewed: bool = False
    rights_confirmed: bool = False


class ListingValidation(BaseModel):
    passed: bool
    errors: list[str]
    warnings: list[str]
    sha256: str | None = None


def validate_listing_draft(draft: ListingDraft) -> ListingValidation:
    errors: list[str] = []
    warnings: list[str] = []
    text = f"{draft.title} {draft.description}".casefold()
    if any(phrase in text for phrase in FORBIDDEN_PHRASES):
        errors.append("El texto contiene una afirmación comercial no permitida")
    if not draft.policy_reviewed:
        errors.append("Falta revisión de política y contenido")
    if not draft.rights_confirmed:
        errors.append("Falta confirmar derechos de los materiales")
    if len({tag.casefold() for tag in draft.tags}) != len(draft.tags):
        warnings.append("Hay etiquetas duplicadas")
    if any(not re.fullmatch(r"[\w -]{2,40}", tag, flags=re.UNICODE) for tag in draft.tags):
        errors.append("Hay etiquetas con formato no permitido")
    return ListingValidation(passed=not errors, errors=errors, warnings=warnings)


def create_listing_package(
    database_path: Path,
    output_directory: Path,
    product_id: int,
    platform: str,
    draft: ListingDraft,
) -> tuple[int, ListingValidation, Path]:
    initialize_database(database_path)
    product = get_product(database_path, product_id)
    if product is None:
        raise ValueError(f"No existe el producto {product_id}")
    product_path = Path(product["output_path"]) if product.get("output_path") else None
    if product_path is None or not product_path.exists():
        raise ValueError("El producto no tiene un archivo local válido")

    validation = validate_listing_draft(draft)
    package_directory = output_directory / f"product-{product_id}"
    package_directory.mkdir(parents=True, exist_ok=True)
    copied_product = package_directory / product_path.name
    shutil.copy2(product_path, copied_product)
    manifest = {
        "product_id": product_id,
        "platform": platform,
        "status": "pending_approval" if validation.passed else "needs_review",
        "draft": draft.model_dump(),
        "validation": validation.model_dump(),
        "local_file": copied_product.name,
        "sha256": hashlib.sha256(copied_product.read_bytes()).hexdigest(),
    }
    (package_directory / "listing.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (package_directory / "description.txt").write_text(draft.description, encoding="utf-8")
    (package_directory / "tags.txt").write_text("\n".join(draft.tags), encoding="utf-8")

    with managed_connection(database_path) as connection:
        cursor = connection.execute(
            "INSERT INTO listings "
            "(product_id, platform, title, description, price, currency, status, created_at, policy_review) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                product_id,
                platform,
                draft.title,
                draft.description,
                draft.price,
                draft.currency.upper(),
                manifest["status"],
                now_utc(),
                json.dumps(validation.model_dump(), ensure_ascii=False),
            ),
        )
        listing_id = int(cursor.lastrowid)
        connection.execute(
            "INSERT INTO approvals "
            "(action_type, target_type, target_id, requested_at, decision, notes) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (
                "listing_review",
                "listing",
                str(listing_id),
                now_utc(),
                "pending",
                "Aprobación humana necesaria antes de publicar",
            ),
        )
    record_audit_event(
        database_path,
        "listing_package_created",
        actor="system",
        entity_type="listing",
        entity_id=str(listing_id),
        details_redacted=f"platform={platform}; status={manifest['status']}",
    )
    return listing_id, validation, package_directory


if __name__ == "__main__":
    draft = ListingDraft(
        title="Planificador de presupuesto personal",
        description="Plantilla local para organizar presupuesto y gasto real por categorías.",
        tags=["presupuesto", "planificación", "gastos"],
        price=4.99,
        policy_reviewed=True,
        rights_confirmed=True,
    )
    listing_id, validation, package_directory = create_listing_package(
        DATABASE_PATH,
        PROJECT_ROOT / "outputs" / "listings",
        product_id=1,
        platform="manual-upload",
        draft=draft,
    )
    print(f"Paquete creado: {listing_id}")
    print(f"Validación: {'OK' if validation.passed else 'REVISAR'}")
    print(f"Directorio: {package_directory}")
