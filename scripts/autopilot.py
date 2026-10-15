from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from batch_generation import DEFAULT_BATCH, request_human_review
from factory_core import (
    create_opportunity,
    initialize_database,
    managed_connection,
    now_utc,
    record_audit_event,
)
from listing_package import ListingDraft, create_listing_package
from product_generation import ProductBrief, create_first_product


DEFAULT_MIN_SCORE = 75.0
DEFAULT_MIN_CONFIDENCE = 60.0
DEFAULT_MAX_PRODUCTS = 3


@dataclass(frozen=True)
class AutopilotReport:
    discovered_opportunity_ids: tuple[int, ...]
    selected_opportunity_ids: tuple[int, ...]
    product_ids: tuple[int, ...]
    listing_ids: tuple[int, ...]
    skipped: tuple[str, ...]
    report_path: Path


def register_idea_candidates(
    database_path: Path,
    briefs: Iterable[ProductBrief] = DEFAULT_BATCH,
) -> tuple[int, ...]:
    """Registra hipótesis de producto sin presentarlas como demanda validada."""
    initialize_database(database_path)
    created: list[int] = []
    for brief in briefs:
        with managed_connection(database_path) as connection:
            existing = connection.execute(
                "SELECT id FROM opportunities WHERE name = ? ORDER BY id DESC LIMIT 1",
                (brief.name,),
            ).fetchone()
        if existing is not None:
            continue
        opportunity_id = create_opportunity(
            database_path,
            name=brief.name,
            category="plantillas digitales",
            subcategory=brief.family,
            target_audience=brief.audience,
            problem=brief.problem,
            proposed_solution="Generar una plantilla XLSX de presupuesto para revisión humana",
            decision="review",
            decision_reason=(
                "Hipótesis generada por la aplicación; no contiene evidencia de demanda "
                "y requiere investigación y revisión humana"
            ),
        )
        created.append(opportunity_id)
    if created:
        record_audit_event(
            database_path,
            "idea_candidates_registered",
            actor="system",
            entity_type="opportunity_batch",
            entity_id=",".join(map(str, created)),
            details_redacted=f"count={len(created)}; evidence_status=hypothesis",
        )
    return tuple(created)


def select_publishable_opportunities(
    database_path: Path,
    max_products: int = DEFAULT_MAX_PRODUCTS,
    min_score: float = DEFAULT_MIN_SCORE,
    min_confidence: float = DEFAULT_MIN_CONFIDENCE,
) -> tuple[dict, ...]:
    """Selecciona únicamente oportunidades aceptadas y no producidas todavía."""
    if max_products < 1:
        raise ValueError("max_products debe ser positivo")
    if not 0 <= min_score <= 100 or not 0 <= min_confidence <= 100:
        raise ValueError("Los umbrales deben estar entre 0 y 100")
    initialize_database(database_path)
    with managed_connection(database_path) as connection:
        rows = connection.execute(
            "SELECT o.* FROM opportunities o "
            "WHERE o.decision = 'accept' "
            "AND o.score >= ? AND o.confidence >= ? "
            "AND NOT EXISTS ("
            "  SELECT 1 FROM products p WHERE p.opportunity_id = o.id"
            ") "
            "ORDER BY o.score DESC, o.confidence DESC, o.updated_at DESC, o.id ASC "
            "LIMIT ?",
            (min_score, min_confidence, max_products),
        ).fetchall()
        return tuple(dict(row) for row in rows)


def _brief_from_opportunity(opportunity: dict) -> ProductBrief:
    category = (opportunity.get("category") or "General").strip()
    subcategory = (opportunity.get("subcategory") or "Otros").strip()
    opportunity_text = f"{opportunity.get('name', '')} {opportunity.get('target_audience', '')}".casefold()
    is_family = any(token in opportunity_text for token in ("familiar", "familia", "pareja", "hogar", "couple"))
    family = "family_budget_planner" if is_family else "budget_planner"
    if is_family:
        categories = ["Vivienda", "Alimentación", "Transporte", "Salud", "Educación", "Ocio", "Otros"]
    else:
        categories = [category[:24], subcategory[:24], "Otros"]
    categories = list(dict.fromkeys(value for value in categories if value))
    return ProductBrief(
        family=family,
        name=str(opportunity["name"]),
        audience=(opportunity.get("target_audience") or "Personas que quieren organizar sus gastos"),
        problem=(opportunity.get("problem") or "Necesitan una forma sencilla de revisar presupuesto y gasto real"),
        categories=categories or ["Otros"],
    )


def _listing_draft(brief: ProductBrief) -> ListingDraft:
    raw_tags = ["presupuesto", "planificador", "plantilla", "xlsx", *brief.categories]
    tags: list[str] = []
    for tag in raw_tags:
        normalized = re.sub(r"[^\w -]", "", tag, flags=re.UNICODE).strip()
        if normalized and normalized.casefold() not in {item.casefold() for item in tags}:
            tags.append(normalized[:40])
    return ListingDraft(
        title=brief.name,
        description=(
            f"Plantilla digital en formato XLSX para {brief.audience.lower()}. "
            f"Ayuda a revisar {brief.problem.lower()}. Incluye hojas de instrucciones, "
            "presupuesto y vista previa. Revisa el archivo y las condiciones de venta antes de publicar."
        ),
        tags=tags[:13],
        price=4.99,
        currency=brief.currency,
        # El piloto automático no puede confirmar derechos ni política por sí solo.
        policy_reviewed=False,
        rights_confirmed=False,
    )


def run_autopilot(
    database_path: Path,
    products_directory: Path,
    listings_directory: Path,
    *,
    max_products: int = DEFAULT_MAX_PRODUCTS,
    min_score: float = DEFAULT_MIN_SCORE,
    min_confidence: float = DEFAULT_MIN_CONFIDENCE,
) -> AutopilotReport:
    """Genera productos y paquetes locales para oportunidades ya aceptadas.

    Nunca hace peticiones a Etsy ni marca un anuncio como publicado.
    """
    selected = select_publishable_opportunities(
        database_path,
        max_products=max_products,
        min_score=min_score,
        min_confidence=min_confidence,
    )
    product_ids: list[int] = []
    listing_ids: list[int] = []
    skipped: list[str] = []

    for opportunity in selected:
        opportunity_id = int(opportunity["id"])
        try:
            brief = _brief_from_opportunity(opportunity)
            safe_name = re.sub(r"[^a-zA-Z0-9._-]+", "_", brief.name).strip("_")
            product_id, quality, _ = create_first_product(
                database_path,
                products_directory,
                brief,
                output_filename=f"autopilot_{opportunity_id}_{safe_name[:60]}_0.1.0.xlsx",
                opportunity_id=opportunity_id,
            )
            if not quality.passed:
                skipped.append(f"Oportunidad {opportunity_id}: producto {product_id} necesita revisión de calidad")
                continue
            product_ids.append(product_id)
            request_human_review(database_path, product_id)

            listing_id, validation, _ = create_listing_package(
                database_path,
                listings_directory,
                product_id,
                platform="etsy-dry-run",
                draft=_listing_draft(brief),
            )
            listing_ids.append(listing_id)
            if not validation.passed:
                skipped.append(
                    f"Producto {product_id}: paquete queda en needs_review por política/derechos"
                )
        except Exception as error:
            skipped.append(f"Oportunidad {opportunity_id}: {error}")

    report_path = products_directory.parent / "reports" / "autopilot_report.json"
    report_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "generated_at": now_utc(),
        "mode": "local_supervised_draft",
        "selected_opportunity_ids": [int(row["id"]) for row in selected],
        "product_ids": product_ids,
        "listing_ids": listing_ids,
        "skipped": skipped,
        "publication_requested": False,
        "status": "pending_human_review" if product_ids or listing_ids else "no_eligible_opportunities",
    }
    report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    record_audit_event(
        database_path,
        "autopilot_completed",
        actor="system",
        entity_type="autopilot_run",
        entity_id=report_path.name,
        details_redacted=(
            f"selected={len(selected)}; products={len(product_ids)}; listings={len(listing_ids)}; "
            "publication_requested=false"
        ),
    )
    return AutopilotReport(
        discovered_opportunity_ids=(),
        selected_opportunity_ids=tuple(int(row["id"]) for row in selected),
        product_ids=tuple(product_ids),
        listing_ids=tuple(listing_ids),
        skipped=tuple(skipped),
        report_path=report_path,
    )


if __name__ == "__main__":
    from factory_core import DATABASE_PATH, PROJECT_ROOT

    report = run_autopilot(
        DATABASE_PATH,
        PROJECT_ROOT / "outputs" / "products",
        PROJECT_ROOT / "outputs" / "listings",
    )
    print(f"Oportunidades seleccionadas: {len(report.selected_opportunity_ids)}")
    print(f"Productos preparados: {len(report.product_ids)}")
    print(f"Paquetes Etsy locales: {len(report.listing_ids)}")
    print(f"Informe: {report.report_path}")
