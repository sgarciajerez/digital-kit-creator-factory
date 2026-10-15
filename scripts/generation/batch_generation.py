from __future__ import annotations

import json
import time
from dataclasses import dataclass
from pathlib import Path

from factory_core import initialize_database, managed_connection, now_utc, record_audit_event
from product_generation import DEFAULT_BRIEF, ProductBrief, QualityReport, create_first_product


DEFAULT_BATCH: tuple[ProductBrief, ...] = (
    DEFAULT_BRIEF,
    ProductBrief(
        name="Planificador de presupuesto para autónomos",
        audience="Autónomos que separan gastos profesionales y personales",
        problem="Necesitan revisar gastos recurrentes y desviaciones por categoría",
        categories=["Operación", "Herramientas", "Transporte", "Impuestos", "Otros"],
    ),
    ProductBrief(
        name="Planificador de presupuesto familiar",
        audience="Familias que quieren coordinar gastos domésticos",
        problem="Necesitan una vista compartida y sencilla del gasto mensual",
        categories=["Casa", "Comida", "Colegio", "Salud", "Ocio", "Otros"],
    ),
    ProductBrief(
        name="Planificador de ahorro para objetivos",
        audience="Personas que ahorran para una meta concreta",
        problem="Necesitan comparar aportaciones previstas y reales",
        categories=["Objetivo", "Aportación", "Gasto variable", "Reserva", "Otros"],
    ),
    ProductBrief(
        name="Planificador de presupuesto para estudiantes",
        audience="Estudiantes que administran un presupuesto limitado",
        problem="Necesitan detectar rápidamente qué categorías se desvían",
        categories=["Alojamiento", "Comida", "Transporte", "Material", "Ocio", "Otros"],
    ),
)


@dataclass(frozen=True)
class BatchReport:
    product_ids: tuple[int, ...]
    total_time_seconds: float
    mean_quality_score: float
    max_name_similarity: float
    review_request_ids: tuple[int, ...]
    report_path: Path


def _token_similarity(left: str, right: str) -> float:
    left_tokens = {token.lower() for token in left.split() if len(token) > 2}
    right_tokens = {token.lower() for token in right.split() if len(token) > 2}
    union = left_tokens | right_tokens
    return 0.0 if not union else len(left_tokens & right_tokens) / len(union)


def request_human_review(database_path: Path, product_id: int) -> int:
    initialize_database(database_path)
    with managed_connection(database_path) as connection:
        cursor = connection.execute(
            "INSERT INTO approvals "
            "(action_type, target_type, target_id, requested_at, decision, notes) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            ("product_review", "product", str(product_id), now_utc(), "pending", "Revisión humana requerida"),
        )
        approval_id = int(cursor.lastrowid)
    record_audit_event(
        database_path,
        "human_review_requested",
        actor="system",
        entity_type="product",
        entity_id=str(product_id),
        details_redacted=f"approval_id={approval_id}",
    )
    return approval_id


def generate_small_batch(
    database_path: Path,
    output_directory: Path,
    briefs: tuple[ProductBrief, ...] = DEFAULT_BATCH,
) -> BatchReport:
    if not 5 <= len(briefs) <= 10:
        raise ValueError("El lote debe contener entre 5 y 10 productos")
    started = time.perf_counter()
    product_ids: list[int] = []
    quality_reports: list[QualityReport] = []
    review_ids: list[int] = []

    for index, brief in enumerate(briefs, start=1):
        product_id, report, _ = create_first_product(
            database_path,
            output_directory,
            brief,
            output_filename=f"presupuesto_variante_{index:02d}_0.1.0.xlsx",
        )
        product_ids.append(product_id)
        quality_reports.append(report)
        review_ids.append(request_human_review(database_path, product_id))

    similarities = [
        _token_similarity(left.name, right.name)
        for index, left in enumerate(briefs)
        for right in briefs[index + 1 :]
    ]
    report_path = output_directory.parent / "reports" / "batch_report.json"
    report_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "product_ids": product_ids,
        "total_time_seconds": round(time.perf_counter() - started, 3),
        "mean_quality_score": round(sum(report.score for report in quality_reports) / len(quality_reports), 2),
        "max_name_similarity": round(max(similarities, default=0.0), 4),
        "review_request_ids": review_ids,
        "ai_cost_eur": 0.0,
        "status": "pending_human_review",
    }
    report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return BatchReport(
        product_ids=tuple(product_ids),
        total_time_seconds=payload["total_time_seconds"],
        mean_quality_score=payload["mean_quality_score"],
        max_name_similarity=payload["max_name_similarity"],
        review_request_ids=tuple(review_ids),
        report_path=report_path,
    )


if __name__ == "__main__":
    from factory_core import DATABASE_PATH, PROJECT_ROOT

    report = generate_small_batch(DATABASE_PATH, PROJECT_ROOT / "outputs" / "products")
    print(f"Lote generado: {len(report.product_ids)} productos")
    print(f"Calidad media: {report.mean_quality_score:.2f}/100")
    print(f"Similitud máxima de nombres: {report.max_name_similarity:.4f}")
    print(f"Revisión humana: {len(report.review_request_ids)} solicitudes")
    print(f"Informe: {report.report_path}")
