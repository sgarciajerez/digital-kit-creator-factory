from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from factory_core import initialize_database, managed_connection, now_utc, record_audit_event


DEFAULT_WEIGHTS = {
    "evidencia de demanda": 25.0,
    "intensidad del problema y utilidad": 15.0,
    "competencia y diferenciación posible": 15.0,
    "viabilidad de producir con calidad": 15.0,
    "margen potencial tras costes y comisiones": 15.0,
    "facilidad de llegar al público objetivo": 10.0,
    "riesgo legal, de plataforma y mantenimiento": 5.0,
}

REQUIRED_EVIDENCE_COLUMNS = {
    "observed_at",
    "query",
    "market",
    "language",
    "metric_name",
    "metric_value",
    "unit",
    "evidence_type",
    "confidence",
    "limitations",
    "raw_reference",
}


@dataclass(frozen=True)
class EvidenceImportResult:
    research_run_id: int
    source_id: int
    evidence_count: int


@dataclass(frozen=True)
class ScoreComponentInput:
    criterion: str
    weight: float
    score: float
    rationale: str
    evidence_ids: tuple[int, ...] = ()


def _parse_confidence(value: str | None) -> float | None:
    if value is None or not value.strip():
        return None
    confidence = float(value)
    if not 0 <= confidence <= 100:
        raise ValueError("confidence debe estar entre 0 y 100")
    return confidence


def import_evidence_csv(
    database_path: Path,
    csv_path: Path,
    source_name: str,
    source_type: str = "manual_csv",
    source_url: str | None = None,
) -> EvidenceImportResult:
    """Importa observaciones aportadas por el usuario; no hace scraping ni llamadas web."""
    initialize_database(database_path)
    with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        columns = set(reader.fieldnames or ())
        missing = REQUIRED_EVIDENCE_COLUMNS - columns
        if missing:
            raise ValueError(f"Faltan columnas de evidencia: {', '.join(sorted(missing))}")
        rows = list(reader)

    timestamp = now_utc()
    query_config = json.dumps(
        {"method": "manual_csv", "file_name": csv_path.name}, ensure_ascii=False
    )
    with managed_connection(database_path) as connection:
        run_cursor = connection.execute(
            "INSERT INTO research_runs "
            "(started_at, completed_at, query_config, status, summary) VALUES (?, ?, ?, ?, ?)",
            (timestamp, timestamp, query_config, "completed", f"{len(rows)} evidencias importadas"),
        )
        research_run_id = int(run_cursor.lastrowid)
        source_cursor = connection.execute(
            "INSERT INTO sources "
            "(source_type, name, url, access_method, terms_checked_at) VALUES (?, ?, ?, ?, ?)",
            (source_type, source_name, source_url, "manual_csv", timestamp),
        )
        source_id = int(source_cursor.lastrowid)

        for row in rows:
            evidence_type = (row.get("evidence_type") or "observed").strip().lower()
            if evidence_type not in {"observed", "estimated", "hypothesis"}:
                raise ValueError(f"evidence_type no válido: {evidence_type}")
            connection.execute(
                "INSERT INTO evidence "
                "(source_id, research_run_id, observed_at, query, market, language, "
                "metric_name, metric_value, unit, evidence_type, confidence, limitations, raw_reference) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    source_id,
                    research_run_id,
                    row["observed_at"],
                    row["query"],
                    row["market"],
                    row["language"],
                    row["metric_name"],
                    row["metric_value"],
                    row["unit"],
                    evidence_type,
                    _parse_confidence(row.get("confidence")),
                    row["limitations"],
                    row["raw_reference"],
                ),
            )

    record_audit_event(
        database_path,
        "research_imported",
        actor="user",
        entity_type="research_run",
        entity_id=str(research_run_id),
        details_redacted=f"source_id={source_id}; evidence_count={len(rows)}",
    )
    return EvidenceImportResult(research_run_id, source_id, len(rows))


def list_evidence(database_path: Path, research_run_id: int | None = None) -> list[dict]:
    initialize_database(database_path)
    with managed_connection(database_path) as connection:
        if research_run_id is None:
            rows = connection.execute("SELECT * FROM evidence ORDER BY observed_at DESC, id DESC")
        else:
            rows = connection.execute(
                "SELECT * FROM evidence WHERE research_run_id = ? "
                "ORDER BY observed_at DESC, id DESC",
                (research_run_id,),
            )
        return [dict(row) for row in rows.fetchall()]


def score_opportunity(
    database_path: Path,
    opportunity_id: int,
    components: Iterable[ScoreComponentInput],
    confidence: float,
    decision_reason: str | None = None,
) -> float:
    initialize_database(database_path)
    components = list(components)
    if not components:
        raise ValueError("Se necesita al menos un componente de puntuación")
    total_weight = sum(float(component.weight) for component in components)
    if abs(total_weight - 100.0) > 0.01:
        raise ValueError("Los pesos de puntuación deben sumar 100")
    if not 0 <= confidence <= 100:
        raise ValueError("confidence debe estar entre 0 y 100")

    total = 0.0
    demand_score: float | None = None
    legal_score: float | None = None
    for component in components:
        if not 0 <= component.score <= 5:
            raise ValueError("Cada score de criterio debe estar entre 0 y 5")
        if component.weight < 0:
            raise ValueError("Los pesos no pueden ser negativos")
        total += (component.score / 5.0) * component.weight
        normalized = component.criterion.strip().lower()
        if normalized == "evidencia de demanda":
            demand_score = component.score
        if normalized.startswith("riesgo legal"):
            legal_score = component.score

    if legal_score is not None and legal_score < 3:
        decision = "reject"
        reason = "Bloqueo: riesgo legal, de plataforma o mantenimiento demasiado alto"
    elif demand_score is not None and demand_score < 3:
        decision = "review"
        reason = "Evidencia de demanda débil; requiere investigación adicional"
    elif total >= 75:
        decision = "accept"
        reason = "Puntuación suficiente; pendiente de validar calidad de evidencia"
    elif total >= 55:
        decision = "review"
        reason = "Puntuación intermedia; requiere revisión o evidencia adicional"
    else:
        decision = "reject"
        reason = "Puntuación inferior al umbral de producción"
    reason = decision_reason or reason

    with managed_connection(database_path) as connection:
        exists = connection.execute(
            "SELECT 1 FROM opportunities WHERE id = ?", (opportunity_id,)
        ).fetchone()
        if exists is None:
            raise ValueError(f"No existe la oportunidad {opportunity_id}")
        connection.execute("DELETE FROM score_components WHERE opportunity_id = ?", (opportunity_id,))
        for component in components:
            connection.execute(
                "INSERT INTO score_components "
                "(opportunity_id, criterion, weight, score, rationale, evidence_ids) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                (
                    opportunity_id,
                    component.criterion,
                    component.weight,
                    component.score,
                    component.rationale,
                    json.dumps(component.evidence_ids),
                ),
            )
        connection.execute(
            "UPDATE opportunities SET score = ?, confidence = ?, decision = ?, "
            "decision_reason = ?, updated_at = ? WHERE id = ?",
            (total, confidence, decision, reason, now_utc(), opportunity_id),
        )

    record_audit_event(
        database_path,
        "opportunity_scored",
        actor="system",
        entity_type="opportunity",
        entity_id=str(opportunity_id),
        details_redacted=f"score={total:.2f}; confidence={confidence:.2f}; decision={decision}",
    )
    return total


def list_score_components(database_path: Path, opportunity_id: int) -> list[dict]:
    initialize_database(database_path)
    with managed_connection(database_path) as connection:
        rows = connection.execute(
            "SELECT * FROM score_components WHERE opportunity_id = ? ORDER BY id",
            (opportunity_id,),
        ).fetchall()
        return [dict(row) for row in rows]
