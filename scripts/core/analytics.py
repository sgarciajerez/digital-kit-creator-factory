from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from factory_core import (
    DATABASE_PATH,
    PROJECT_ROOT,
    initialize_database,
    managed_connection,
    now_utc,
    record_audit_event,
)


METRIC_COLUMNS = {
    "date",
    "impressions",
    "visits",
    "favorites",
    "orders",
    "units_sold",
    "gross_revenue",
    "refunds",
    "fees",
    "ad_spend",
    "source",
    "data_quality_notes",
}


@dataclass(frozen=True)
class MetricsImportResult:
    listing_id: int
    rows_imported: int


@dataclass(frozen=True)
class ListingMetricsReport:
    listing_id: int
    date_from: str | None
    date_to: str | None
    impressions: int
    visits: int
    orders: int
    units_sold: int
    gross_revenue: float
    refunds: float
    fees: float
    ad_spend: float
    ai_cost_attributed: float
    conversion_rate: float | None
    contribution_margin_estimated: float | None
    warnings: tuple[str, ...]


def _number(row: dict[str, str], name: str, integer: bool = False) -> int | float:
    value = (row.get(name) or "").strip()
    if not value:
        return 0
    parsed = float(value)
    if integer:
        return int(parsed)
    return parsed


def import_metrics_csv(
    database_path: Path,
    listing_id: int,
    csv_path: Path,
) -> MetricsImportResult:
    initialize_database(database_path)
    with csv_path.open("r", newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        missing = METRIC_COLUMNS - set(reader.fieldnames or ())
        if missing:
            raise ValueError(f"Faltan columnas de métricas: {', '.join(sorted(missing))}")
        rows = list(reader)

    with managed_connection(database_path) as connection:
        exists = connection.execute("SELECT 1 FROM listings WHERE id = ?", (listing_id,)).fetchone()
        if exists is None:
            raise ValueError(f"No existe el listing {listing_id}")
        for row in rows:
            date.fromisoformat(row["date"])
            connection.execute(
                "INSERT INTO metrics_daily "
                "(listing_id, date, impressions, visits, favorites, orders, units_sold, "
                "gross_revenue, refunds, fees, ad_spend, source, data_quality_notes) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
                "ON CONFLICT(listing_id, date) DO UPDATE SET "
                "impressions=excluded.impressions, visits=excluded.visits, "
                "favorites=excluded.favorites, orders=excluded.orders, "
                "units_sold=excluded.units_sold, gross_revenue=excluded.gross_revenue, "
                "refunds=excluded.refunds, fees=excluded.fees, ad_spend=excluded.ad_spend, "
                "source=excluded.source, data_quality_notes=excluded.data_quality_notes",
                (
                    listing_id,
                    row["date"],
                    _number(row, "impressions", True),
                    _number(row, "visits", True),
                    _number(row, "favorites", True),
                    _number(row, "orders", True),
                    _number(row, "units_sold", True),
                    _number(row, "gross_revenue"),
                    _number(row, "refunds"),
                    _number(row, "fees"),
                    _number(row, "ad_spend"),
                    row["source"],
                    row["data_quality_notes"],
                ),
            )
    record_audit_event(
        database_path,
        "metrics_imported",
        actor="user",
        entity_type="listing",
        entity_id=str(listing_id),
        details_redacted=f"rows={len(rows)}; source=csv",
    )
    return MetricsImportResult(listing_id, len(rows))


def calculate_listing_report(database_path: Path, listing_id: int) -> ListingMetricsReport:
    initialize_database(database_path)
    with managed_connection(database_path) as connection:
        listing = connection.execute(
            "SELECT product_id FROM listings WHERE id = ?", (listing_id,)
        ).fetchone()
        if listing is None:
            raise ValueError(f"No existe el listing {listing_id}")
        row = connection.execute(
            "SELECT MIN(date), MAX(date), COALESCE(SUM(impressions), 0), "
            "COALESCE(SUM(visits), 0), COALESCE(SUM(orders), 0), "
            "COALESCE(SUM(units_sold), 0), COALESCE(SUM(gross_revenue), 0), "
            "COALESCE(SUM(refunds), 0), COALESCE(SUM(fees), 0), "
            "COALESCE(SUM(ad_spend), 0) FROM metrics_daily WHERE listing_id = ?",
            (listing_id,),
        ).fetchone()
        ai_cost = connection.execute(
            "SELECT COALESCE(SUM(amount), 0) FROM expenses "
            "WHERE category = 'ai' AND currency = 'EUR' AND attributed_product_id = ?",
            (listing["product_id"],),
        ).fetchone()[0]
        notes = [
            item[0]
            for item in connection.execute(
                "SELECT data_quality_notes FROM metrics_daily "
                "WHERE listing_id = ? AND data_quality_notes IS NOT NULL AND data_quality_notes <> ''",
                (listing_id,),
            ).fetchall()
        ]

    date_from, date_to, impressions, visits, orders, units_sold, gross, refunds, fees, ad_spend = row
    warnings: list[str] = list(notes)
    if not date_from:
        warnings.append("No hay métricas importadas")
    if visits == 0:
        warnings.append("No hay visitas; no se puede interpretar conversión")
    if orders == 0:
        warnings.append("No hay pedidos observados; no concluir falta de demanda")
    conversion = None if visits == 0 else orders / visits * 100
    margin = None if not date_from else gross - refunds - fees - ad_spend - float(ai_cost or 0)
    return ListingMetricsReport(
        listing_id=listing_id,
        date_from=date_from,
        date_to=date_to,
        impressions=int(impressions),
        visits=int(visits),
        orders=int(orders),
        units_sold=int(units_sold),
        gross_revenue=float(gross),
        refunds=float(refunds),
        fees=float(fees),
        ad_spend=float(ad_spend),
        ai_cost_attributed=float(ai_cost or 0),
        conversion_rate=conversion,
        contribution_margin_estimated=margin,
        warnings=tuple(warnings),
    )


def create_experiment(
    database_path: Path,
    name: str,
    hypothesis: str,
    primary_metric: str,
    budget_cap: float,
) -> int:
    initialize_database(database_path)
    if budget_cap < 0:
        raise ValueError("budget_cap no puede ser negativo")
    with managed_connection(database_path) as connection:
        cursor = connection.execute(
            "INSERT INTO experiments "
            "(name, hypothesis, primary_metric, budget_cap, start_date, status) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (name, hypothesis, primary_metric, budget_cap, date.today().isoformat(), "planned"),
        )
        experiment_id = int(cursor.lastrowid)
    record_audit_event(
        database_path,
        "experiment_created",
        actor="system",
        entity_type="experiment",
        entity_id=str(experiment_id),
        details_redacted=f"budget_cap={budget_cap}; metric={primary_metric}",
    )
    return experiment_id


def export_listing_report(database_path: Path, listing_id: int, output_path: Path) -> Path:
    report = calculate_listing_report(database_path, listing_id)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report.__dict__, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


if __name__ == "__main__":
    print("Este módulo requiere un CSV de métricas aportado manualmente; no inventa datos.")
    print(f"Base configurada: {DATABASE_PATH}")
    print(f"Fecha UTC: {now_utc()}")
