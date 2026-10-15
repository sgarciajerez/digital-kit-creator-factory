from __future__ import annotations

import csv
import json
import os
import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterator

try:
    from dotenv import load_dotenv
except ImportError:  # Permite usar el núcleo antes de instalar dependencias opcionales.
    def load_dotenv(*args: Any, **kwargs: Any) -> bool:
        return False


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DATABASE_PATH = PROJECT_ROOT / "data" / "factory.sqlite3"

# El esquema se mantiene en SQL sencillo para que una copia SQLite siga siendo portable.
# Las columnas añadidas después de la primera versión se completan en initialize_database().
SCHEMA = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS app_runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at TEXT NOT NULL,
    completed_at TEXT,
    status TEXT NOT NULL,
    source TEXT NOT NULL,
    details TEXT
);
CREATE TABLE IF NOT EXISTS research_runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at TEXT NOT NULL,
    completed_at TEXT,
    query_config TEXT,
    status TEXT NOT NULL DEFAULT 'pending',
    cost_estimate REAL,
    cost_actual REAL,
    summary TEXT
);
CREATE TABLE IF NOT EXISTS sources (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_type TEXT NOT NULL,
    name TEXT NOT NULL,
    url TEXT,
    access_method TEXT,
    terms_checked_at TEXT,
    notes TEXT
);
CREATE TABLE IF NOT EXISTS evidence (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_id INTEGER,
    research_run_id INTEGER,
    observed_at TEXT NOT NULL,
    query TEXT,
    market TEXT,
    language TEXT,
    metric_name TEXT,
    metric_value TEXT,
    unit TEXT,
    evidence_type TEXT NOT NULL DEFAULT 'observed',
    confidence REAL,
    limitations TEXT,
    raw_reference TEXT,
    FOREIGN KEY(source_id) REFERENCES sources(id),
    FOREIGN KEY(research_run_id) REFERENCES research_runs(id),
    CHECK(evidence_type IN ('observed', 'estimated', 'hypothesis'))
);
CREATE TABLE IF NOT EXISTS opportunities (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    category TEXT,
    subcategory TEXT,
    target_audience TEXT,
    problem TEXT,
    proposed_solution TEXT,
    market_research_json TEXT,
    score REAL,
    confidence REAL,
    decision TEXT NOT NULL DEFAULT 'review',
    decision_reason TEXT,
    next_review_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    CHECK(decision IN ('accept', 'reject', 'review'))
);
CREATE TABLE IF NOT EXISTS score_components (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    opportunity_id INTEGER NOT NULL,
    criterion TEXT NOT NULL,
    weight REAL NOT NULL,
    score REAL NOT NULL,
    rationale TEXT,
    evidence_ids TEXT,
    FOREIGN KEY(opportunity_id) REFERENCES opportunities(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS opportunity_evidence (
    opportunity_id INTEGER NOT NULL,
    evidence_id INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    PRIMARY KEY(opportunity_id, evidence_id),
    FOREIGN KEY(opportunity_id) REFERENCES opportunities(id) ON DELETE CASCADE,
    FOREIGN KEY(evidence_id) REFERENCES evidence(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    opportunity_id INTEGER,
    name TEXT NOT NULL,
    product_type TEXT NOT NULL,
    version TEXT NOT NULL DEFAULT '0.1.0',
    status TEXT NOT NULL DEFAULT 'draft',
    brief_json TEXT,
    output_path TEXT,
    created_at TEXT NOT NULL,
    quality_score REAL,
    quality_report TEXT,
    generation_cost REAL,
    production_time_seconds INTEGER,
    FOREIGN KEY(opportunity_id) REFERENCES opportunities(id)
);
CREATE TABLE IF NOT EXISTS listings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER NOT NULL,
    platform TEXT NOT NULL,
    external_listing_id TEXT,
    title TEXT,
    description TEXT,
    price REAL,
    currency TEXT NOT NULL DEFAULT 'EUR',
    status TEXT NOT NULL DEFAULT 'local',
    created_at TEXT NOT NULL,
    published_at TEXT,
    last_synced_at TEXT,
    policy_review TEXT,
    FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS metrics_daily (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    listing_id INTEGER NOT NULL,
    date TEXT NOT NULL,
    impressions INTEGER,
    visits INTEGER,
    favorites INTEGER,
    orders INTEGER,
    units_sold INTEGER,
    gross_revenue REAL,
    refunds REAL,
    fees REAL,
    ad_spend REAL,
    source TEXT,
    data_quality_notes TEXT,
    FOREIGN KEY(listing_id) REFERENCES listings(id) ON DELETE CASCADE,
    UNIQUE(listing_id, date)
);
CREATE TABLE IF NOT EXISTS expenses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT NOT NULL,
    category TEXT NOT NULL,
    provider TEXT,
    amount REAL NOT NULL,
    currency TEXT NOT NULL DEFAULT 'EUR',
    attributed_product_id INTEGER,
    attributed_research_run_id INTEGER,
    provider_model TEXT,
    units REAL,
    cost_estimate REAL,
    cost_actual REAL,
    scope TEXT,
    scope_id TEXT,
    notes TEXT,
    FOREIGN KEY(attributed_product_id) REFERENCES products(id),
    FOREIGN KEY(attributed_research_run_id) REFERENCES research_runs(id),
    CHECK(amount >= 0)
);
CREATE TABLE IF NOT EXISTS experiments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    hypothesis TEXT NOT NULL,
    variable TEXT,
    primary_metric TEXT,
    budget_cap REAL,
    start_date TEXT,
    end_date TEXT,
    status TEXT NOT NULL DEFAULT 'planned',
    result TEXT,
    decision TEXT
);
CREATE TABLE IF NOT EXISTS approvals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    action_type TEXT NOT NULL,
    target_type TEXT NOT NULL,
    target_id TEXT NOT NULL,
    requested_at TEXT NOT NULL,
    approved_at TEXT,
    approved_by TEXT,
    decision TEXT NOT NULL DEFAULT 'pending',
    notes TEXT
);
CREATE TABLE IF NOT EXISTS audit_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT NOT NULL,
    actor TEXT NOT NULL,
    event_type TEXT NOT NULL,
    entity_type TEXT,
    entity_id TEXT,
    details_redacted TEXT
);
"""

INDEXES = """
CREATE INDEX IF NOT EXISTS idx_app_runs_started_at ON app_runs(started_at);
CREATE INDEX IF NOT EXISTS idx_research_runs_status ON research_runs(status);
CREATE INDEX IF NOT EXISTS idx_evidence_research_run ON evidence(research_run_id);
CREATE INDEX IF NOT EXISTS idx_opportunities_decision ON opportunities(decision);
CREATE INDEX IF NOT EXISTS idx_opportunities_updated_at ON opportunities(updated_at);
CREATE INDEX IF NOT EXISTS idx_score_components_opportunity ON score_components(opportunity_id);
CREATE INDEX IF NOT EXISTS idx_opportunity_evidence_evidence ON opportunity_evidence(evidence_id);
CREATE INDEX IF NOT EXISTS idx_products_status ON products(status);
CREATE INDEX IF NOT EXISTS idx_products_opportunity ON products(opportunity_id);
CREATE INDEX IF NOT EXISTS idx_listings_product ON listings(product_id);
CREATE INDEX IF NOT EXISTS idx_metrics_daily_date ON metrics_daily(date);
CREATE INDEX IF NOT EXISTS idx_expenses_date ON expenses(date);
CREATE INDEX IF NOT EXISTS idx_audit_events_timestamp ON audit_events(timestamp);
"""

MIGRATION_COLUMNS = {
    "opportunities": {
        "subcategory": "TEXT",
        "proposed_solution": "TEXT",
        "market_research_json": "TEXT",
        "next_review_at": "TEXT",
    },
    "products": {
        "brief_json": "TEXT",
        "quality_report": "TEXT",
        "generation_cost": "REAL",
        "production_time_seconds": "INTEGER",
    },
    "expenses": {
        "attributed_product_id": "INTEGER",
        "attributed_research_run_id": "INTEGER",
        "provider_model": "TEXT",
        "units": "REAL",
        "cost_estimate": "REAL",
        "cost_actual": "REAL",
        "scope": "TEXT",
        "scope_id": "TEXT",
    },
}

EXPORTABLE_TABLES = {
    "app_runs",
    "research_runs",
    "sources",
    "evidence",
    "opportunities",
    "score_components",
    "opportunity_evidence",
    "products",
    "listings",
    "metrics_daily",
    "expenses",
    "experiments",
    "approvals",
    "audit_events",
}


@dataclass(frozen=True)
class Settings:
    monthly_budget_eur: float = 15.0
    daily_budget_eur: float | None = None
    task_budget_eur: float | None = None
    product_budget_eur: float | None = None
    batch_budget_eur: float | None = None
    autonomy_mode: int = 2
    database_path: Path = DATABASE_PATH
    ai_provider: str = "mock"
    ai_model: str = ""
    ai_api_key_configured: bool = False


def now_utc() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def load_settings() -> Settings:
    load_dotenv(PROJECT_ROOT / ".env", override=False)

    raw_path = os.getenv("DATABASE_PATH", str(DATABASE_PATH)).strip()
    path = Path(raw_path or DATABASE_PATH)
    if not path.is_absolute():
        path = PROJECT_ROOT / path

    raw_budget = os.getenv("MONTHLY_BUDGET_EUR", "15.0")
    try:
        monthly_budget = float(raw_budget)
    except ValueError as error:
        raise ValueError("MONTHLY_BUDGET_EUR debe ser un número válido") from error
    if monthly_budget < 0:
        raise ValueError("MONTHLY_BUDGET_EUR no puede ser negativo")

    def optional_budget(name: str) -> float | None:
        raw = os.getenv(name, "").strip()
        if not raw:
            return None
        try:
            value = float(raw)
        except ValueError as error:
            raise ValueError(f"{name} debe ser un número válido") from error
        if value < 0:
            raise ValueError(f"{name} no puede ser negativo")
        return value

    raw_autonomy = os.getenv("AUTONOMY_MODE", "2")
    try:
        autonomy_mode = int(raw_autonomy)
    except ValueError as error:
        raise ValueError("AUTONOMY_MODE debe ser un entero entre 0 y 4") from error
    if autonomy_mode not in range(5):
        raise ValueError("AUTONOMY_MODE debe ser un entero entre 0 y 4")

    return Settings(
        monthly_budget_eur=monthly_budget,
        daily_budget_eur=optional_budget("DAILY_BUDGET_EUR"),
        task_budget_eur=optional_budget("TASK_BUDGET_EUR"),
        product_budget_eur=optional_budget("PRODUCT_BUDGET_EUR"),
        batch_budget_eur=optional_budget("BATCH_BUDGET_EUR"),
        autonomy_mode=autonomy_mode,
        database_path=path,
        ai_provider=os.getenv("AI_PROVIDER", "mock").strip() or "mock",
        ai_model=os.getenv("AI_MODEL", "").strip(),
        ai_api_key_configured=bool(os.getenv("AI_API_KEY", "").strip()),
    )


def connect(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


@contextmanager
def managed_connection(path: Path) -> Iterator[sqlite3.Connection]:
    connection = connect(path)
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def _table_columns(connection: sqlite3.Connection, table: str) -> set[str]:
    return {row[1] for row in connection.execute(f"PRAGMA table_info({table})")}


def _apply_column_migrations(connection: sqlite3.Connection) -> None:
    for table, columns in MIGRATION_COLUMNS.items():
        existing = _table_columns(connection, table)
        for name, definition in columns.items():
            if name not in existing:
                connection.execute(f"ALTER TABLE {table} ADD COLUMN {name} {definition}")


def initialize_database(path: Path) -> None:
    with managed_connection(path) as connection:
        version = int(connection.execute("PRAGMA user_version").fetchone()[0])
        if version < 1:
            connection.executescript(SCHEMA)
            connection.execute("PRAGMA user_version = 1")
        _apply_column_migrations(connection)
        connection.execute(
            "CREATE TABLE IF NOT EXISTS opportunity_evidence ("
            "opportunity_id INTEGER NOT NULL, evidence_id INTEGER NOT NULL, created_at TEXT NOT NULL, "
            "PRIMARY KEY(opportunity_id, evidence_id), "
            "FOREIGN KEY(opportunity_id) REFERENCES opportunities(id) ON DELETE CASCADE, "
            "FOREIGN KEY(evidence_id) REFERENCES evidence(id) ON DELETE CASCADE)"
        )
        connection.executescript(INDEXES)
        if version < 2:
            connection.execute("PRAGMA user_version = 2")
        if version < 3:
            connection.execute("PRAGMA user_version = 3")


def _insert_audit(
    connection: sqlite3.Connection,
    actor: str,
    event_type: str,
    entity_type: str | None = None,
    entity_id: str | None = None,
    details_redacted: str = "",
) -> int:
    cursor = connection.execute(
        "INSERT INTO audit_events "
        "(timestamp, actor, event_type, entity_type, entity_id, details_redacted) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (now_utc(), actor, event_type, entity_type, entity_id, details_redacted),
    )
    return int(cursor.lastrowid)


def record_audit_event(
    path: Path,
    event_type: str,
    actor: str = "system",
    entity_type: str | None = None,
    entity_id: str | None = None,
    details_redacted: str = "",
) -> int:
    initialize_database(path)
    with managed_connection(path) as connection:
        return _insert_audit(
            connection,
            actor,
            event_type,
            entity_type,
            entity_id,
            details_redacted,
        )


def record_run(path: Path, source: str = "startup") -> int:
    initialize_database(path)
    timestamp = now_utc()
    with managed_connection(path) as connection:
        cursor = connection.execute(
            "INSERT INTO app_runs (started_at, completed_at, status, source) "
            "VALUES (?, ?, ?, ?)",
            (timestamp, timestamp, "completed", source),
        )
        _insert_audit(
            connection,
            "system",
            "app_started",
            "app_run",
            str(cursor.lastrowid),
        )
        return int(cursor.lastrowid)


def _require_text(value: str, field: str) -> str:
    normalized = value.strip()
    if not normalized:
        raise ValueError(f"{field} no puede estar vacío")
    return normalized


def _optional_score(value: float | None, field: str) -> float | None:
    if value is None:
        return None
    numeric = float(value)
    if not 0 <= numeric <= 100:
        raise ValueError(f"{field} debe estar entre 0 y 100")
    return numeric


def _row_dict(row: sqlite3.Row | None) -> dict[str, Any] | None:
    return dict(row) if row is not None else None


def create_opportunity(
    path: Path,
    name: str,
    category: str | None = None,
    subcategory: str | None = None,
    target_audience: str | None = None,
    problem: str | None = None,
    proposed_solution: str | None = None,
    market_research_json: dict[str, Any] | str | None = None,
    score: float | None = None,
    confidence: float | None = None,
    decision: str = "review",
    decision_reason: str | None = None,
    next_review_at: str | None = None,
) -> int:
    initialize_database(path)
    name = _require_text(name, "name")
    if decision not in {"accept", "reject", "review"}:
        raise ValueError("decision debe ser accept, reject o review")
    score = _optional_score(score, "score")
    confidence = _optional_score(confidence, "confidence")
    if isinstance(market_research_json, dict):
        market_research_json = json.dumps(market_research_json, ensure_ascii=False)
    elif market_research_json is not None:
        try:
            json.loads(market_research_json)
        except (TypeError, json.JSONDecodeError) as error:
            raise ValueError("market_research_json debe ser un JSON válido") from error
    timestamp = now_utc()
    with managed_connection(path) as connection:
        cursor = connection.execute(
            "INSERT INTO opportunities "
            "(name, category, subcategory, target_audience, problem, proposed_solution, market_research_json, "
            "score, confidence, decision, decision_reason, next_review_at, created_at, updated_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                name,
                category,
                subcategory,
                target_audience,
                problem,
                proposed_solution,
                market_research_json,
                score,
                confidence,
                decision,
                decision_reason,
                next_review_at,
                timestamp,
                timestamp,
            ),
        )
        opportunity_id = int(cursor.lastrowid)
        _insert_audit(connection, "system", "opportunity_created", "opportunity", str(opportunity_id))
        return opportunity_id


def get_opportunity(path: Path, opportunity_id: int) -> dict[str, Any] | None:
    initialize_database(path)
    with managed_connection(path) as connection:
        row = connection.execute(
            "SELECT * FROM opportunities WHERE id = ?", (opportunity_id,)
        ).fetchone()
        return _row_dict(row)


def link_evidence_to_opportunity(path: Path, opportunity_id: int, evidence_ids: list[int]) -> int:
    """Attach research evidence to an opportunity without assigning a score."""
    initialize_database(path)
    normalized = list(dict.fromkeys(int(evidence_id) for evidence_id in evidence_ids))
    if not normalized:
        return 0
    timestamp = now_utc()
    with managed_connection(path) as connection:
        if not connection.execute("SELECT 1 FROM opportunities WHERE id = ?", (opportunity_id,)).fetchone():
            raise ValueError(f"No existe la oportunidad {opportunity_id}")
        placeholders = ",".join("?" for _ in normalized)
        existing = {
            int(row[0]) for row in connection.execute(
                f"SELECT id FROM evidence WHERE id IN ({placeholders})", normalized
            ).fetchall()
        }
        missing = set(normalized) - existing
        if missing:
            raise ValueError("No existe evidencia con ID: " + ", ".join(map(str, sorted(missing))))
        connection.executemany(
            "INSERT OR IGNORE INTO opportunity_evidence (opportunity_id, evidence_id, created_at) VALUES (?, ?, ?)",
            [(opportunity_id, evidence_id, timestamp) for evidence_id in normalized],
        )
        linked_count = int(connection.execute(
            "SELECT COUNT(*) FROM opportunity_evidence WHERE opportunity_id = ?", (opportunity_id,)
        ).fetchone()[0])
        _insert_audit(connection, "system", "evidence_linked_to_opportunity", "opportunity", str(opportunity_id))
        return linked_count


def list_opportunities(path: Path, decision: str | None = None) -> list[dict[str, Any]]:
    initialize_database(path)
    with managed_connection(path) as connection:
        if decision is None:
            rows = connection.execute(
                "SELECT * FROM opportunities ORDER BY updated_at DESC, id DESC"
            ).fetchall()
        else:
            rows = connection.execute(
                "SELECT * FROM opportunities WHERE decision = ? "
                "ORDER BY updated_at DESC, id DESC",
                (decision,),
            ).fetchall()
        return [dict(row) for row in rows]


def update_opportunity(path: Path, opportunity_id: int, **changes: Any) -> bool:
    initialize_database(path)
    allowed = {
        "name",
        "category",
        "subcategory",
        "target_audience",
        "problem",
        "proposed_solution",
        "market_research_json",
        "score",
        "confidence",
        "decision",
        "decision_reason",
        "next_review_at",
    }
    unknown = set(changes) - allowed
    if unknown:
        raise ValueError(f"Campos no permitidos: {', '.join(sorted(unknown))}")
    if not changes:
        return False
    if "name" in changes:
        changes["name"] = _require_text(changes["name"], "name")
    if "decision" in changes and changes["decision"] not in {"accept", "reject", "review"}:
        raise ValueError("decision debe ser accept, reject o review")
    if isinstance(changes.get("market_research_json"), dict):
        changes["market_research_json"] = json.dumps(changes["market_research_json"], ensure_ascii=False)
    elif changes.get("market_research_json") is not None:
        try:
            json.loads(changes["market_research_json"])
        except (TypeError, json.JSONDecodeError) as error:
            raise ValueError("market_research_json debe ser un JSON válido") from error
    for field in ("score", "confidence"):
        if field in changes:
            changes[field] = _optional_score(changes[field], field)

    changes["updated_at"] = now_utc()
    assignments = ", ".join(f"{field} = ?" for field in changes)
    values = [changes[field] for field in changes]
    values.append(opportunity_id)
    with managed_connection(path) as connection:
        cursor = connection.execute(
            f"UPDATE opportunities SET {assignments} WHERE id = ?", values
        )
        if cursor.rowcount:
            _insert_audit(connection, "system", "opportunity_updated", "opportunity", str(opportunity_id))
        return bool(cursor.rowcount)


def delete_opportunity(path: Path, opportunity_id: int) -> bool:
    initialize_database(path)
    with managed_connection(path) as connection:
        cursor = connection.execute("DELETE FROM opportunities WHERE id = ?", (opportunity_id,))
        if cursor.rowcount:
            _insert_audit(connection, "system", "opportunity_deleted", "opportunity", str(opportunity_id))
        return bool(cursor.rowcount)


def create_product(
    path: Path,
    name: str,
    product_type: str,
    opportunity_id: int | None = None,
    version: str = "0.1.0",
    status: str = "draft",
    brief_json: str | None = None,
    output_path: str | None = None,
    quality_score: float | None = None,
    quality_report: str | None = None,
    generation_cost: float | None = None,
    production_time_seconds: int | None = None,
) -> int:
    initialize_database(path)
    name = _require_text(name, "name")
    product_type = _require_text(product_type, "product_type")
    version = _require_text(version, "version")
    quality_score = _optional_score(quality_score, "quality_score")
    if generation_cost is not None and float(generation_cost) < 0:
        raise ValueError("generation_cost no puede ser negativo")
    if production_time_seconds is not None and int(production_time_seconds) < 0:
        raise ValueError("production_time_seconds no puede ser negativo")
    timestamp = now_utc()
    with managed_connection(path) as connection:
        cursor = connection.execute(
            "INSERT INTO products "
            "(opportunity_id, name, product_type, version, status, brief_json, output_path, "
            "created_at, quality_score, quality_report, generation_cost, production_time_seconds) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                opportunity_id,
                name,
                product_type,
                version,
                status,
                brief_json,
                output_path,
                timestamp,
                quality_score,
                quality_report,
                generation_cost,
                production_time_seconds,
            ),
        )
        product_id = int(cursor.lastrowid)
        _insert_audit(connection, "system", "product_created", "product", str(product_id))
        return product_id


def get_product(path: Path, product_id: int) -> dict[str, Any] | None:
    initialize_database(path)
    with managed_connection(path) as connection:
        row = connection.execute("SELECT * FROM products WHERE id = ?", (product_id,)).fetchone()
        return _row_dict(row)


def list_products(path: Path, status: str | None = None) -> list[dict[str, Any]]:
    initialize_database(path)
    with managed_connection(path) as connection:
        if status is None:
            rows = connection.execute("SELECT * FROM products ORDER BY id DESC").fetchall()
        else:
            rows = connection.execute(
                "SELECT * FROM products WHERE status = ? ORDER BY id DESC", (status,)
            ).fetchall()
        return [dict(row) for row in rows]


def update_product(path: Path, product_id: int, **changes: Any) -> bool:
    initialize_database(path)
    allowed = {
        "name",
        "product_type",
        "version",
        "status",
        "brief_json",
        "output_path",
        "quality_score",
        "quality_report",
        "generation_cost",
        "production_time_seconds",
    }
    unknown = set(changes) - allowed
    if unknown:
        raise ValueError(f"Campos no permitidos: {', '.join(sorted(unknown))}")
    if not changes:
        return False
    for field in ("name", "product_type", "version"):
        if field in changes:
            changes[field] = _require_text(changes[field], field)
    if "quality_score" in changes:
        changes["quality_score"] = _optional_score(changes["quality_score"], "quality_score")
    if "generation_cost" in changes and changes["generation_cost"] is not None:
        if float(changes["generation_cost"]) < 0:
            raise ValueError("generation_cost no puede ser negativo")
    if "production_time_seconds" in changes and changes["production_time_seconds"] is not None:
        if int(changes["production_time_seconds"]) < 0:
            raise ValueError("production_time_seconds no puede ser negativo")

    assignments = ", ".join(f"{field} = ?" for field in changes)
    values = [changes[field] for field in changes]
    values.append(product_id)
    with managed_connection(path) as connection:
        cursor = connection.execute(f"UPDATE products SET {assignments} WHERE id = ?", values)
        if cursor.rowcount:
            _insert_audit(connection, "system", "product_updated", "product", str(product_id))
        return bool(cursor.rowcount)


def delete_product(path: Path, product_id: int) -> bool:
    initialize_database(path)
    with managed_connection(path) as connection:
        cursor = connection.execute("DELETE FROM products WHERE id = ?", (product_id,))
        if cursor.rowcount:
            _insert_audit(connection, "system", "product_deleted", "product", str(product_id))
        return bool(cursor.rowcount)


def record_expense(
    path: Path,
    amount: float,
    category: str,
    provider: str | None = None,
    currency: str = "EUR",
    date: str | None = None,
    attributed_product_id: int | None = None,
    attributed_research_run_id: int | None = None,
    provider_model: str | None = None,
    units: float | None = None,
    cost_estimate: float | None = None,
    cost_actual: float | None = None,
    scope: str | None = None,
    scope_id: str | None = None,
    notes: str | None = None,
) -> int:
    initialize_database(path)
    amount = float(amount)
    if amount < 0:
        raise ValueError("amount no puede ser negativo")
    category = _require_text(category, "category")
    currency = _require_text(currency, "currency").upper()
    if len(currency) != 3:
        raise ValueError("currency debe ser un código de tres letras")
    with managed_connection(path) as connection:
        cursor = connection.execute(
            "INSERT INTO expenses "
            "(date, category, provider, amount, currency, attributed_product_id, "
            "attributed_research_run_id, provider_model, units, cost_estimate, cost_actual, "
            "scope, scope_id, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                date or now_utc()[:10],
                category,
                provider,
                amount,
                currency,
                attributed_product_id,
                attributed_research_run_id,
                provider_model,
                units,
                cost_estimate,
                cost_actual,
                scope,
                scope_id,
                notes,
            ),
        )
        expense_id = int(cursor.lastrowid)
        _insert_audit(connection, "system", "expense_recorded", "expense", str(expense_id))
        return expense_id


def list_expenses(path: Path, currency: str | None = None) -> list[dict[str, Any]]:
    initialize_database(path)
    with managed_connection(path) as connection:
        if currency is None:
            rows = connection.execute("SELECT * FROM expenses ORDER BY date DESC, id DESC").fetchall()
        else:
            rows = connection.execute(
                "SELECT * FROM expenses WHERE currency = ? ORDER BY date DESC, id DESC",
                (currency.upper(),),
            ).fetchall()
        return [dict(row) for row in rows]


def export_table_csv(path: Path, table: str, output_path: Path) -> Path:
    if table not in EXPORTABLE_TABLES:
        raise ValueError(f"Tabla no exportable: {table}")
    initialize_database(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with managed_connection(path) as connection:
        cursor = connection.execute(f"SELECT * FROM {table}")
        rows = cursor.fetchall()
        fieldnames = [description[0] for description in cursor.description or ()]
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows({field: row[field] for field in fieldnames} for row in rows)
    return output_path


def export_table_json(path: Path, table: str, output_path: Path) -> Path:
    if table not in EXPORTABLE_TABLES:
        raise ValueError(f"Tabla no exportable: {table}")
    initialize_database(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with managed_connection(path) as connection:
        rows = [dict(row) for row in connection.execute(f"SELECT * FROM {table}").fetchall()]
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(rows, handle, ensure_ascii=False, indent=2, default=str)
    return output_path


def scalar(path: Path, query: str, parameters: tuple[Any, ...] = ()) -> int | float:
    with managed_connection(path) as connection:
        value = connection.execute(query, parameters).fetchone()[0]
        return value or 0
