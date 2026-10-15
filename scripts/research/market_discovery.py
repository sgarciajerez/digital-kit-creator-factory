from __future__ import annotations

"""Authorized Etsy supply-signal collection for discovering digital-kit niches.

Etsy API listing counts and samples describe supply/competition only; they are
not treated as search volume, buyer demand, or sales evidence.
"""

import json
import re
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from statistics import median
from urllib.parse import urlencode

from etsy_client import api_key, request_json
from factory_core import PROJECT_ROOT, initialize_database, managed_connection, now_utc, record_audit_event


CONFIG_PATH = PROJECT_ROOT / "config" / "market_discovery.json"
ETSY_LISTINGS_DOCS = "https://developers.etsy.com/documentation/reference"


def load_discovery_config(path: Path = CONFIG_PATH) -> dict:
    config = json.loads(path.read_text(encoding="utf-8"))
    if config.get("schema_version") != "1.0.0":
        raise ValueError("Versión de config/market_discovery.json no compatible")
    return config


def discovery_catalog(language: str, config: dict | None = None) -> list[dict]:
    """Return the configured starter hypotheses for a supported language."""
    config = config or load_discovery_config()
    catalog = config.get("starter_catalog", {}).get(language)
    if not catalog:
        raise ValueError(f"No hay catálogo inicial configurado para el idioma: {language}")
    return [dict(item) for item in catalog]


def unscanned_catalog(language: str, database_path: Path, config: dict | None = None) -> list[dict]:
    """Prioritize catalog hypotheses not previously scanned in this language."""
    catalog = discovery_catalog(language, config)
    initialize_database(database_path)
    with managed_connection(database_path) as connection:
        rows = connection.execute(
            "SELECT query_config FROM research_runs WHERE status = 'completed'"
        ).fetchall()
    scanned: set[str] = set()
    for row in rows:
        try:
            query_config = json.loads(row["query_config"])
        except (TypeError, json.JSONDecodeError):
            continue
        if query_config.get("language") == language:
            scanned.update(str(query).strip().casefold() for query in query_config.get("queries", []))
    return [item for item in catalog if item["query"].strip().casefold() not in scanned]


def _digital_listing(row: dict) -> bool:
    listing_type = str(row.get("listing_type") or "").casefold()
    return listing_type in {"download", "digital"}


def _price(row: dict) -> float | None:
    value = row.get("price") or {}
    try:
        amount = float(value["amount"]) / float(value["divisor"])
    except (KeyError, TypeError, ValueError, ZeroDivisionError):
        return None
    return amount if amount > 0 else None


def _phrase_clusters(listings: list[dict], minimum_support: int, stopwords: set[str]) -> list[dict]:
    support: Counter[str] = Counter()
    examples: dict[str, list[dict]] = {}
    for listing in listings:
        words = [
            token for token in re.findall(r"[a-z0-9]+", str(listing.get("title") or "").casefold())
            if len(token) >= 3 and token not in stopwords
        ]
        phrases = set()
        for size in (2, 3):
            phrases.update(" ".join(words[index:index + size]) for index in range(len(words) - size + 1))
        for phrase in phrases:
            support[phrase] += 1
            examples.setdefault(phrase, []).append(listing)
    clusters = []
    for phrase, count in support.most_common():
        if count < minimum_support:
            continue
        rows = examples[phrase]
        prices = [float(row["price"]) for row in rows if row.get("price") is not None]
        clusters.append({
            "phrase": phrase,
            "listing_support": count,
            "median_price": round(median(prices), 2) if prices else None,
            "sample_titles": list(dict.fromkeys(str(row.get("title") or "") for row in rows))[:5],
            "sample_urls": list(dict.fromkeys(str(row.get("url") or "") for row in rows))[:5],
            "signal_type": "competition_phrase_only",
            "status": "hypothesis_to_validate",
        })
    return clusters[:30]


def _insert_run(database_path: Path, queries: list[str], market: str, language: str) -> tuple[int, int]:
    initialize_database(database_path)
    timestamp = now_utc()
    with managed_connection(database_path) as connection:
        cursor = connection.execute(
            "INSERT INTO research_runs (started_at, query_config, status, summary) VALUES (?, ?, 'running', ?) ",
            (timestamp, json.dumps({"queries": queries, "market": market, "language": language}, ensure_ascii=False), "Automated niche discovery scan"),
        )
        run_id = int(cursor.lastrowid)
        source_cursor = connection.execute(
            "INSERT INTO sources (source_type, name, url, access_method, terms_checked_at, notes) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            ("official_api", "Etsy Open API active listings", ETSY_LISTINGS_DOCS, "etsy_open_api_v3", timestamp,
             "Active listing counts and samples are competition/supply signals, not buyer-demand evidence."),
        )
        return run_id, int(source_cursor.lastrowid)


def scan_etsy_supply(
    queries: list[str],
    api_values: dict[str, str],
    *,
    shop_id: str | int | None,
    market: str = "US",
    language: str = "en",
    limit_per_query: int = 40,
) -> dict:
    """Collect active Etsy listing counts/prices through the official API."""
    config = load_discovery_config()
    normalized = list(dict.fromkeys(str(query).strip() for query in queries if str(query).strip()))
    max_queries = int(config["etsy_queries_per_scan"])
    if not normalized:
        raise ValueError("Introduce al menos una búsqueda para Etsy")
    if len(normalized) > max_queries:
        raise ValueError(f"Máximo {max_queries} consultas Etsy por exploración")
    if not 1 <= limit_per_query <= 100:
        raise ValueError("limit_per_query debe estar entre 1 y 100")
    # Validate before making calls so the UI can report a simple local setup issue.
    api_key(api_values)
    all_listings: dict[str, dict] = {}
    search_results = []
    for query in normalized:
        path = "/application/listings/active?" + urlencode({
            "keywords": query,
            "limit": limit_per_query,
            "sort_on": "score",
            "sort_order": "desc",
        })
        response = request_json("GET", path, api_values)
        rows = response.get("results") or []
        sampled_digital = []
        for row in rows:
            listing_id = str(row.get("listing_id") or "")
            if not listing_id or (shop_id is not None and str(row.get("shop_id")) == str(shop_id)):
                continue
            if not _digital_listing(row):
                continue
            listing = {
                "listing_id": listing_id,
                "title": str(row.get("title") or ""),
                "url": str(row.get("url") or f"https://www.etsy.com/listing/{listing_id}"),
                "price": _price(row),
                "currency": str((row.get("price") or {}).get("currency_code") or ""),
                "taxonomy_id": row.get("taxonomy_id"),
                "query": query,
            }
            sampled_digital.append(listing)
            all_listings.setdefault(listing_id, listing)
        search_results.append({
            "query": query,
            "total_active_results": int(response.get("count", len(rows))),
            "sampled_results": len(rows),
            "sampled_digital_competitors": len(sampled_digital),
            "digital_price_median": round(median([row["price"] for row in sampled_digital if row["price"] is not None]), 2)
                if any(row["price"] is not None for row in sampled_digital) else None,
            "listings": sampled_digital[:15],
            "source_url": "https://www.etsy.com/search?" + urlencode({"q": query}),
        })
    config_stopwords = set(config.get("title_stopwords", []))
    clusters = _phrase_clusters(list(all_listings.values()), int(config["title_cluster_minimum_support"]), config_stopwords)
    return {
        "marketplace": "Etsy",
        "market": market,
        "language": language,
        "research_date": datetime.now(timezone.utc).date().isoformat(),
        "results": search_results,
        "unique_digital_listings": len(all_listings),
        "supply_phrase_clusters": clusters,
        "limitations": [
            "Etsy active-listing counts and repeated title phrases measure supply/competition, not buyer search volume or purchases.",
            "Only the result page sample is inspected; the API total may include physical listings and the full market is not crawled.",
            "Digital listings in the sample are inferred from the API listing_type field.",
        ],
    }


def save_discovery_run(
    database_path: Path,
    *,
    queries: list[str],
    market: str,
    language: str,
    etsy_report: dict,
    report_directory: Path | None = None,
) -> dict:
    """Persist raw, dated source observations and a human-readable JSON report."""
    run_id, etsy_source_id = _insert_run(database_path, queries, market, language)
    observed_at = now_utc()
    with managed_connection(database_path) as connection:
        for result in etsy_report["results"]:
            for metric_name, metric_value, unit, limitations in (
                ("active listing count", result["total_active_results"], "listings", "Etsy listing count is supply, not demand."),
                ("sampled digital competitor count", result["sampled_digital_competitors"], "listings", "Sample count is not a full-market count."),
            ):
                connection.execute(
                    "INSERT INTO evidence (source_id, research_run_id, observed_at, query, market, language, "
                    "metric_name, metric_value, unit, evidence_type, confidence, limitations, raw_reference) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'observed', ?, ?, ?)",
                    (etsy_source_id, run_id, observed_at, result["query"], market, language, metric_name,
                     str(metric_value), unit, 90, limitations, result["source_url"]),
                )
            if result["digital_price_median"] is not None:
                connection.execute(
                    "INSERT INTO evidence (source_id, research_run_id, observed_at, query, market, language, "
                    "metric_name, metric_value, unit, evidence_type, confidence, limitations, raw_reference) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'observed', ?, ?, ?)",
                    (etsy_source_id, run_id, observed_at, result["query"], market, language,
                     "median digital competitor price", str(result["digital_price_median"]), "listing_currency_price",
                     75, "Median of sampled digital listings; prices may use mixed currencies.", result["source_url"]),
                )
        connection.execute(
            "UPDATE research_runs SET completed_at = ?, status = 'completed', summary = ? WHERE id = ?",
            (observed_at, f"Etsy supply queries={len(etsy_report['results'])}; demand metrics unavailable", run_id),
        )
    payload = {
        "research_run_id": run_id,
        "generated_at": observed_at,
        "market": market,
        "language": language,
        "etsy_supply_scan": etsy_report,
        "status": "needs_human_validation",
        "next_step": "Contrasta oferta observada con pruebas de demanda e intención de compra antes de aceptar o prototipar.",
    }
    if report_directory is not None:
        report_directory.mkdir(parents=True, exist_ok=True)
        report_path = report_directory / f"market-discovery-{run_id}.json"
        payload["report_path"] = str(report_path)
        report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    record_audit_event(
        database_path,
        "market_discovery_completed",
        actor="system",
        entity_type="research_run",
        entity_id=str(run_id),
        details_redacted=f"etsy_queries={len(etsy_report['results'])}; demand_metrics=none",
    )
    return payload


def run_market_discovery(
    database_path: Path,
    seed_keywords: list[str],
    etsy_values: dict[str, str],
    *,
    shop_id: str | int | None,
    market: str,
    language: str,
    report_directory: Path | None = None,
) -> dict:
    """Scan Etsy supply for selected catalog or user-provided hypotheses."""
    config = load_discovery_config()
    normalized = list(dict.fromkeys(value.strip() for value in seed_keywords if value.strip()))
    if not normalized:
        raise ValueError("Introduce palabras semilla para iniciar la exploración")
    query_limit = int(config["etsy_queries_per_scan"])
    etsy_queries = list(dict.fromkeys(normalized))
    if len(etsy_queries) > query_limit:
        raise ValueError(f"Máximo {query_limit} consultas Etsy por exploración")
    report = scan_etsy_supply(
        etsy_queries,
        etsy_values,
        shop_id=shop_id,
        market=market,
        language=language,
        limit_per_query=int(config["etsy_listings_per_query"]),
    )
    return save_discovery_run(
        database_path,
        queries=etsy_queries,
        market=market,
        language=language,
        etsy_report=report,
        report_directory=report_directory,
    )
