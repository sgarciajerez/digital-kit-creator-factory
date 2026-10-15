from __future__ import annotations

"""Evidence-gated, portfolio-aware recommendations for the next digital kit."""

import json
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

from factory_core import PROJECT_ROOT, initialize_database, managed_connection, now_utc, record_audit_event


CONFIG_PATH = PROJECT_ROOT / "config" / "product_selection.json"
DEFAULT_MAX_PRODUCTS = 3


@dataclass(frozen=True)
class ProductSelectionReport:
    candidates: tuple[dict, ...]
    selected_opportunity_ids: tuple[int, ...]
    report_path: Path | None = None


def load_selection_config(path: Path = CONFIG_PATH) -> dict:
    config = json.loads(path.read_text(encoding="utf-8"))
    if config.get("schema_version") != "1.0.0":
        raise ValueError("Versión de config/product_selection.json no compatible")
    return config


def _tokens(value: str, stopwords: set[str]) -> set[str]:
    return {
        token
        for token in re.findall(r"[a-z0-9áéíóúüñ]+", value.casefold())
        if len(token) >= 4 and token not in stopwords
    }


def _similarity(left: str, right: str, stopwords: set[str]) -> float:
    left_tokens = _tokens(left, stopwords)
    right_tokens = _tokens(right, stopwords)
    if not left_tokens or not right_tokens:
        return 0.0
    return len(left_tokens & right_tokens) / len(left_tokens | right_tokens)


def _catalog_text(product: dict) -> str:
    values = [
        product.get("title", product.get("name", "")),
        product.get("slug", ""),
        product.get("target_buyer", product.get("target_audience", "")),
        product.get("subtitle", ""),
        " ".join(product.get("features", [])) if isinstance(product.get("features"), list) else "",
        product.get("problem", ""),
        product.get("category", ""),
        product.get("subcategory", ""),
    ]
    return " ".join(str(value) for value in values if value)


def load_existing_product_catalog(project_root: Path = PROJECT_ROOT) -> list[dict]:
    """Use the factory's declared products plus generated product folders as the portfolio."""
    from create_next_product_designs import PRODUCTS
    from factory_config import active_product_ids, finalized_product_ids

    known_ids = set(active_product_ids()) | set(finalized_product_ids())
    output_root = project_root / "outputs" / "listings"
    for folder in output_root.glob("product-*"):
        match = re.fullmatch(r"product-(\d+)", folder.name)
        if match:
            known_ids.add(int(match.group(1)))
    catalog = []
    for product in PRODUCTS:
        product_id = int(product["id"])
        if product_id not in known_ids:
            continue
        catalog.append({**product, "id": product_id})
    return catalog


def _parse_observed_date(value: str) -> date | None:
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).date()
    except (TypeError, ValueError):
        return None


def _source_key(row: dict) -> str | None:
    source_url = str(row.get("source_url") or "").strip()
    raw_reference = str(row.get("raw_reference") or "").strip()
    for candidate in (source_url, raw_reference):
        parsed = urlparse(candidate)
        if parsed.netloc:
            host = parsed.netloc.casefold().removeprefix("www.")
            # Etsy demand and supply endpoints are different signals from one platform,
            # not independent publishers for the minimum-source policy.
            if host == "etsy.com" or host.endswith(".etsy.com"):
                return "etsy.com"
            return host
    name = str(row.get("source_name") or "").strip()
    return name.casefold() if name else None


def _classify_metric(metric_name: str, terms: dict) -> str | None:
    normalized = re.sub(r"[_-]+", " ", metric_name.casefold())
    # Search result/listing counts describe supply, not buyer demand.
    if any(term in normalized for term in terms.get("competition_terms", [])):
        return "competition"
    if any(term in normalized for term in terms.get("demand_terms", [])):
        return "demand"
    return None


def _evidence_for_opportunities(connection, terms: dict, freshness_days: int) -> dict[int, dict]:
    rows = connection.execute(
        "SELECT e.*, s.name AS source_name, s.url AS source_url "
        "FROM evidence e LEFT JOIN sources s ON s.id = e.source_id "
        "ORDER BY e.observed_at DESC, e.id DESC"
    ).fetchall()
    today = date.today()
    grouped: dict[int, set[int]] = defaultdict(set)
    links = connection.execute(
        "SELECT DISTINCT opportunity_id, evidence_ids FROM score_components"
    ).fetchall()
    for link in links:
        try:
            ids = json.loads(link["evidence_ids"] or "[]")
        except (TypeError, json.JSONDecodeError):
            ids = []
        for evidence_id in ids:
            grouped[int(link["opportunity_id"])].add(int(evidence_id))
    direct_links = connection.execute(
        "SELECT opportunity_id, evidence_id FROM opportunity_evidence"
    ).fetchall()
    for link in direct_links:
        grouped[int(link["opportunity_id"])].add(int(link["evidence_id"]))

    by_id = {int(row["id"]): dict(row) for row in rows}
    summaries: dict[int, dict] = {}
    for opportunity_id, references in grouped.items():
        observations = []
        for evidence_id in sorted(references):
            row = by_id.get(evidence_id)
            if not row:
                continue
            observed_date = _parse_observed_date(row.get("observed_at", ""))
            age_days = (today - observed_date).days if observed_date else None
            classification = _classify_metric(str(row.get("metric_name") or ""), terms)
            row["age_days"] = age_days
            row["classification"] = classification
            row["source_key"] = _source_key(row)
            observations.append(row)

        fresh_observed = [
            row for row in observations
            if row.get("evidence_type") == "observed"
            and row.get("age_days") is not None
            and 0 <= row["age_days"] <= freshness_days
            and row.get("source_key")
            and row.get("raw_reference")
        ]
        source_count = len({row["source_key"] for row in fresh_observed})
        summaries[opportunity_id] = {
            "evidence_count": len(fresh_observed),
            "source_count": source_count,
            "demand_count": sum(row.get("classification") == "demand" for row in fresh_observed),
            "competition_count": sum(row.get("classification") == "competition" for row in fresh_observed),
            "oldest_days": max((row["age_days"] for row in fresh_observed), default=None),
            "evidence_ids": [int(row["id"]) for row in fresh_observed],
            "references": [
                {
                    "evidence_id": int(row["id"]),
                    "observed_at": row.get("observed_at"),
                    "query": row.get("query"),
                    "metric_name": row.get("metric_name"),
                    "metric_value": row.get("metric_value"),
                    "unit": row.get("unit"),
                    "market": row.get("market"),
                    "language": row.get("language"),
                    "evidence_type": row.get("evidence_type"),
                    "confidence": row.get("confidence"),
                    "source": row.get("source_name"),
                    "reference": row.get("raw_reference") or row.get("source_url"),
                    "limitations": row.get("limitations"),
                }
                for row in fresh_observed
            ],
            "evidence_types": sorted({str(row["evidence_type"]) for row in observations}),
        }
    return summaries


def _category_shop_signals(connection, minimum_visits: int, minimum_orders: int) -> dict[str, dict]:
    rows = connection.execute(
        "SELECT lower(trim(coalesce(o.category, o.subcategory, ''))) AS category, "
        "m.visits, m.orders, m.gross_revenue, m.refunds, m.fees, m.ad_spend, "
        "m.data_quality_notes "
        "FROM metrics_daily m "
        "JOIN listings l ON l.id = m.listing_id "
        "JOIN products p ON p.id = l.product_id "
        "LEFT JOIN opportunities o ON o.id = p.opportunity_id "
        "WHERE lower(l.status) IN ('active', 'published')"
    ).fetchall()
    totals: dict[str, dict] = defaultdict(lambda: {"visits": 0, "orders": 0, "margin": 0.0})
    for row in rows:
        notes = str(row["data_quality_notes"] or "").casefold()
        if any(marker in notes for marker in ("fictional", "test_data", "fixture", "sample only")):
            continue
        category = str(row["category"] or "").strip()
        if not category:
            continue
        totals[category]["visits"] += int(row["visits"] or 0)
        totals[category]["orders"] += int(row["orders"] or 0)
        totals[category]["margin"] += float(row["gross_revenue"] or 0) - sum(
            float(row[field] or 0) for field in ("refunds", "fees", "ad_spend")
        )
    return {
        category: {
            **values,
            "conversion_rate": values["orders"] / values["visits"] if values["visits"] else None,
            "usable": values["visits"] >= minimum_visits and values["orders"] >= minimum_orders,
        }
        for category, values in totals.items()
    }


def _candidate_performance_signal(category: str, signals: dict[str, dict]) -> dict | None:
    normalized = category.strip().casefold()
    signal = signals.get(normalized)
    if not signal or not signal["usable"]:
        return None
    usable = [item["conversion_rate"] for item in signals.values() if item["usable"] and item["conversion_rate"] is not None]
    if not usable:
        return None
    low, high = min(usable), max(usable)
    percentile = 50.0 if high == low else 100 * (signal["conversion_rate"] - low) / (high - low)
    return {**signal, "relative_score": round(percentile, 2)}


def _confidence(evidence: dict) -> float:
    # A small, transparent confidence heuristic; it is deliberately separate from the commercial score.
    score = min(40.0, evidence["evidence_count"] * 8.0)
    score += min(30.0, evidence["source_count"] * 15.0)
    score += 15.0 if evidence["demand_count"] else 0.0
    score += 15.0 if evidence["competition_count"] else 0.0
    if evidence["oldest_days"] is not None:
        score -= min(20.0, max(0, evidence["oldest_days"] - 30) / 10)
    return round(max(0.0, min(100.0, score)), 2)


def _get_catalog_similarity(opportunity: dict, catalog: list[dict], stopwords: set[str]) -> tuple[float, dict | None]:
    candidate_text = " ".join(
        str(opportunity.get(key) or "")
        for key in ("name", "category", "subcategory", "target_audience", "problem", "proposed_solution")
    )
    best_score, best_product = 0.0, None
    for product in catalog:
        product_title = str(product.get("title", product.get("name", "")))
        name_tokens = _tokens(str(opportunity.get("name") or ""), stopwords)
        title_tokens = _tokens(product_title, stopwords)
        name_overlap = (
            len(name_tokens & title_tokens) / min(len(name_tokens), len(title_tokens))
            if name_tokens and title_tokens
            else 0.0
        )
        score = max(
            _similarity(candidate_text, _catalog_text(product), stopwords),
            _similarity(str(opportunity.get("name") or ""), product_title, stopwords),
            name_overlap,
        )
        if score > best_score:
            best_score, best_product = score, product
    return best_score, best_product


def _get_candidate_rows(connection) -> list[dict]:
    rows = connection.execute(
        "SELECT o.* FROM opportunities o WHERE o.decision <> 'reject' "
        "AND NOT EXISTS (SELECT 1 FROM products p WHERE p.opportunity_id = o.id) "
        "ORDER BY o.updated_at DESC, o.id ASC"
    ).fetchall()
    return [dict(row) for row in rows]


def recommend_products(
    database_path: Path,
    *,
    max_products: int = DEFAULT_MAX_PRODUCTS,
    catalog: list[dict] | None = None,
    config: dict | None = None,
    report_path: Path | None = None,
) -> ProductSelectionReport:
    """Rank researched opportunities, with explicit evidence and portfolio-diversity gates."""
    if max_products < 1:
        raise ValueError("max_products debe ser positivo")
    config = config or load_selection_config()
    policy = config["policy"]
    ranking = config["ranking"]
    catalog = load_existing_product_catalog() if catalog is None else catalog
    stopwords = set(ranking.get("catalog_stopwords", []))
    initialize_database(database_path)

    with managed_connection(database_path) as connection:
        opportunities = _get_candidate_rows(connection)
        evidence_summaries = _evidence_for_opportunities(
            connection,
            config["evidence_classification"],
            int(policy["evidence_freshness_days"]),
        )
        shop_signals = _category_shop_signals(
            connection,
            int(policy["shop_signal_minimum_visits"]),
            int(policy["shop_signal_minimum_orders"]),
        )

    candidates = []
    for opportunity in opportunities:
        opportunity_id = int(opportunity["id"])
        evidence = evidence_summaries.get(opportunity_id, {
            "evidence_count": 0, "source_count": 0, "demand_count": 0,
            "competition_count": 0, "oldest_days": None, "evidence_ids": [], "references": [], "evidence_types": [],
        })
        confidence = _confidence(evidence)
        similarity, closest = _get_catalog_similarity(opportunity, catalog, stopwords)
        status_reasons = []
        if opportunity.get("decision") != "accept":
            status_reasons.append("La oportunidad no está aceptada.")
        if float(opportunity.get("score") or 0) < float(policy["minimum_opportunity_score"]):
            status_reasons.append("No alcanza la puntuación comercial mínima.")
        if confidence < float(policy["minimum_confidence"]):
            status_reasons.append("Confianza de evidencia insuficiente.")
        if evidence["evidence_count"] < int(policy["minimum_observed_evidence"]):
            status_reasons.append("Faltan observaciones de mercado recientes y verificables.")
        if evidence["source_count"] < int(policy["minimum_independent_sources"]):
            status_reasons.append("Faltan fuentes independientes.")
        if evidence["demand_count"] < 1:
            status_reasons.append("Falta evidencia observada de demanda; resultados de búsqueda no bastan.")
        if evidence["competition_count"] < 1:
            status_reasons.append("Falta evidencia observada de competencia/precios.")
        if similarity > float(policy["maximum_catalog_similarity"]):
            similar_name = (closest or {}).get("title") or (closest or {}).get("name") or "producto existente"
            status_reasons.append(f"Solapamiento alto con el catálogo ({similarity:.2f}): {similar_name}.")

        novelty_score = round(max(0.0, 100 * (1 - similarity)), 2)
        base_score = float(opportunity.get("score") or 0)
        performance = _candidate_performance_signal(str(opportunity.get("category") or ""), shop_signals)
        components = [(base_score, float(ranking["opportunity_score_weight"])), (novelty_score, float(ranking["novelty_weight"]))]
        if performance is not None:
            components.append((performance["relative_score"], float(ranking["shop_performance_weight"])))
        denominator = sum(weight for _, weight in components)
        selection_score = round(sum(value * weight for value, weight in components) / denominator, 2) if denominator else 0.0
        candidates.append({
            "opportunity_id": opportunity_id,
            "name": opportunity["name"],
            "category": opportunity.get("category") or "Sin categoría",
            "subcategory": opportunity.get("subcategory") or "",
            "decision": opportunity.get("decision"),
            "base_score": base_score,
            "confidence": confidence,
            "novelty_score": novelty_score,
            "catalog_similarity": round(similarity, 3),
            "closest_catalog_product": (closest or {}).get("title") or (closest or {}).get("name"),
            "selection_score": selection_score,
            "evidence": evidence,
            "shop_signal": performance,
            "status": "eligible" if not status_reasons else "blocked",
            "reasons": status_reasons,
        })

    eligible = sorted(
        (candidate for candidate in candidates if candidate["status"] == "eligible"),
        key=lambda candidate: (-candidate["selection_score"], -candidate["confidence"], candidate["opportunity_id"]),
    )
    selected = []
    category_counts: dict[str, int] = defaultdict(int)
    max_per_category = int(policy["maximum_per_category"])
    for candidate in eligible:
        category = candidate["category"].strip().casefold()
        if category_counts[category] >= max_per_category:
            candidate["status"] = "deferred_for_diversity"
            candidate["reasons"].append("Se pospone para diversificar por categoría en este lote.")
            continue
        candidate["status"] = "recommended"
        selected.append(candidate)
        category_counts[category] += 1
        if len(selected) >= max_products:
            break

    candidates.sort(key=lambda candidate: (
        0 if candidate["status"] == "recommended" else 1,
        -candidate["selection_score"],
        candidate["opportunity_id"],
    ))
    report = ProductSelectionReport(
        candidates=tuple(candidates),
        selected_opportunity_ids=tuple(candidate["opportunity_id"] for candidate in selected),
        report_path=report_path,
    )
    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "generated_at": now_utc(),
            "method": "evidence_gated_portfolio_diversified_ranking",
            "selected_opportunity_ids": list(report.selected_opportunity_ids),
            "candidates": candidates,
            "limitations": [
                "Etsy search-result counts are treated as competition, not proof of buyer demand.",
                "Shop-performance signals are used only when category-linked metrics meet configured minimums and are not marked as test data.",
                "Similarity is a transparent lexical screen; it flags candidates for review and is not semantic proof of duplication.",
            ],
        }
        report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        record_audit_event(
            database_path,
            "product_selection_recommended",
            actor="system",
            entity_type="opportunity_batch",
            entity_id=",".join(map(str, report.selected_opportunity_ids)),
            details_redacted=f"candidates={len(candidates)}; selected={len(selected)}; report={report_path.name}",
        )
    return report


if __name__ == "__main__":
    from factory_core import DATABASE_PATH

    output = PROJECT_ROOT / "outputs" / "reports" / "product_selection_report.json"
    result = recommend_products(DATABASE_PATH, report_path=output)
    print(f"Candidatos analizados: {len(result.candidates)}")
    print(f"Productos recomendados: {len(result.selected_opportunity_ids)}")
    print(f"Informe: {output}")
