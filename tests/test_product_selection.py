from __future__ import annotations

from datetime import date
from pathlib import Path

from factory_core import create_opportunity, create_product, managed_connection, now_utc
from product_selection import recommend_products
from research import ScoreComponentInput, import_evidence_csv, list_evidence, score_opportunity


def _researched_opportunity(tmp_path: Path, database_path: Path, name: str, category: str, score: float = 90) -> int:
    opportunity_id = create_opportunity(
        database_path,
        name=name,
        category=category,
        subcategory="Specific tools",
        target_audience=f"People who need {name.lower()}",
        problem=f"They need a reliable way to manage {name.lower()} decisions",
        proposed_solution=f"A focused digital kit for {name.lower()}",
        score=score,
        confidence=90,
        decision="accept",
    )
    csv_path = tmp_path / f"evidence-{opportunity_id}.csv"
    csv_path.write_text(
        "observed_at,query,market,language,metric_name,metric_value,unit,evidence_type,"
        "confidence,limitations,raw_reference\n"
        f"{date.today().isoformat()},{name},US,en,monthly searches,1200,count,observed,90,Keyword tool estimate,https://keywords.example/{opportunity_id}\n"
        f"{date.today().isoformat()},{name},US,en,active listings,450,count,observed,85,Search results show supply,https://etsy.example/{opportunity_id}\n"
        f"{date.today().isoformat()},{name},US,en,buyer interviews,4,count,observed,70,Small exploratory sample,https://interviews.example/{opportunity_id}\n",
        encoding="utf-8",
    )
    imported = import_evidence_csv(database_path, csv_path, f"market sources {opportunity_id}")
    evidence_ids = tuple(row["id"] for row in list_evidence(database_path, imported.research_run_id))
    score_opportunity(
        database_path,
        opportunity_id,
        [ScoreComponentInput("evidencia de demanda", 100, 5, "Evidence attached", evidence_ids)],
        confidence=90,
    )
    return opportunity_id


def test_high_score_without_market_evidence_is_blocked(tmp_path: Path) -> None:
    database_path = tmp_path / "factory.sqlite3"
    opportunity_id = create_opportunity(
        database_path,
        "Popular planner idea",
        decision="accept",
        score=99,
        confidence=99,
    )

    report = recommend_products(database_path, catalog=[])

    assert report.selected_opportunity_ids == ()
    candidate = next(item for item in report.candidates if item["opportunity_id"] == opportunity_id)
    assert candidate["status"] == "blocked"
    assert any("observaciones de mercado" in reason for reason in candidate["reasons"])
    assert any("demanda" in reason for reason in candidate["reasons"])


def test_recent_independent_demand_and_competition_evidence_recommend_candidate(tmp_path: Path) -> None:
    database_path = tmp_path / "factory.sqlite3"
    opportunity_id = _researched_opportunity(
        tmp_path, database_path, "Freshwater tank volume calculator", "aquarium tools"
    )

    report = recommend_products(database_path, catalog=[])

    assert report.selected_opportunity_ids == (opportunity_id,)
    candidate = report.candidates[0]
    assert candidate["status"] == "recommended"
    assert candidate["evidence"]["demand_count"] >= 1
    assert candidate["evidence"]["competition_count"] >= 1
    assert candidate["evidence"]["source_count"] == 3
    assert all(item["reference"].startswith("https://") for item in candidate["evidence"]["references"])
    assert candidate["confidence"] >= 60


def test_catalog_similarity_blocks_repackaging_existing_product(tmp_path: Path) -> None:
    database_path = tmp_path / "factory.sqlite3"
    opportunity_id = _researched_opportunity(
        tmp_path, database_path, "Freshwater aquarium setup planner", "aquarium planning"
    )
    catalog = [{
        "title": "Freshwater Aquarium Setup Planner",
        "target_buyer": "People planning a freshwater aquarium setup",
        "subtitle": "A practical freshwater aquarium setup planning kit",
        "features": ["Freshwater aquarium setup budget", "Aquarium setup timeline"],
    }]

    report = recommend_products(database_path, catalog=catalog)

    assert report.selected_opportunity_ids == ()
    candidate = next(item for item in report.candidates if item["opportunity_id"] == opportunity_id)
    assert candidate["status"] == "blocked"
    assert candidate["catalog_similarity"] > 0.62
    assert any("Solapamiento alto" in reason for reason in candidate["reasons"])


def test_batch_selection_limits_category_repetition(tmp_path: Path) -> None:
    database_path = tmp_path / "factory.sqlite3"
    first_id = _researched_opportunity(tmp_path, database_path, "Aquarium volume helper", "aquarium")
    second_id = _researched_opportunity(tmp_path, database_path, "Aquarium maintenance organizer", "aquarium")
    third_id = _researched_opportunity(tmp_path, database_path, "Wedding seating assistant", "weddings")

    report = recommend_products(database_path, max_products=3, catalog=[])

    assert len(report.selected_opportunity_ids) == 2
    assert third_id in report.selected_opportunity_ids
    assert not ({first_id, second_id} <= set(report.selected_opportunity_ids))
    deferred = [row for row in report.candidates if row["status"] == "deferred_for_diversity"]
    assert len(deferred) == 1


def test_category_shop_results_influence_rank_only_with_usable_data(tmp_path: Path) -> None:
    database_path = tmp_path / "factory.sqlite3"
    aquarium_id = _researched_opportunity(tmp_path, database_path, "Tank volume helper", "aquarium")
    wedding_id = _researched_opportunity(tmp_path, database_path, "Wedding guest seating helper", "weddings")

    for category, visits, orders in (("aquarium", 100, 3), ("weddings", 100, 20)):
        historical_opportunity = create_opportunity(
            database_path,
            f"Historical {category} product",
            category=category,
            decision="accept",
            score=80,
            confidence=80,
        )
        product_id = create_product(
            database_path, f"Historical {category} product", "digital_kit", historical_opportunity
        )
        with managed_connection(database_path) as connection:
            listing_cursor = connection.execute(
                "INSERT INTO listings (product_id, platform, title, status, created_at) "
                "VALUES (?, 'etsy', ?, 'active', ?)",
                (product_id, f"Historical {category} product", now_utc()),
            )
            connection.execute(
                "INSERT INTO metrics_daily (listing_id, date, visits, orders, units_sold, gross_revenue, "
                "refunds, fees, ad_spend, source, data_quality_notes) "
                "VALUES (?, ?, ?, ?, ?, ?, 0, 0, 0, 'etsy_csv', '')",
                (int(listing_cursor.lastrowid), date.today().isoformat(), visits, orders, orders, orders * 8),
            )

    report = recommend_products(database_path, max_products=3, catalog=[])
    by_id = {row["opportunity_id"]: row for row in report.candidates}

    assert by_id[aquarium_id]["shop_signal"]["relative_score"] == 0
    assert by_id[wedding_id]["shop_signal"]["relative_score"] == 100
    assert by_id[wedding_id]["selection_score"] > by_id[aquarium_id]["selection_score"]
