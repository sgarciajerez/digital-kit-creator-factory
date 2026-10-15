from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

import market_discovery
from factory_core import initialize_database, managed_connection, scalar


def test_catalog_returns_localized_seed_hypotheses() -> None:
    english = market_discovery.discovery_catalog("en")
    spanish = market_discovery.discovery_catalog("es")

    assert len(english) == 24
    assert len(spanish) == 24
    assert english[0]["query"] == "indoor plant care planner"
    assert spanish[0]["query"] == "planificador cuidado plantas interior"
    with pytest.raises(ValueError, match="No hay catálogo"):
        market_discovery.discovery_catalog("fr")


def test_catalog_prioritizes_topics_not_previously_scanned(tmp_path: Path) -> None:
    database_path = tmp_path / "factory.sqlite3"
    initialize_database(database_path)
    initial = market_discovery.unscanned_catalog("en", database_path)
    assert len(initial) == 24

    run_id, _ = market_discovery._insert_run(database_path, [initial[0]["query"]], "US", "en")
    with managed_connection(database_path) as connection:
        connection.execute("UPDATE research_runs SET status = 'completed' WHERE id = ?", (run_id,))

    remaining = market_discovery.unscanned_catalog("en", database_path)
    assert len(remaining) == 23
    assert initial[0]["query"] not in {item["query"] for item in remaining}


def test_etsy_scan_counts_supply_and_excludes_own_shop(monkeypatch) -> None:
    def fake_request(method, path, values):
        assert method == "GET"
        if "small+garden" in path:
            return {
                "count": 100,
                "results": [
                    {"listing_id": 1, "shop_id": 10, "listing_type": "download", "title": "Small garden planner", "url": "https://etsy.test/1", "price": {"amount": 800, "divisor": 100, "currency_code": "USD"}},
                    {"listing_id": 2, "shop_id": 20, "listing_type": "download", "title": "Small garden planner", "url": "https://etsy.test/2", "price": {"amount": 900, "divisor": 100, "currency_code": "USD"}},
                    {"listing_id": 3, "shop_id": 21, "listing_type": "physical", "title": "Garden tools", "url": "https://etsy.test/3"},
                ],
            }
        return {"count": 40, "results": [
            {"listing_id": 2, "shop_id": 20, "listing_type": "download", "title": "Small garden planner", "url": "https://etsy.test/2", "price": {"amount": 900, "divisor": 100, "currency_code": "USD"}},
            {"listing_id": 4, "shop_id": 22, "listing_type": "download", "title": "Small garden budget checklist", "url": "https://etsy.test/4", "price": {"amount": 700, "divisor": 100, "currency_code": "USD"}},
        ]}

    monkeypatch.setattr(market_discovery, "request_json", fake_request)
    report = market_discovery.scan_etsy_supply(
        ["small garden", "balcony garden"],
        {"ETSY_API_KEY": "test", "ETSY_API_SECRET": "test"},
        shop_id=10,
    )

    assert report["unique_digital_listings"] == 2
    assert report["results"][0]["total_active_results"] == 100
    assert report["results"][0]["sampled_digital_competitors"] == 1
    assert report["supply_phrase_clusters"]
    assert report["supply_phrase_clusters"][0]["signal_type"] == "competition_phrase_only"
    assert "not buyer search volume" in report["limitations"][0]


def test_discovery_persists_supply_evidence_without_creating_opportunity(tmp_path: Path) -> None:
    database_path = tmp_path / "factory.sqlite3"
    initialize_database(database_path)
    etsy_report = {
        "marketplace": "Etsy",
        "market": "US",
        "language": "en",
        "results": [{
            "query": "balcony vegetable garden",
            "total_active_results": 200,
            "sampled_digital_competitors": 18,
            "digital_price_median": 7.5,
            "source_url": "https://www.etsy.com/search?q=balcony+vegetable+garden",
        }],
        "supply_phrase_clusters": [],
        "limitations": [],
    }

    payload = market_discovery.save_discovery_run(
        database_path,
        queries=["balcony vegetable garden"],
        market="US",
        language="en",
        etsy_report=etsy_report,
        report_directory=tmp_path / "reports",
    )

    with sqlite3.connect(database_path) as connection:
        types = {row[0] for row in connection.execute("SELECT DISTINCT evidence_type FROM evidence")}
    assert payload["status"] == "needs_human_validation"
    assert "google_keyword_ideas" not in payload
    assert Path(payload["report_path"]).exists()
    assert types == {"observed"}
    assert scalar(database_path, "SELECT COUNT(*) FROM opportunities") == 0
    assert scalar(database_path, "SELECT status FROM research_runs WHERE id = ?", (payload["research_run_id"],)) == "completed"


def test_run_discovery_scans_selected_catalog_queries_and_enforces_limit(tmp_path: Path, monkeypatch) -> None:
    captured = {}

    def fake_scan(queries, api_values, **kwargs):
        captured["queries"] = queries
        return {
            "marketplace": "Etsy", "market": kwargs["market"], "language": kwargs["language"],
            "results": [], "supply_phrase_clusters": [], "limitations": [], "unique_digital_listings": 0,
        }

    monkeypatch.setattr(market_discovery, "scan_etsy_supply", fake_scan)
    payload = market_discovery.run_market_discovery(
        tmp_path / "factory.sqlite3", [" indoor plants ", "indoor plants"],
        {"ETSY_API_KEY": "test"}, shop_id=None, market="US", language="en",
        report_directory=tmp_path / "reports",
    )
    assert captured["queries"] == ["indoor plants"]
    assert payload["status"] == "needs_human_validation"

    with pytest.raises(ValueError, match="Máximo"):
        market_discovery.run_market_discovery(
            tmp_path / "factory.sqlite3", [f"query {index}" for index in range(9)],
            {"ETSY_API_KEY": "test"}, shop_id=None, market="US", language="en",
        )
