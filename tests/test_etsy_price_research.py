from __future__ import annotations

import etsy_price_research as research


def test_market_band_reports_middle_half_and_resists_single_extreme():
    band = research.market_band([2, 3, 4, 5, 6, 7, 100])

    assert band == {
        "sample_size": 7,
        "minimum": 2,
        "q1": 3.5,
        "median": 5,
        "q3": 6.5,
        "maximum": 100,
    }


def test_currency_conversion_uses_ecb_base_euro_rates():
    rates = {"EUR": 1.0, "USD": 1.1, "GBP": 0.8}

    assert research.convert_currency(11, "USD", "EUR", rates) == 10
    assert research.convert_currency(10, "EUR", "GBP", rates) == 8


def test_live_etsy_search_keeps_digital_comps_deduplicates_and_excludes_own_shop(monkeypatch):
    def listing(listing_id, shop_id, amount, currency="EUR", listing_type="download"):
        return {
            "listing_id": listing_id,
            "shop_id": shop_id,
            "listing_type": listing_type,
            "title": f"Comparable {listing_id}",
            "url": f"https://www.etsy.com/listing/{listing_id}",
            "price": {"amount": amount * 100, "divisor": 100, "currency_code": currency},
        }

    answers = iter([
        {"count": 3, "results": [listing(1, 11, 5), listing(2, 12, 7), listing(3, 13, 9, listing_type="physical")]},
        {"count": 4, "results": [listing(1, 11, 5), listing(4, 10, 6), listing(5, 14, 8), listing(6, 15, 10)]},
    ])
    monkeypatch.setattr(research, "request_json", lambda *args, **kwargs: next(answers))
    monkeypatch.setattr(research, "fetch_ecb_rates", lambda: ({"EUR": 1.0, "USD": 1.1}, "2026-10-15"))

    report = research.search_etsy_prices(
        ["puppy budget", "dog expense"], {}, shop_id=10, minimum_comparables=4,
    )

    assert report["market_band"]["sample_size"] == 4
    assert report["market_band"]["median"] == 7.5
    assert len(report["observations"]) == 4
    assert all(row["listing_id"] not in {"3", "4"} for row in report["observations"])
    assert report["observations"][0]["price_in_shop_currency"] == 5
