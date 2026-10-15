"""Fresh Etsy-market price sampling and currency-normalized market bands."""
from __future__ import annotations

import math
from datetime import date
from statistics import median
from typing import Any
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from xml.etree import ElementTree

from etsy_client import request_json


ECB_DAILY_RATES_URL = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml"
ECB_RATES_SOURCE = "https://www.ecb.europa.eu/stats/policy_and_exchange_rates/euro_reference_exchange_rates/html/index.en.html"


def quantile(values: list[float], fraction: float) -> float:
    """Linearly interpolated quantile, stable for small comparison sets."""
    if not values or not 0 <= fraction <= 1:
        raise ValueError("Se necesita una muestra no vacía y un percentil entre 0 y 1")
    ordered = sorted(float(value) for value in values)
    position = (len(ordered) - 1) * fraction
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def market_band(values: list[float]) -> dict[str, float | int]:
    """Summarize the observed market using Q1–Q3 and a robust median."""
    clean = [float(value) for value in values if math.isfinite(float(value)) and float(value) > 0]
    if not clean:
        raise ValueError("La investigación no contiene precios válidos")
    return {
        "sample_size": len(clean),
        "minimum": round(min(clean), 2),
        "q1": round(quantile(clean, 0.25), 2),
        "median": round(float(median(clean)), 2),
        "q3": round(quantile(clean, 0.75), 2),
        "maximum": round(max(clean), 2),
    }


def fetch_ecb_rates() -> tuple[dict[str, float], str]:
    request = Request(ECB_DAILY_RATES_URL, headers={"User-Agent": "DigitalKitFactory/1.0"})
    with urlopen(request, timeout=20) as response:
        root = ElementTree.fromstring(response.read())
    rates = {"EUR": 1.0}
    rate_date = date.today().isoformat()
    for node in root.iter():
        if node.attrib.get("time"):
            rate_date = node.attrib["time"]
        currency = node.attrib.get("currency")
        rate = node.attrib.get("rate")
        if currency and rate:
            rates[currency.upper()] = float(rate)
    if len(rates) < 10:
        raise RuntimeError("El BCE devolvió un catálogo de tipos de cambio incompleto")
    return rates, rate_date


def convert_currency(amount: float, currency: str, target_currency: str, rates: dict[str, float]) -> float:
    source = currency.upper()
    target = target_currency.upper()
    if source not in rates or target not in rates:
        raise ValueError(f"No hay tipo de cambio BCE para {source} o {target}")
    # The ECB series quotes units of each currency for one EUR.
    return float(amount) / rates[source] * rates[target]


def search_etsy_prices(
    queries: list[str],
    api_values: dict[str, str],
    *,
    shop_id: str | int | None,
    target_currency: str = "EUR",
    limit_per_query: int = 40,
    minimum_comparables: int = 5,
) -> dict[str, Any]:
    """Search Etsy's active marketplace listings, dedupe, and derive a price band.

    Only digital-download listings are sampled. Own-shop listings and free/invalid
    prices are excluded. Every sample carries its Etsy listing URL for auditing.
    """
    normalized_queries = list(dict.fromkeys(str(query).strip() for query in queries if str(query).strip()))[:3]
    if not normalized_queries:
        raise ValueError("La investigación de precios necesita al menos una búsqueda Etsy")
    if not 1 <= int(limit_per_query) <= 100:
        raise ValueError("limit_per_query debe estar entre 1 y 100")

    listings: dict[str, dict[str, Any]] = {}
    query_counts: list[dict[str, Any]] = []
    for query in normalized_queries:
        path = "/application/listings/active?" + urlencode({
            "keywords": query,
            "limit": int(limit_per_query),
            "sort_on": "score",
            "sort_order": "desc",
        })
        result = request_json("GET", path, api_values)
        rows = result.get("results") or []
        query_counts.append({"query": query, "market_result_count": int(result.get("count", len(rows))), "sampled": len(rows)})
        for row in rows:
            listing_id = str(row.get("listing_id") or "")
            if not listing_id or listing_id in listings:
                continue
            if shop_id is not None and str(row.get("shop_id")) == str(shop_id):
                continue
            if str(row.get("listing_type", "")).casefold() not in {"download", "digital"}:
                continue
            price = row.get("price") or {}
            try:
                amount = float(price["amount"]) / float(price["divisor"])
                currency = str(price["currency_code"]).upper()
            except (KeyError, TypeError, ValueError, ZeroDivisionError):
                continue
            if not math.isfinite(amount) or amount <= 0 or not currency:
                continue
            listings[listing_id] = {
                "listing_id": listing_id,
                "title": str(row.get("title") or ""),
                "price": round(amount, 2),
                "currency": currency,
                "url": str(row.get("url") or f"https://www.etsy.com/listing/{listing_id}"),
                "search_query": query,
            }

    rates, fx_date = fetch_ecb_rates()
    observations = []
    unsupported_currencies: dict[str, int] = {}
    for listing in listings.values():
        try:
            converted = convert_currency(listing["price"], listing["currency"], target_currency, rates)
        except ValueError:
            currency = str(listing["currency"])
            unsupported_currencies[currency] = unsupported_currencies.get(currency, 0) + 1
            continue
        observations.append({**listing, "price_in_shop_currency": round(converted, 2)})
    if len(observations) < minimum_comparables:
        raise ValueError(
            f"Etsy devolvió solo {len(observations)} comparables digitales con moneda convertible "
            f"(mínimo {minimum_comparables}); monedas omitidas: {unsupported_currencies or 'ninguna'}. "
            "Amplía o afina las búsquedas."
        )
    band = market_band([float(row["price_in_shop_currency"]) for row in observations])
    return {
        "marketplace": "Etsy",
        "currency": target_currency.upper(),
        "research_date": date.today().isoformat(),
        "search_queries": query_counts,
        "fx": {"base_currency": "EUR", "rates_date": fx_date, "source": ECB_RATES_SOURCE},
        "excluded_unsupported_currencies": unsupported_currencies,
        "market_band": band,
        "recommended_launch_price": band["median"],
        "launch_position": "market_median",
        "observations": observations,
    }
