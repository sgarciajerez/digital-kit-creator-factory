from __future__ import annotations

"""Central runtime configuration for the Excel factory pipeline."""

import json
from functools import lru_cache
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
PIPELINE_PATH = ROOT / "config" / "FACTORY_PIPELINE.json"


@lru_cache(maxsize=1)
def load_pipeline_config() -> dict:
    return json.loads(PIPELINE_PATH.read_text(encoding="utf-8"))


def active_product_ids() -> tuple[int, ...]:
    return tuple(int(value) for value in load_pipeline_config().get("active_product_ids", []))


def finalized_product_ids() -> tuple[int, ...]:
    return tuple(int(value) for value in load_pipeline_config().get("finalized_product_ids", []))


def all_product_ids() -> tuple[int, ...]:
    configured = active_product_ids() + finalized_product_ids()
    return tuple(dict.fromkeys(configured))


def product_ids(include_finalized: bool = False) -> tuple[int, ...]:
    return all_product_ids() if include_finalized else active_product_ids()
