"""Bind quality decisions to the exact workbook and declarations checked."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path


def sha256(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def artifact_identity(workbook: Path, master: Path, contract: dict, spec: dict) -> dict:
    def digest(value: dict) -> str:
        return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()
    # Runtime keys are not part of the persisted design declaration.
    declaration = {key: value for key, value in spec.items() if not key.startswith("_") and key not in {"model_contract", "functional_contract", "run_functional_qa"}}
    return {"workbook_sha256": sha256(workbook), "master_sha256": sha256(master),
            "contract_sha256": digest(contract), "spec_sha256": digest(declaration)}


def approval_is_current(report: dict, workbook: Path, master: Path, contract: dict, spec: dict) -> bool:
    functional = report.get("validation", {}).get("model_quality", {}).get("functional", {})
    if not report.get("validation", {}).get("passed") or functional.get("status") != "executed" or not functional.get("passed"):
        return False
    try:
        return report.get("artifact_identity") == artifact_identity(workbook, master, contract, spec)
    except OSError:
        return False
