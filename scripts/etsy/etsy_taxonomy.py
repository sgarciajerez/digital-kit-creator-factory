from __future__ import annotations

import argparse
import re
from typing import Any

from etsy_client import load_env, request_json, update_env_value


def flatten(nodes: list[dict[str, Any]], parents: tuple[str, ...] = ()) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for node in nodes:
        name = str(node.get("name", "")).strip()
        path = parents + ((name,) if name else ())
        output.append({"id": int(node["id"]), "name": name, "path": path})
        output.extend(flatten(node.get("children") or [], path))
    return output


def score(node: dict[str, Any], terms: list[str]) -> int:
    text = " ".join(node["path"]).casefold()
    leaf = node["name"].casefold()
    value = 0
    for term in terms:
        if term in leaf:
            value += 4
        elif term in text:
            value += 1
    if "template" in leaf:
        value += 5
    if "planner" in leaf:
        value += 4
    if leaf == "digital":
        value -= 4
    if "spreadsheet" in leaf:
        value += 5
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description="Busca la taxonomía de vendedor de Etsy")
    parser.add_argument("query", nargs="+", help="Términos de la categoría")
    parser.add_argument("--write-env", action="store_true", help="Guarda solo una coincidencia clara")
    parser.add_argument("--limit", type=int, default=10)
    args = parser.parse_args()

    values = load_env()
    payload = request_json("GET", "/application/seller-taxonomy/nodes", values)
    nodes = flatten(payload.get("results") or [])
    terms = [term.casefold() for term in re.findall(r"[\w-]+", " ".join(args.query), flags=re.UNICODE)]
    ranked = sorted(
        ((score(node, terms), node) for node in nodes),
        key=lambda item: (-item[0], len(item[1]["path"])),
    )
    ranked = [item for item in ranked if item[0] > 0]
    if not ranked:
        print("No se encontraron candidatos.")
        return 2
    print(f"Candidatos: {len(ranked)}")
    for points, node in ranked[: args.limit]:
        print(f"score={points} taxonomy_id={node['id']} path={' > '.join(node['path'])}")

    best_score, best = ranked[0]
    second_score = ranked[1][0] if len(ranked) > 1 else 0
    if args.write_env:
        if best_score < 5 or best_score - second_score < 2:
            print("No se guarda: la coincidencia no es suficientemente clara.")
            return 3
        update_env_value("ETSY_TAXONOMY_ID", str(best["id"]))
        print("ETSY_TAXONOMY_ID guardado para la coincidencia más clara.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

