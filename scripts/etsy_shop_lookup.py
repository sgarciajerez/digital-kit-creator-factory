from __future__ import annotations

import argparse
import json
import urllib.parse
import urllib.request
from pathlib import Path

from dotenv import dotenv_values


PROJECT_ROOT = Path(__file__).resolve().parent.parent


def find_shops(shop_name: str, env_path: Path = PROJECT_ROOT / ".env") -> list[dict]:
    values = dotenv_values(env_path)
    key = (values.get("ETSY_API_KEY") or "").strip()
    secret = (values.get("ETSY_API_SECRET") or "").strip()
    if not key or not secret:
        raise RuntimeError("Faltan ETSY_API_KEY o ETSY_API_SECRET en .env")
    query = urllib.parse.urlencode({"shop_name": shop_name, "limit": 25})
    request = urllib.request.Request(
        f"https://api.etsy.com/v3/application/shops?{query}",
        headers={"x-api-key": f"{key}:{secret}", "Accept": "application/json"},
    )
    with urllib.request.urlopen(request, timeout=20) as response:
        payload = json.loads(response.read().decode("utf-8"))
    return list(payload.get("results", []))


def write_shop_id(env_path: Path, shop_id: int) -> None:
    text = env_path.read_text(encoding="utf-8") if env_path.exists() else ""
    lines = text.splitlines()
    replacement = f"ETSY_SHOP_ID={shop_id}"
    for index, line in enumerate(lines):
        if line.startswith("ETSY_SHOP_ID="):
            lines[index] = replacement
            break
    else:
        lines.append(replacement)
    env_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Busca el shop_id de una tienda Etsy por nombre")
    parser.add_argument("shop_name", nargs="?", default="maquetasexcel")
    parser.add_argument("--write-env", action="store_true", help="Escribe la coincidencia exacta en .env")
    args = parser.parse_args()
    results = find_shops(args.shop_name)
    print(f"Resultados encontrados: {len(results)}")
    for result in results:
        print(f"shop_id={result.get('shop_id')} shop_name={result.get('shop_name')}")
    exact = [row for row in results if str(row.get("shop_name", "")).casefold() == args.shop_name.casefold()]
    if args.write_env:
        if len(exact) != 1:
            raise RuntimeError("No hay una única coincidencia exacta; no se modifica .env")
        shop_id = int(exact[0]["shop_id"])
        write_shop_id(PROJECT_ROOT / ".env", shop_id)
        print("ETSY_SHOP_ID escrito en .env para la coincidencia exacta")


if __name__ == "__main__":
    main()
