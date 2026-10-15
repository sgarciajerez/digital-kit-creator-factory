from __future__ import annotations

"""One-time migration for existing product design specs."""

import argparse
import json

from design_architecture import ROOT, apply_design_architecture


def main() -> int:
    parser = argparse.ArgumentParser(description="Migra specs existentes al sistema de diseño estable")
    parser.add_argument("--product", type=int, help="Migrar un solo producto")
    args = parser.parse_args()
    root = ROOT / "outputs" / "listings"
    paths = [root / f"product-{args.product}" / "design_spec.json"] if args.product else sorted(root.glob("product-*/design_spec.json"))
    for path in paths:
        if not path.is_file():
            continue
        spec = json.loads(path.read_text(encoding="utf-8"))
        apply_design_architecture(spec)
        path.write_text(json.dumps(spec, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"normalized={path.parent.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
