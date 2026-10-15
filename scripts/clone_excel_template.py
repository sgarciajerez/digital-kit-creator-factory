from __future__ import annotations

import argparse
import shutil
from pathlib import Path

import win32com.client as win32

from design_architecture import ROOT
from model_quality import validate_model_file
from model_contracts import family_base_contract, load_model_contract, write_model_behavior_markdown, write_model_contract_json


DEFAULT_TEMPLATE = ROOT / "templates" / "excel-first-base"


def main() -> int:
    parser = argparse.ArgumentParser(description="Clona una plantilla maestra Excel sin rediseñar el libro")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Carpeta de la plantilla maestra")
    parser.add_argument("--output", type=Path, required=True, help="Ruta del nuevo XLSX")
    parser.add_argument("--title", help="Título opcional para la cabecera de todas las hojas")
    parser.add_argument("--subtitle", help="Subtítulo opcional para la cabecera de todas las hojas")
    parser.add_argument("--icon", type=Path, help="SVG de identidad del producto; se incrusta en todas las hojas")
    args = parser.parse_args()

    template_dir = args.template if args.template.is_absolute() else ROOT / args.template
    output = args.output if args.output.is_absolute() else ROOT / args.output
    source = template_dir / "master_workbook.xlsx"
    if not source.is_file():
        raise SystemExit(f"No existe master_workbook.xlsx: {source}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output)
    if args.title or args.subtitle:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = None
        try:
            workbook = excel.Workbooks.Open(str(output), UpdateLinks=0, ReadOnly=False, AddToMru=False)
            for sheet in workbook.Worksheets:
                if args.title:
                    sheet.Range("B3").Value = args.title
                if args.subtitle:
                    sheet.Range("B4").Value = args.subtitle
            workbook.Save()
            workbook.Close(SaveChanges=True)
            workbook = None
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            excel.Quit()
    if args.icon:
        icon = args.icon if args.icon.is_absolute() else ROOT / args.icon
        if not icon.is_file():
            raise SystemExit(f"No existe el SVG de identidad: {icon}")
        from embed_svg_icon import embed_icon
        embed_icon(output, output, icon)
    contract_path = template_dir / "model_contract.json"
    model_contract = load_model_contract(contract_path) if contract_path.is_file() else None
    if model_contract is None and source.name == "master_workbook.xlsx":
        model_contract = family_base_contract()
    behavior_path = output.parent.parent / "model_behavior.md"
    if model_contract:
        write_model_behavior_markdown(behavior_path, model_contract)
        write_model_contract_json(output.parent.parent / "model_contract.json", model_contract)
    quality = validate_model_file(
        output,
        expected_sheets=model_contract.get("sheets") if model_contract else None,
        model_contract=model_contract,
        model_contract_path=behavior_path if model_contract else None,
    )
    if not quality["passed"]:
        raise SystemExit(f"El clon no supera la validación global: {quality['errors']}")
    print(f"TEMPLATE_CLONED={output}")
    print(f"SOURCE_TEMPLATE={source}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
