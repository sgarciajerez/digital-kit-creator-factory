"""Behaviour shared by the native project-budget family."""
from __future__ import annotations


def extend_contract(contract: dict) -> dict:
    if contract.get("product_id") not in {1, 12, 14}:
        return contract
    if contract.get("product_id") == 1:
        names = {"Fixtures": "Equipment", "Finishes": "Soil", "Permits": "Workshops", "Design": "Planning", "Room": "Location"}
        for table in contract["tables"]:
            for column in table["headers"]:
                column["name"] = names.get(column["name"], column["name"])
        for policy in contract.get("input_text_policy", []):
            policy["column"] = names.get(policy["column"], policy["column"])
    for rule in contract["formula_rules"]:
        if rule["sheet"] == "Dashboard" and rule["cell"] == "E7":
            rule["contains"] = "SUM(tblRenovationPurchases[Amount])"
    area = next(item["values"][0] for item in contract["enum_contract"] if item["column"] == "Area")
    def case(name, inputs, assertions, actions=None):
        return {"name": name, "purpose": name, "inputs": inputs, "assertions": assertions, "actions": actions or []}
    def assertion(sheet, cell, value):
        return {"sheet": sheet, "cell": cell, "operator": "equals", "value": value}
    contract["cases"].extend([
        case("Refund reduces total spend", [{"sheet": "Purchases", "cell": "F8", "value": -100}], [assertion("Dashboard", "E7", 3800), assertion("Budget", "D8", -100)]),
        case("Empty purchase amounts produce zero spend", [{"sheet": "Purchases", "cell": f"F{row}", "value": None} for row in range(8, 15)], [assertion("Dashboard", "E7", 0), assertion("Dashboard", "K7", 24000)]),
        case("Overdue task never has negative days left", [{"sheet": "Timeline", "cell": "E8", "value": 1}], [assertion("Timeline", "G8", 0)]),
        case("Added table rows remain included beyond row 105", [], [assertion("Dashboard", "E7", 5010), assertion("Budget", "D8", 1110)], [{"operation": "append_rows", "sheet": "Purchases", "table": "tblRenovationPurchases", "count": 111, "values": {"Area": area, "Amount": 10}}]),
        case("Deleting a purchase updates totals", [], [assertion("Dashboard", "E7", 2100)], [{"operation": "delete_row", "sheet": "Purchases", "table": "tblRenovationPurchases", "index": 2}]),
    ])
    return contract


def apply_project_behaviour(workbook, spec: dict) -> None:
    """Use table references so Excel insertion/deletion updates all totals."""
    purchases = workbook.Worksheets("Purchases").ListObjects("tblRenovationPurchases")
    if spec.get("product_id") == 1:
        for sheet_name, replacements in {
            "Budget": {"Fixtures": "Equipment", "Finishes": "Soil", "Permits": "Workshops", "Design": "Planning"},
            "Purchases": {"Room": "Location"},
        }.items():
            for table in workbook.Worksheets(sheet_name).ListObjects:
                for old, new in replacements.items():
                    table.ListColumns(old).Name = new
        workbook.Worksheets("Budget").Range("B6").Value = "GARDEN BUDGET · INPUTS IN YELLOW"
        workbook.Worksheets("Budget").Range("B18").Value = "Plan shared garden costs and compare each area with recorded purchases."
        workbook.Worksheets("Dashboard").Range("B10").Value = "BUDGET BY GARDEN AREA"
    budget = workbook.Worksheets("Budget")
    for row in range(8, 15):
        budget.Range(f"D{row}").Formula = f'=SUMIF(tblRenovationPurchases[Area],B{row},tblRenovationPurchases[Amount])'
    workbook.Worksheets("Dashboard").Range("E7").Formula = "=SUM(tblRenovationPurchases[Amount])"
    workbook.Worksheets("Purchases").Range("F15").Formula = "=SUM(tblRenovationPurchases[Amount])"
    timeline = workbook.Worksheets("Timeline")
    timeline.Range("C20").Formula = "=COUNTA(tblRenovationTimeline[Task])"
    timeline.Range("F20").Formula = '=COUNTIF(tblRenovationTimeline[Status],"Open")'
    timeline.Range("G20").Formula = '=IFERROR(AVERAGE(tblRenovationTimeline[Days left]),0)'
    workbook.Worksheets("Start here").Range("B25").Value = "Add records using Insert > Table Rows inside Purchases. Edit highlighted inputs only. Negative amounts record refunds. Keep the totals outside the table."
    # Explicitly expose the input region through table expansion, not a fixed
    # row capacity. Excel copies validation and formatting on ListRows.Add.
    for column in ("Amount",):
        purchases.ListColumns(column).DataBodyRange.NumberFormat = '#,##0.00'

