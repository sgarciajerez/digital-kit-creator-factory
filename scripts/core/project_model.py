"""Behaviour shared by the native project-budget family."""
from __future__ import annotations


def supports_quote_forecast(contract: dict) -> bool:
    """Recognize this behaviour by its data model, not by a product ID."""
    required = {
        "Budget": {"Area", "Planned", "Actual spend", "Forecast"},
        "Purchases": {"Area", "Amount"},
        "Quotes": {"Area", "Amount", "Selected", "Decision"},
    }
    by_sheet = {
        table.get("sheet"): {str(column.get("name", "")).strip() for column in table.get("headers", [])}
        for table in contract.get("tables", [])
    }
    has_area_values = any(
        item.get("column") == "Area" and item.get("values")
        for item in contract.get("enum_contract", [])
    )
    return has_area_values and all(columns <= by_sheet.get(sheet, set()) for sheet, columns in required.items())


def extend_contract(contract: dict) -> dict:
    if contract.get("product_id") == 1:
        names = {"Fixtures": "Equipment", "Finishes": "Soil", "Permits": "Workshops", "Design": "Planning", "Room": "Location"}
        for table in contract["tables"]:
            for column in table["headers"]:
                column["name"] = names.get(column["name"], column["name"])
        for policy in contract.get("input_text_policy", []):
            policy["column"] = names.get(policy["column"], policy["column"])
    if not supports_quote_forecast(contract):
        return contract
    for rule in contract["formula_rules"]:
        if rule["sheet"] == "Dashboard" and rule["cell"] == "E7":
            rule["contains"] = "SUM(tblRenovationPurchases[Amount])"
    forecast_rule = {
        "sheet": "Budget", "cell": "O8", "contains": "SUMIFS(tblRenovationQuotes[Amount]",
        "description": "Forecast includes selected preferred quotes for the matching budget area; actual spend stays tied to purchases.",
    }
    if not any(rule.get("sheet") == "Budget" and rule.get("cell") == "O8" for rule in contract["formula_rules"]):
        contract["formula_rules"].append(forecast_rule)
    relationship = "Budget.Forecast includes selected Preferred quotes for the matching Area; quotes never count as paid Actual spend."
    if relationship not in contract["relationships"]:
        contract["relationships"].append(relationship)
    area = next(item["values"][0] for item in contract["enum_contract"] if item["column"] == "Area")
    def case(name, inputs, assertions, actions=None):
        return {"name": name, "purpose": name, "inputs": inputs, "assertions": assertions, "actions": actions or []}
    def assertion(sheet, cell, value):
        return {"sheet": sheet, "cell": cell, "operator": "equals", "value": value}
    new_cases = [
        case("Selected quote flows into its budget forecast, not actual spend", [
            *[{"sheet": "Quotes", "cell": f"F{row}", "value": "No"} for row in range(8, 28)],
            {"sheet": "Budget", "cell": "C8", "value": 1},
            {"sheet": "Quotes", "cell": "B8", "value": "New supplier"},
            {"sheet": "Quotes", "cell": "C8", "value": area},
            {"sheet": "Quotes", "cell": "E8", "value": 12345},
            {"sheet": "Quotes", "cell": "F8", "value": "Yes"},
            {"sheet": "Quotes", "cell": "I8", "value": "Preferred"},
        ], [assertion("Budget", "O8", 12345), {"sheet": "Budget", "cell": "D8", "operator": "unchanged"}, {"sheet": "Dashboard", "cell": "E7", "operator": "unchanged"}]),
        case("Unselected or backup quote is excluded from the forecast", [
            *[{"sheet": "Quotes", "cell": f"F{row}", "value": "No"} for row in range(8, 28)],
            {"sheet": "Budget", "cell": "C8", "value": 1},
            {"sheet": "Quotes", "cell": "C8", "value": area},
            {"sheet": "Quotes", "cell": "E8", "value": 12345},
            {"sheet": "Quotes", "cell": "I8", "value": "Backup"},
        ], [assertion("Budget", "O8", 1), {"sheet": "Dashboard", "cell": "E7", "operator": "unchanged"}]),
        case("Refund reduces total spend", [{"sheet": "Purchases", "cell": "F8", "value": -100}], [assertion("Dashboard", "E7", 3800), assertion("Budget", "D8", -100)]),
        case("Empty purchase amounts produce zero spend", [{"sheet": "Purchases", "cell": f"F{row}", "value": None} for row in range(8, 15)], [assertion("Dashboard", "E7", 0), assertion("Dashboard", "K7", 24000)]),
        case("Overdue task never has negative days left", [{"sheet": "Timeline", "cell": "E8", "value": 1}], [assertion("Timeline", "G8", 0)]),
        case("Added table rows remain included beyond row 105", [], [assertion("Dashboard", "E7", 5010), assertion("Budget", "D8", 1110)], [{"operation": "append_rows", "sheet": "Purchases", "table": "tblRenovationPurchases", "count": 111, "values": {"Area": area, "Amount": 10}}]),
        case("Deleting a purchase updates totals", [], [assertion("Dashboard", "E7", 2100)], [{"operation": "delete_row", "sheet": "Purchases", "table": "tblRenovationPurchases", "index": 2}]),
    ]
    existing_case_names = {item.get("name") for item in contract["cases"]}
    contract["cases"].extend(item for item in new_cases if item["name"] not in existing_case_names)
    return contract


def apply_project_behaviour(workbook, spec: dict) -> None:
    """Use table references so Excel insertion/deletion updates all totals."""
    if not supports_quote_forecast(spec.get("model_contract", {})):
        return
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
    for row in range(8, 15):
        budget.Range(f"O{row}").Formula = (
            f'=MAX(C{row},D{row}+SUMIFS(tblRenovationQuotes[Amount],'
            f'tblRenovationQuotes[Area],B{row},tblRenovationQuotes[Selected],"Yes",'
            f'tblRenovationQuotes[Decision],"Preferred"))'
        )
    workbook.Worksheets("Start here").Range("B25").Value = "Add records using Insert > Table Rows inside Purchases. Edit highlighted inputs only. Negative amounts record refunds. Keep the totals outside the table."
    # Explicitly expose the input region through table expansion, not a fixed
    # row capacity. Excel copies validation and formatting on ListRows.Add.
    for sheet, address in (
        ("Budget", "C15:O15"), ("Budget", "O8:O14"),
        ("Purchases", "F15"), ("Quotes", "E8:E27"),
        ("Dashboard", "B7:B7"), ("Dashboard", "E7:E7"),
        ("Dashboard", "K7:K7"), ("Dashboard", "C19:E19"),
    ):
        workbook.Worksheets(sheet).Range(address).ShrinkToFit = True

