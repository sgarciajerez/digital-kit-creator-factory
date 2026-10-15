from copy import deepcopy
from openpyxl import Workbook
from model_contracts import contract_for_product, validate_contract_schema
from model_quality import validate_functional_contract_schema, run_functional_contract
from preflight_visual_contract import table_layout_contract
from prepare_autonomous_drafts import launch_price
from project_model import extend_contract


def test_project_contracts_cover_growth_refunds_empty_and_deleted_rows():
    for product_id in (1, 12, 14):
        contract = contract_for_product(product_id)
        assert not validate_contract_schema(contract)
        assert not validate_functional_contract_schema(contract, contract["sheets"])
        assert len(contract["cases"]) >= 6
        actions = [action for case in contract["cases"] for action in case.get("actions", [])]
        assert any(action.get("count", 0) > 105 for action in actions)
        assert any(action["operation"] == "delete_row" for action in actions)
        assert any("SUMIFS(tblRenovationQuotes[Amount]" in rule["contains"] for rule in contract["formula_rules"])
        assert any("Selected quote flows" in case["name"] for case in contract["cases"])


def test_quote_forecast_contract_is_applied_by_schema_for_future_product_ids():
    future_contract = deepcopy(contract_for_product(14))
    future_contract["product_id"] = 999
    future_contract["formula_rules"] = [
        rule for rule in future_contract["formula_rules"]
        if not (rule.get("sheet") == "Budget" and rule.get("cell") == "O8")
    ]
    future_contract["relationships"] = [
        relationship for relationship in future_contract["relationships"]
        if "Budget.Forecast includes selected Preferred quotes" not in relationship
    ]
    future_contract["cases"] = [
        case for case in future_contract["cases"]
        if not case["name"].startswith(("Selected quote flows", "Unselected or backup quote"))
    ]

    extend_contract(future_contract)
    extend_contract(future_contract)

    assert sum(rule.get("sheet") == "Budget" and rule.get("cell") == "O8" for rule in future_contract["formula_rules"]) == 1
    assert sum("Budget.Forecast includes selected Preferred quotes" in item for item in future_contract["relationships"]) == 1
    assert sum(case["name"].startswith("Selected quote flows") for case in future_contract["cases"]) == 1
    assert any("SUMIFS(tblRenovationQuotes[Amount]" in rule["contains"] for rule in future_contract["formula_rules"])


def test_quote_forecast_contract_does_not_apply_to_unrelated_schema():
    unrelated = {"product_id": 999, "tables": [], "enum_contract": [], "formula_rules": [], "relationships": [], "cases": []}
    assert extend_contract(unrelated) == unrelated
    assert unrelated["formula_rules"] == []


def test_missing_functional_cases_cannot_pass():
    assert not run_functional_contract(None, {})["passed"]
    assert validate_functional_contract_schema({"cases": [{"inputs": []}]}, ["Data"])


def test_freeze_panes_must_match_declaration():
    book = Workbook()
    book.active.title = "Data"
    book.active.freeze_panes = "B8"
    assert not table_layout_contract(book)["passed"]
    assert table_layout_contract(book, {"freeze_panes": {"Data": "B8"}})["passed"]


def test_price_uses_uncapped_comparable_market_median():
    reference, raw, price, observations = launch_price([5, 10, 15])
    assert reference == 10
    assert raw == 10
    assert price == 10
    assert observations == [5, 10, 15]


def test_price_converts_comparables_before_taking_the_median():
    reference, raw, price, observations = launch_price([2.74, 8.37, 12.97], observed_to_shop_currency=0.89325)
    assert reference == 7.48
    assert raw == 7.48
    assert price == 7.48
    assert observations == [2.74, 8.37, 12.97]
