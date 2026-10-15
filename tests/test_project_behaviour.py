from openpyxl import Workbook
from model_contracts import contract_for_product, validate_contract_schema
from model_quality import validate_functional_contract_schema, run_functional_contract
from preflight_visual_contract import table_layout_contract
from prepare_autonomous_drafts import launch_price


def test_project_contracts_cover_growth_refunds_empty_and_deleted_rows():
    for product_id in (1, 12, 14):
        contract = contract_for_product(product_id)
        assert not validate_contract_schema(contract)
        assert not validate_functional_contract_schema(contract, contract["sheets"])
        assert len(contract["cases"]) >= 6
        actions = [action for case in contract["cases"] for action in case.get("actions", [])]
        assert any(action.get("count", 0) > 105 for action in actions)
        assert any(action["operation"] == "delete_row" for action in actions)


def test_missing_functional_cases_cannot_pass():
    assert not run_functional_contract(None, {})["passed"]
    assert validate_functional_contract_schema({"cases": [{"inputs": []}]}, ["Data"])


def test_freeze_panes_must_match_declaration():
    book = Workbook()
    book.active.title = "Data"
    book.active.freeze_panes = "B8"
    assert not table_layout_contract(book)["passed"]
    assert table_layout_contract(book, {"freeze_panes": {"Data": "B8"}})["passed"]


def test_price_cap_does_not_rewrite_observed_evidence():
    reference, raw, price, observations = launch_price([5, 10, 15])
    assert reference == 10
    assert observations == [5, 10, 15]
    assert price == 3
