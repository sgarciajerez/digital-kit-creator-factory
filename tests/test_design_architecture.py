import json

from design_architecture import apply_design_architecture, load_design_system, validate_design_spec


def test_design_architecture_generates_ordered_template_contract() -> None:
    spec = {
        "product_id": 99,
        "tabs": ["Start here", "Dashboard", "Items", "Settings"],
    }

    apply_design_architecture(spec)

    assert list(spec["screen_contracts"]) == spec["tabs"]
    assert spec["screen_contracts"]["Dashboard"]["template"] == "dashboard"
    assert spec["screen_contracts"]["Items"]["template"] == "table"
    assert spec["navigation_contract"]["button_width_mode"] == "shared_base_buttons"
    assert spec["navigation_contract"]["component"] == "rounded_in_sheet_buttons_plus_native_excel_sheet_tabs"
    assert validate_design_spec(spec) == []


def test_design_architecture_rejects_freeform_component_and_override() -> None:
    system = load_design_system()
    spec = {
        "design_system_contract": "config/FACTORY_DESIGN_SYSTEM.json",
        "layout_mode": "template_inheritance",
        "tabs": ["Dashboard"],
        "screen_contracts": {
            "Dashboard": {
                "template": "dashboard",
                "components": ["header", "identity", "freeform"],
                "overrides": {"custom_css": True},
                "content": {"geometry_source": "browser_only"},
            }
        },
        "navigation_contract": {
            "tabs": ["Dashboard"],
            "component": "arbitrary_buttons",
            "button_width_mode": "fixed",
        },
    }

    errors = validate_design_spec(spec, system)

    assert any("componentes no permitidos" in error for error in errors)
    assert any("overrides" in error for error in errors)
    assert any("navegación" in error for error in errors)
