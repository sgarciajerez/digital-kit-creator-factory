from __future__ import annotations

"""Declarative model contracts for the Excel factory.

This module is deliberately independent from workbook inspection when a
contract is authored or rendered.  A contract describes what a product is
supposed to do; the validator in :mod:`model_quality` compares that contract
with the generated workbook afterwards.
"""

import json
from copy import deepcopy
from pathlib import Path
from typing import Any

from openpyxl.utils.cell import range_boundaries


ROOT = Path(__file__).resolve().parents[2]


def _columns(*items: tuple[str, str, str]) -> list[dict[str, str]]:
    return [
        {"name": name, "role": role, "kind": kind}
        for name, role, kind in items
    ]


def _master_input_text_policy() -> list[dict[str, str]]:
    """Classify every editable text field in the master-derived model."""
    enum = "enum"
    free = "free_text"
    return [
        {"sheet": "Settings", "column": "Setting", "mode": free},
        {"sheet": "Budget", "column": "Notes", "mode": free},
        {"sheet": "Purchases", "column": "Area", "mode": enum},
        {"sheet": "Purchases", "column": "Item", "mode": free},
        {"sheet": "Purchases", "column": "Supplier", "mode": free},
        {"sheet": "Purchases", "column": "Paid?", "mode": enum},
        {"sheet": "Purchases", "column": "Status", "mode": enum},
        {"sheet": "Purchases", "column": "Phase", "mode": free},
        {"sheet": "Purchases", "column": "Quote ref", "mode": free},
        {"sheet": "Purchases", "column": "Room", "mode": free},
        {"sheet": "Purchases", "column": "Payment method", "mode": free},
        {"sheet": "Purchases", "column": "Receipt", "mode": free},
        {"sheet": "Purchases", "column": "Notes", "mode": free},
        {"sheet": "Quotes", "column": "Supplier", "mode": free},
        {"sheet": "Quotes", "column": "Area", "mode": enum},
        {"sheet": "Quotes", "column": "Selected", "mode": enum},
        {"sheet": "Quotes", "column": "Notes", "mode": free},
        {"sheet": "Quotes", "column": "Lead time", "mode": free},
        {"sheet": "Quotes", "column": "Decision", "mode": enum},
        {"sheet": "Timeline", "column": "Task", "mode": free},
        {"sheet": "Timeline", "column": "Area", "mode": enum},
        {"sheet": "Timeline", "column": "Status", "mode": enum},
        {"sheet": "Timeline", "column": "Owner", "mode": free},
        {"sheet": "Timeline", "column": "Phase", "mode": free},
        {"sheet": "Timeline", "column": "Notes", "mode": free},
    ]


def _table(
    sheet: str,
    ref: str,
    headers: list[dict[str, str]],
    *,
    total_ref: str | None = None,
    total_label: str | None = None,
    name: str | None = None,
) -> dict[str, Any]:
    item: dict[str, Any] = {
        "sheet": sheet,
        "ref": ref,
        "headers": headers,
        "required": True,
    }
    if name:
        item["name"] = name
    if total_ref:
        item["total_ref"] = total_ref
    if total_label:
        item["total_label"] = total_label
    return item


def _formula(sheet: str, cell: str, contains: str, description: str) -> dict[str, str]:
    return {"sheet": sheet, "cell": cell, "contains": contains, "description": description}


def _case(
    name: str,
    inputs: list[dict[str, Any]],
    assertions: list[dict[str, Any]],
    purpose: str,
) -> dict[str, Any]:
    return {"name": name, "purpose": purpose, "inputs": inputs, "assertions": assertions}


def family_base_contract() -> dict[str, Any]:
    """Expected behavior of the canonical family-budget master workbook."""
    money = "amount"
    contract = {
        "contract_id": "family-budget-master-v1",
        "title": "Base de presupuesto familiar",
        "product_id": None,
        "purpose": (
            "Registrar movimientos una sola vez, comparar presupuesto y gasto real "
            "por mes y categoría, y convertir el resultado en flujo de caja, metas y deudas."
        ),
        "sheets": [
            "Inicio", "Resumen", "Presupuesto", "Gastos reales", "Movimientos",
            "Flujo de caja", "Metas y fondo", "Deudas", "Listas",
        ],
        "tables": [
            _table("Resumen", "B11:F18", _columns(
                ("Categoría", "input", "text"), ("Plan", "derived", money),
                ("Real", "derived", money), ("Diferencia", "derived", money),
                ("% gastado", "derived", "percent"),
            )),
            _table("Presupuesto", "B7:O14", _columns(
                ("Categoría", "input", "text"), *( (month, "input", money) for month in
                ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")),
                ("Total anual", "derived", money),
            ), total_ref="B15:O15", total_label="TOTAL DEL HOGAR"),
            _table("Gastos reales", "B7:O14", _columns(
                ("Categoría", "derived", "text"), *( (month, "derived", money) for month in
                ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")),
                ("Total anual", "derived", money),
            ), total_ref="B15:O15", total_label="TOTAL REAL"),
            _table("Movimientos", "B7:I27", _columns(
                ("Fecha", "input", "date"), ("Mes", "input", "text"), ("Tipo", "input", "text"),
                ("Categoría", "input", "text"), ("Concepto", "input", "text"),
                ("Importe", "input", money), ("Pagado por", "input", "text"), ("Estado", "input", "text"),
            )),
            _table("Flujo de caja", "B7:J19", _columns(
                ("Mes", "derived", "text"), ("Ingresos plan", "derived", money),
                ("Ingresos real", "derived", money), ("Gastos plan", "derived", money),
                ("Gastos real", "derived", money), ("Ahorro plan", "derived", money),
                ("Ahorro real", "derived", money), ("Disponible", "derived", money),
                ("Comentarios", "input", "text"),
            )),
            _table("Metas y fondo", "B9:J12", _columns(
                ("Meta", "input", "text"), ("Gasto mensual", "input", money),
                ("Meses", "input", "count"), ("Meta necesaria", "derived", money),
                ("Ahorrado", "input", money), ("Pendiente", "derived", money),
                ("Progreso", "derived", "percent"), ("Fecha objetivo", "input", "date"),
                ("Estado", "derived", "text"),
            )),
            _table("Deudas", "B7:H10", _columns(
                ("Deuda", "input", "text"), ("Saldo actual", "input", money),
                ("Interés anual", "input", "percent"), ("Pago mensual", "input", money),
                ("Meses restantes", "input", "count"), ("Interés estimado", "derived", money),
                ("Estado", "derived", "text"),
            )),
            _table("Listas", "B7:C13", _columns(("Parámetro", "input", "text"), ("Valor", "input", "general"))),
        ],
        "formula_rules": [
            _formula("Resumen", "E12", "C12-D12", "La diferencia es presupuesto menos gasto real."),
            _formula("Resumen", "F12", "D12/C12", "El porcentaje compara gasto real con presupuesto."),
            _formula("Gastos reales", "C8", "SUMIFS", "El gasto real nace de Movimientos y no se duplica manualmente."),
            _formula("Flujo de caja", "I8", "D8-F8-H8", "Disponible = ingreso real - gasto real - ahorro real."),
            _formula("Metas y fondo", "G10", "MAX(E10-F10,0)", "Pendiente nunca puede ser negativa."),
            _formula("Metas y fondo", "H10", "F10/E10", "Progreso = ahorrado / meta necesaria."),
        ],
        "relationships": [
            "Movimientos es la única entrada operativa para gastos reales y flujo de caja.",
            "Presupuesto y Gastos reales comparten categorías y meses.",
            "Las filas TOTAL quedan fuera del rango de datos de cada tabla.",
            "Las columnas de conteo nunca se expresan como porcentaje.",
        ],
        "cases": [
            _case("income flows into monthly cash flow", [{"sheet": "Listas", "cell": "C9", "value": 3500}], [
                {"sheet": "Flujo de caja", "cell": "C8", "operator": "equals", "value": 3500},
            ], "Comprueba que el ingreso neto declarado alimenta la lectura mensual."),
        ],
        "failure_criteria": [
            "Una tabla tiene una cabecera vacía, genérica, duplicada o con un significado distinto.",
            "Una fila TOTAL queda dentro del rango filtrable de datos.",
            "Gastos reales deja de depender de Movimientos o Disponible deja de seguir su ecuación.",
            "Una columna de conteo usa formato porcentaje o una columna monetaria usa porcentaje.",
        ],
    }
    for table in contract["tables"]:
        if table.get("sheet") == "Presupuesto":
            table["headers"][0]["role"] = "derived"
    contract["formula_rules"].append(
        _formula("Presupuesto", "B8", "Listas!I7", "La categoría presupuestaria se obtiene de la lista editable de Listas.")
    )
    return contract


def wedding_base_contract() -> dict[str, Any]:
    """Expected behavior of the legacy wedding base kept as a product base."""
    money = "amount"
    return {
        "contract_id": "wedding-planner-base-v1",
        "title": "Base de planificador de boda",
        "product_id": None,
        "purpose": "Comparar presupuesto y pagos, seguir proveedores, invitados y tareas sin duplicar entradas.",
        "sheets": ["Inicio", "Resumen", "Presupuesto", "Pagos reales", "Proveedores", "Flujo de pagos", "Checklist", "Invitados", "Listas"],
        "tables": [
            _table("Resumen", "B11:F18", _columns(("Categoría", "derived", "text"), ("Plan", "derived", money), ("Pagado", "derived", money), ("Pendiente", "derived", money), ("% pagado", "derived", "percent"))),
            _table("Presupuesto", "B7:O14", _columns(("Categoría", "input", "text"), *( (month, "input", money) for month in
                ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")), ("Total", "derived", money)), total_ref="B15:O15", total_label="TOTAL DE LA BODA"),
            _table("Pagos reales", "B7:O14", _columns(("Categoría", "derived", "text"), *( (month, "derived", money) for month in
                ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")), ("Total", "derived", money)), total_ref="B15:O15", total_label="TOTAL PAGADO"),
            _table("Proveedores", "B7:I27", _columns(("Fecha", "input", "date"), ("Mes", "input", "text"), ("Tipo", "input", "text"), ("Categoría", "input", "text"), ("Concepto", "input", "text"), ("Importe", "input", money), ("Pagado por", "input", "text"), ("Estado", "input", "text"))),
            _table("Flujo de pagos", "B7:J19", _columns(("Mes", "derived", "text"), ("Presupuesto plan", "derived", money), ("Ingresos", "derived", money), ("Pagos previstos", "derived", money), ("Pagos reales", "derived", money), ("Ajustes", "derived", money), ("Saldo previsto", "derived", money), ("Comentarios", "input", "text"), ("Estado", "derived", "text"))),
            _table("Checklist", "B9:J12", _columns(("Tarea", "input", "text"), ("Área", "input", "text"), ("Responsable", "input", "text"), ("Fecha límite", "input", "date"), ("Estado", "input", "text"), ("Prioridad", "input", "text"), ("Días", "derived", "count"), ("Notas", "input", "text"), ("Progreso", "derived", "percent"))),
            _table("Invitados", "B7:H10", _columns(("Grupo", "input", "text"), ("Invitados previstos", "input", "count"), ("Confirmados", "input", "count"), ("Menú", "input", "text"), ("Mesa", "input", "count"), ("Estado", "input", "text"), ("Notas", "input", "text")), total_ref="B11:H11", total_label="TOTAL"),
            _table("Listas", "B7:C13", _columns(("Parámetro", "input", "text"), ("Valor", "input", "general"))),
        ],
        "formula_rules": [
            _formula("Resumen", "E12", "C12-D12", "Pendiente = plan menos pagado."),
            _formula("Pagos reales", "C8", "SUMIFS", "Los pagos reales proceden del diario de proveedores."),
            _formula("Flujo de pagos", "H8", "C8-F8+D8-G8", "Saldo previsto combina presupuesto, ingresos, pagos y ajustes."),
            _formula("Checklist", "J10", "IF(F10=\"Confirmado\",1,0)", "Progreso refleja el estado de la tarea."),
            _formula("Invitados", "C11", "SUM(C8:C10)", "El total de invitados previstos suma grupos, no porcentajes."),
        ],
        "relationships": [
            "Proveedores es el diario operativo; Pagos reales y Flujo de pagos se alimentan de él.",
            "Invitados previstos, Confirmados y Mesa son conteos enteros.",
            "Los porcentajes solo describen progreso o pagado, nunca número de invitados.",
            "Las filas TOTAL están fuera de las tablas de datos.",
        ],
        "cases": [
            _case("guest total follows group counts", [{"sheet": "Invitados", "cell": "C8", "value": 40}], [
                {"sheet": "Invitados", "cell": "C11", "operator": "equals", "value": 130},
            ], "Comprueba que el total de invitados suma personas y no interpreta la columna como porcentaje."),
        ],
        "failure_criteria": [
            "Invitados previstos, Confirmados o Mesa usan formato porcentaje.",
            "Los pagos reales se introducen duplicados en vez de salir de Proveedores.",
            "Pendiente o saldo previsto deja de ser una relación calculada.",
            "Una fila TOTAL queda dentro del rango filtrable de la tabla.",
        ],
    }


def _product_contracts() -> dict[int, dict[str, Any]]:
    money = "amount"
    contracts = {
        8: {
            "contract_id": "wedding-planner-simple-v1",
            "title": "Planificador de boda",
            "product_id": 8,
            "purpose": "Planificar presupuesto, invitados, proveedores, pagos y tareas desde una única boda conectada.",
            "sheets": ["Start here", "Dashboard", "Budget", "Guests", "Vendors", "Payments", "Checklist", "Settings"],
            "tables": [
                _table("Settings", "B7:D13", _columns(("Setting", "input", "text"), ("Value", "input", "general"), ("Used by", "derived", "text"))),
                _table("Budget", "B7:F13", _columns(("Category", "input", "text"), ("Planned", "input", money), ("Actual", "input", money), ("Remaining", "derived", money), ("Status", "input", "text")), total_ref="B14:F14", total_label="TOTAL"),
                _table("Guests", "B7:F14", _columns(("Name", "input", "text"), ("Party", "input", "count"), ("RSVP", "input", "text"), ("Meal", "input", "text"), ("Notes", "input", "text"))),
                _table("Vendors", "B7:G13", _columns(("Vendor", "input", "text"), ("Category", "input", "text"), ("Estimate", "input", money), ("Deposit", "input", money), ("Balance", "derived", money), ("Status", "input", "text"))),
                _table("Payments", "B7:E8", _columns(("Total paid", "derived", money), ("Due next 60 days", "derived", money), ("Open items", "derived", "count"), ("Currency", "derived", "text"))),
                _table("Payments", "B11:G16", _columns(("Due date", "input", "date"), ("Vendor", "input", "text"), ("Category", "input", "text"), ("Amount", "input", money), ("Paid", "input", "text"), ("Notes", "input", "text"))),
                _table("Checklist", "B7:E14", _columns(("Task", "input", "text"), ("Owner", "input", "text"), ("Due", "input", "date"), ("Status", "input", "text"))),
                _table("Dashboard", "B7:E8", _columns(("Total budget", "derived", money), ("Paid so far", "derived", money), ("Guests confirmed", "derived", "count"), ("Days to go", "derived", "count"))),
                _table("Dashboard", "B11:F17", _columns(("Category", "derived", "text"), ("Planned", "derived", money), ("Actual", "derived", money), ("Remaining", "derived", money), ("Status", "derived", "text")), total_ref="B18:F18", total_label="TOTAL"),
            ],
            "formula_rules": [
                _formula("Budget", "E8", "C8-D8", "Saldo por categoría = planificado menos real."),
                _formula("Budget", "C14", "SUM(C8:C13)", "El total está fuera de la tabla y suma solo las filas de datos."),
                _formula("Vendors", "F8", "D8-E8", "Saldo del proveedor = estimación menos depósito."),
                _formula("Payments", "B8", "SUMIF(F12:F107,\"Yes\",E12:E107)", "Total pagado depende del registro de pagos."),
                _formula("Dashboard", "B8", "Settings!C9", "El presupuesto del panel sigue al parámetro de Settings."),
                _formula("Dashboard", "D8", "SUMIF(Guests!D8:D107", "Los invitados confirmados suman Party solo para RSVP confirmado."),
            ],
            "relationships": [
                "Party es un número de personas, no un porcentaje.",
                "Guests!Party solo se suma cuando RSVP = Confirmed.",
                "Payments resume el log de pagos; no se introducen totales manualmente.",
                "Las filas TOTAL están fuera de los ListObjects.",
            ],
            "cases": [
                _case("wedding budget and guests", [{"sheet": "Settings", "cell": "C9", "value": 25000}], [
                    {"sheet": "Dashboard", "cell": "B8", "operator": "equals", "value": 25000},
                    {"sheet": "Dashboard", "cell": "D8", "operator": "equals", "value": 8},
                    {"sheet": "Budget", "cell": "E8", "operator": "equals", "value": 1600},
                ], "Comprueba que un parámetro, una suma de personas y un saldo de tabla están conectados."),
            ],
            "failure_criteria": [
                "Guests.Party se formatea como porcentaje o se cuenta como una fila en vez de sumar personas.",
                "Una columna Amount, Planned, Actual, Balance o Estimate se expresa como porcentaje.",
                "El presupuesto del Dashboard no sigue a Settings o los pagos no siguen el log de Payments.",
                "Una fila TOTAL forma parte del rango editable o filtrable.",
            ],
        },
        9: {
            "contract_id": "debt-payoff-planner-v1",
            "title": "Planificador de deudas",
            "product_id": 9,
            "purpose": "Registrar deudas y pagos, separar interés de principal y comparar el ritmo de amortización.",
            "sheets": ["Start here", "Dashboard", "Debts", "Payments", "Simulator", "Progress", "Settings"],
            "tables": [
                _table("Debts", "A5:F8", _columns(("Debt", "input", "text"), ("Type", "input", "text"), ("Balance", "input", money), ("APR", "input", "percent"), ("Minimum", "input", money), ("Priority", "input", "count"))),
                _table("Payments", "A5:E8", _columns(("Date", "input", "date"), ("Debt", "input", "text"), ("Amount", "input", money), ("Interest", "derived", money), ("Principal", "derived", money))),
                _table("Simulator", "A5:C9", _columns(("Input", "input", "text"), ("Value", "input", "general"), ("Result", "derived", "general"))),
                _table("Progress", "A5:D13", _columns(("Month", "derived", "text"), ("Balance", "derived", money), ("Paid down", "derived", money), ("Progress", "derived", "percent"))),
                _table("Settings", "A5:B10", _columns(("Setting", "input", "text"), ("Value", "input", "general"))),
            ],
            "formula_rules": [
                _formula("Payments", "D6", "INDEX(Debts!$C$6:$C$8", "Interest uses the matching debt balance and APR divided by 12."),
                _formula("Payments", "E6", "MAX(C6-D6,0)", "Principal is payment minus interest and never negative."),
                _formula("Dashboard", "A6", "SUM(Debts!C6:C8)", "Total balance is the sum of debt balances."),
                _formula("Dashboard", "G6", "Settings!B9", "Extra payment comes from the declared settings input."),
                _formula("Simulator", "C8", "ROUNDUP(SUM(Debts!C6:C8)/(SUM(Debts!E6:E8)+B6)", "Projected months uses minimum payments plus the extra payment."),
            ],
            "relationships": [
                "Interest is calculated monthly from the matching balance and APR.",
                "Principal cannot exceed the payment and cannot be negative.",
                "The simulator's extra payment is a numeric amount, never a percentage.",
                "Progress is a ratio and is the only percentage column in its table.",
            ],
            "cases": [
                _case("debt payment allocation", [{"sheet": "Settings", "cell": "B9", "value": 250}], [
                    {"sheet": "Dashboard", "cell": "A6", "operator": "equals", "value": 18640},
                    {"sheet": "Dashboard", "cell": "G6", "operator": "equals", "value": 250},
                    {"sheet": "Payments", "cell": "D6", "operator": "equals", "value": 66.4},
                    {"sheet": "Payments", "cell": "E6", "operator": "equals", "value": 113.6},
                    {"sheet": "Simulator", "cell": "C8", "operator": "equals", "value": 19},
                ], "Comprueba que el pago se divide correctamente entre interés y principal y que el simulador usa el importe extra."),
            ],
            "failure_criteria": [
                "APR, Progress o una tasa se formatea como importe o como conteo.",
                "Priority, cantidad de meses o personas se formatea como porcentaje.",
                "El interés no depende del saldo y APR de la deuda seleccionada.",
                "El caso de pago produce un resultado distinto de 66,40 de interés, 113,60 de principal y 19 meses.",
            ],
        },
        10: {
            "contract_id": "small-business-cashflow-v1",
            "title": "Flujo de caja para pequeño negocio",
            "product_id": 10,
            "purpose": "Registrar entradas y salidas, anticipar facturas y conservar una previsión de caja encadenada.",
            "sheets": ["Start here", "Dashboard", "Transactions", "Invoices", "Recurring", "Forecast", "Year end", "Settings"],
            "tables": [
                _table("Transactions", "A5:F10", _columns(("Date", "input", "date"), ("Description", "input", "text"), ("Category", "input", "text"), ("Money in", "input", money), ("Money out", "input", money), ("Status", "input", "text"))),
                _table("Invoices", "A5:F9", _columns(("Invoice", "input", "text"), ("Client", "input", "text"), ("Due", "input", "date"), ("Amount", "input", money), ("Status", "input", "text"), ("Notes", "input", "text"))),
                _table("Recurring", "A5:F9", _columns(("Item", "input", "text"), ("Category", "input", "text"), ("Frequency", "input", "text"), ("Next due", "input", "date"), ("Amount", "input", money), ("Type", "input", "text"))),
                _table("Forecast", "A4:E10", _columns(("Month", "derived", "text"), ("Opening", "derived", money), ("Income", "derived", money), ("Expenses", "derived", money), ("Closing", "derived", money))),
                _table("Year end", "A5:D9", _columns(("Category", "input", "text"), ("Income", "input", money), ("Expenses", "input", money), ("Net", "derived", money)), total_ref="A10:D10", total_label="TOTAL"),
                _table("Settings", "A5:B11", _columns(("Setting", "input", "text"), ("Value", "input", "general"))),
            ],
            "formula_rules": [
                _formula("Dashboard", "A6", "Settings!B8+SUM(Transactions!D6:D105)-SUM(Transactions!E6:E105)", "Cash on hand starts at the declared opening balance."),
                _formula("Forecast", "B6", "E5", "Each forecast month opens with the previous closing balance."),
                _formula("Forecast", "E5", "B5+C5-D5", "Closing = opening + income - expenses."),
                _formula("Year end", "D6", "B6-C6", "Net = income minus expenses."),
            ],
            "relationships": [
                "Money in and Money out are separate amount columns; they are not percentages.",
                "Cash on hand adds inflows and subtracts outflows from the opening balance.",
                "Forecast is a chain: next opening equals prior closing.",
                "The total row in Year end is outside the editable table.",
            ],
            "cases": [
                _case("cash flow opening balance", [{"sheet": "Settings", "cell": "B8", "value": 6000}], [
                    {"sheet": "Dashboard", "cell": "A6", "operator": "equals", "value": 7238},
                    {"sheet": "Dashboard", "cell": "C6", "operator": "equals", "value": 2100},
                    {"sheet": "Dashboard", "cell": "E6", "operator": "equals", "value": 862},
                    {"sheet": "Forecast", "cell": "B6", "operator": "formula_contains", "value": "E5"},
                ], "Comprueba que el saldo inicial se propaga al panel y que la previsión conserva la continuidad mensual."),
            ],
            "failure_criteria": [
                "Opening balance no cambia Cash on hand.",
                "La apertura de un mes no referencia el cierre del mes anterior.",
                "Money in, Money out, Amount o Closing usan formato porcentaje.",
                "La fila TOTAL de Year end se incluye en el ListObject.",
            ],
        },
        11: {
            "contract_id": "meal-planner-v1",
            "title": "Planificador de comidas",
            "product_id": 11,
            "purpose": "Planificar comidas, reutilizar recetas, consolidar la compra y comparar gasto previsto con presupuesto semanal.",
            "sheets": ["Start here", "Dashboard", "This week", "Monthly view", "Grocery list", "Recipes", "Pantry", "Budget", "Settings"],
            "tables": [
                _table("This week", "A5:D26", _columns(("Day", "input", "text"), ("Meal", "input", "text"), ("Recipe", "input", "text"), ("Est. cost", "derived", money))),
                _table("Recipes", "A5:F10", _columns(("Recipe", "input", "text"), ("Meal type", "input", "text"), ("Prep", "input", "text"), ("Cost/serving", "input", money), ("Rating", "input", "count"), ("Tags", "input", "text"))),
                _table("Grocery list", "A5:F10", _columns(("Done", "input", "text"), ("Item", "input", "text"), ("Aisle", "input", "text"), ("Qty", "input", "count"), ("Est. cost", "input", money), ("For meals", "input", "text"))),
                _table("Pantry", "A5:F9", _columns(("Item", "input", "text"), ("Location", "input", "text"), ("Qty", "input", "count"), ("Expiry", "input", "date"), ("Status", "input", "text"), ("Suggested use", "input", "text"))),
                _table("Budget", "A4:E8", _columns(("Week", "input", "text"), ("Planned", "input", money), ("Actual", "input", money), ("Difference", "derived", money), ("Note", "derived", "text"))),
                _table("Monthly view", "A5:D9", _columns(("Week", "input", "text"), ("Focus", "input", "text"), ("Planned spend", "input", money), ("Notes", "input", "text"))),
                _table("Settings", "A5:B10", _columns(("Setting", "input", "text"), ("Value", "input", "general"))),
            ],
            "formula_rules": [
                _formula("This week", "D6", "INDEX(Recipes!D$6:D$10", "El coste estimado se consulta desde la receta elegida."),
                _formula("Budget", "D5", "B5-C5", "Diferencia = presupuesto planificado menos gasto real."),
                _formula("Dashboard", "A6", "Settings!B7", "El presupuesto semanal viene de Settings."),
                _formula("Dashboard", "C6", "SUM('Grocery list'!E6:E105)", "El gasto previsto suma la lista de compra."),
                _formula("Dashboard", "E6", "COUNTA('This week'!C6:C105)", "Las comidas planificadas cuentan recetas, no importes."),
            ],
            "relationships": [
                "Qty y Rating son conteos o cantidades y no porcentajes.",
                "Est. cost es un importe y debe conservar formato numérico/monetario.",
                "Budget compara Planned con Actual por semana.",
                "La lista de compra alimenta el gasto previsto del Dashboard.",
            ],
            "cases": [
                _case("grocery spend and budget", [{"sheet": "Grocery list", "cell": "E6", "value": 10}], [
                    {"sheet": "Dashboard", "cell": "A6", "operator": "equals", "value": 120},
                    {"sheet": "Dashboard", "cell": "C6", "operator": "equals", "value": 25.1},
                    {"sheet": "Dashboard", "cell": "G6", "operator": "equals", "value": 5},
                    {"sheet": "Budget", "cell": "D5", "operator": "equals", "value": 22},
                ], "Comprueba que una modificación en la lista de compra modifica el resumen y que la diferencia presupuestaria conserva su signo."),
            ],
            "failure_criteria": [
                "Una cantidad, número de comidas o valoración usa formato porcentaje.",
                "El gasto previsto no suma Grocery list o el presupuesto no resta Actual a Planned.",
                "El coste de una comida deja de buscarse en Recipes.",
                "Una columna de importe se interpreta como texto sin formato numérico.",
            ],
        },
        12: {
            "contract_id": "home-renovation-planner-v1",
            "title": "Planificador de presupuesto y obra de reforma",
            "product_id": 12,
            "purpose": (
                "Conectar presupuesto por partidas, compras, comparativa de presupuestos "
                "y calendario de tareas para controlar una reforma sin duplicar entradas."
            ),
            "sheets": ["Start here", "Dashboard", "Budget", "Purchases", "Quotes", "Timeline", "Settings"],
            "tables": [
                _table("Settings", "B7:C13", _columns(
                    ("Setting", "input", "text"), ("Value", "input", "general"),
                )),
                _table("Budget", "B7:O14", _columns(
                ("Area", "derived", "text"), ("Planned", "input", money),
                    ("Actual spend", "derived", money), ("Remaining", "derived", money),
                    ("Status", "derived", "text"), ("Notes", "input", "text"),
                    ("Materials", "derived", money), ("Labour", "derived", money),
                    ("Fixtures", "derived", money), ("Finishes", "derived", money),
                    ("Permits", "derived", money), ("Design", "derived", money),
                    ("Contingency", "derived", money), ("Forecast", "derived", money),
                ), total_ref="B15:O15", total_label="TOTAL PROJECT"),
                _table("Purchases", "B7:O14", _columns(
                    ("Purchase date", "input", "date"), ("Area", "input", "text"),
                    ("Item", "input", "text"), ("Supplier", "input", "text"),
                    ("Amount", "input", money), ("Paid?", "input", "text"),
                    ("Status", "input", "text"), ("Phase", "input", "text"),
                    ("Quote ref", "input", "text"), ("Room", "input", "text"),
                    ("Payment method", "input", "text"), ("Due date", "input", "date"),
                    ("Receipt", "input", "text"), ("Notes", "input", "text"),
                ), total_ref="B15:O15", total_label="TOTAL PURCHASES"),
                _table("Quotes", "B7:I27", _columns(
                    ("Supplier", "input", "text"), ("Area", "input", "text"),
                    ("Quote date", "input", "date"), ("Amount", "input", money),
                    ("Selected", "input", "text"), ("Notes", "input", "text"),
                    ("Lead time", "input", "text"), ("Decision", "input", "text"),
                )),
                _table("Timeline", "B7:J19", _columns(
                    ("Task", "input", "text"), ("Area", "input", "text"),
                    ("Start date", "input", "date"), ("Due date", "input", "date"),
                    ("Status", "input", "text"), ("Days left", "derived", "count"),
                    ("Owner", "input", "text"), ("Phase", "input", "text"),
                    ("Notes", "input", "text"),
                )),
                _table("Dashboard", "B11:F18", _columns(
                    ("Area", "derived", "text"), ("Planned", "derived", money),
                    ("Actual spend", "derived", money), ("Remaining", "derived", money),
                    ("Status", "derived", "text"),
                ), total_ref="B19:F19", total_label="TOTAL PROJECT"),
            ],
            "enum_contract": [
                {"sheet": "Purchases", "range": "C8:C105", "column": "Area", "values": ["Kitchen units", "Flooring", "Lighting", "Plumbing", "Paint", "Permits", "Contingency"]},
                {"sheet": "Quotes", "range": "C8:C105", "column": "Area", "values": ["Kitchen units", "Flooring", "Lighting", "Plumbing", "Paint", "Permits", "Contingency"]},
                {"sheet": "Quotes", "range": "I8:I105", "column": "Decision", "values": ["Preferred", "Backup", "Review"]},
                {"sheet": "Timeline", "range": "C8:C105", "column": "Area", "values": ["Kitchen units", "Flooring", "Lighting", "Plumbing", "Paint", "Permits", "Contingency"]},
                {"sheet": "Purchases", "range": "G8:G105", "column": "Paid?", "values": ["Yes", "No"]},
                {"sheet": "Purchases", "range": "H8:H105", "column": "Status", "values": ["Planned", "Ordered", "Purchased"]},
                {"sheet": "Quotes", "range": "F8:F105", "column": "Selected", "values": ["Yes", "No"]},
                {"sheet": "Timeline", "range": "F8:F105", "column": "Status", "values": ["Open", "In progress", "Done"]},
            ],
            "input_text_policy": _master_input_text_policy(),
            "formula_rules": [
                _formula("Budget", "D8", "SUMIF", "Actual spend is the sum of logged purchases for the matching area."),
                _formula("Budget", "E8", "C8-D8", "Remaining is planned amount minus actual spend."),
                _formula("Dashboard", "B7", "Settings!C9", "The dashboard budget cap follows the Settings input."),
                _formula("Dashboard", "E7", "SUM(Purchases!F8:F14)", "Actual spend follows the Purchases table."),
                _formula("Dashboard", "K7", "B7-E7", "Available budget is the cap less actual spend."),
                _formula("Timeline", "G8", "MAX(E8-TODAY(),0)", "Days left never becomes negative."),
            ],
            "relationships": [
                "Purchases is the only operational source for Budget.Actual spend.",
                "Purchases.Area must match Budget.Area so each purchase is attributed once.",
                "Dashboard derives its budget cap, planned work, actual spend and available budget from Settings and Budget.",
                "Timeline.Status is an input and Timeline.Days left is derived from Due date.",
                "Contingency % is a ratio and never represents an amount, count or date.",
                "Rows labelled TOTAL remain outside the editable and filterable table ranges.",
            ],
            "cases": [
                _case(
                    "renovation spend updates the control panel",
                    [
                        {"sheet": "Settings", "cell": "C9", "value": 24000},
                        {"sheet": "Purchases", "cell": "F8", "value": 3500},
                    ],
                    [
                        {"sheet": "Dashboard", "cell": "B7", "operator": "equals", "value": 24000},
                        {"sheet": "Budget", "cell": "D8", "operator": "equals", "value": 3500},
                        {"sheet": "Budget", "cell": "E8", "operator": "equals", "value": 5500},
                        {"sheet": "Dashboard", "cell": "E7", "operator": "equals", "value": 7400},
                        {"sheet": "Dashboard", "cell": "K7", "operator": "equals", "value": 16600},
                    ],
                    "Comprueba que una nueva compra actualiza su partida, el gasto acumulado y el presupuesto disponible.",
                ),
            ],
            "failure_criteria": [
                "Una compra no aparece en el Actual spend de su misma Area.",
                "Budget.Remaining deja de ser Planned menos Actual spend.",
                "Budget cap, actual spend o available budget no se actualizan en Dashboard.",
                "Contingency % usa formato de importe o una columna de importe usa formato porcentaje.",
                "Days left puede ser negativo o Status se calcula desde una fÃ³rmula que impide editarlo.",
                "Una fila TOTAL entra dentro del rango de la tabla nativa.",
            ],
        },
    }
    freelance = deepcopy(contracts[10])
    freelance.update({
        "contract_id": "freelance-client-invoice-tracker-v1",
        "title": "Freelance Client, Project & Invoice Tracker",
        "product_id": 13,
        "purpose": "Registrar ingresos de clientes, facturas, costes recurrentes y previsiones de caja para decidir el siguiente trabajo con visibilidad.",
        "relationships": [
            "Transactions is the operational source for cash in and cash out.",
            "Invoices describe expected client income before it enters Transactions.",
            "Forecast chains each month's opening balance from the previous closing balance.",
            "Year end keeps the total row outside the editable table.",
        ],
        "enum_contract": [
            {"sheet": "Transactions", "range": "F6:F105", "column": "Status", "values": ["Cleared", "Expected", "Planned", "Reserved"]},
        ],
    })
    contracts[13] = freelance
    creator = deepcopy(contracts[12])
    creator.update({
        "contract_id": "creator-content-sponsorship-tracker-v1",
        "title": "Creator Content & Sponsorship Tracker",
        "product_id": 14,
        "purpose": "Probar una variante de creator business sobre las mismas ranuras nativas de presupuesto, compras, comparativas y timeline del master.",
        "relationships": [
            "Purchases is the operational source for creator production costs.",
            "Quotes describe sponsorship or supplier alternatives before commitment.",
            "Dashboard derives the main creator budget and available amount from Settings and Purchases.",
            "Timeline keeps publishing and delivery dates outside the editable totals.",
        ],
        "enum_contract": [
            {"sheet": "Purchases", "range": "C8:C105", "column": "Area", "values": ["Brand deals", "Video production", "Editing", "Software", "Studio", "Travel", "Tax reserve"]},
            {"sheet": "Quotes", "range": "C8:C105", "column": "Area", "values": ["Brand deals", "Video production", "Editing", "Software", "Studio", "Travel", "Tax reserve"]},
            {"sheet": "Quotes", "range": "I8:I105", "column": "Decision", "values": ["Preferred", "Backup", "Review"]},
            {"sheet": "Timeline", "range": "C8:C105", "column": "Area", "values": ["Brand deals", "Video production", "Editing", "Software", "Studio", "Travel", "Tax reserve"]},
            {"sheet": "Purchases", "range": "G8:G105", "column": "Paid?", "values": ["Yes", "No"]},
            {"sheet": "Purchases", "range": "H8:H105", "column": "Status", "values": ["Planned", "Ordered", "Purchased"]},
            {"sheet": "Quotes", "range": "F8:F105", "column": "Selected", "values": ["Yes", "No"]},
            {"sheet": "Timeline", "range": "F8:F105", "column": "Status", "values": ["Open", "In progress", "Done"]},
        ],
    })
    creator["cases"][0]["name"] = "creator production spend updates the control panel"
    creator["cases"][0]["purpose"] = (
        "Comprueba que un coste de producción de creator actualiza su área, el gasto acumulado "
        "y el presupuesto disponible."
    )
    contracts[14] = creator
    aquarium = deepcopy(creator)
    aquarium.update({
        "contract_id": "freshwater-aquarium-setup-planner-v1",
        "title": "Freshwater Aquarium Setup Planner",
        "product_id": 15,
        "purpose": "Planificar el presupuesto y las compras de un acuario de agua dulce, comparar equipos y ordenar las tareas de montaje antes de introducir animales.",
        "relationships": [
            "Purchases is the operational source for aquarium setup costs.",
            "Purchases.Area must match Budget.Area so equipment and setup costs are attributed once.",
            "Quotes compare aquarium equipment and services before a purchase is committed.",
            "Timeline keeps setup, equipment checks and readiness milestones visible while Days left remains derived.",
            "Rows labelled TOTAL remain outside editable and filterable table ranges.",
        ],
    })
    aquarium_areas = ["Tank and stand", "Filtration", "Lighting and heat", "Substrate and hardscape", "Plants", "Water testing", "Contingency"]
    for enum in aquarium["enum_contract"]:
        if enum.get("column") == "Area":
            enum["values"] = aquarium_areas
    aquarium["cases"][0]["name"] = "aquarium equipment spend updates the control panel"
    aquarium["cases"][0]["purpose"] = "Comprueba que el coste de un filtro atribuido a su partida actualiza gasto real y presupuesto disponible sin duplicarse."
    contracts[15] = aquarium
    puppy = deepcopy(aquarium)
    puppy.update({
        "contract_id": "puppy-first-year-planner-v1",
        "title": "New Puppy Planner & First-Year Budget",
        "product_id": 16,
        "purpose": "Organizar el presupuesto, las compras, las preguntas para profesionales y los hitos del primer año de un cachorro sin sustituir el consejo veterinario.",
        "relationships": [
            "Purchases is the operational source for puppy setup and first-year costs.",
            "Purchases.Area must match Budget.Area so each cost is attributed once.",
            "Quotes compare local veterinary, training and grooming services before a choice is made.",
            "Timeline records owner-managed milestones and does not prescribe a medical schedule.",
            "Rows labelled TOTAL remain outside editable and filterable table ranges.",
        ],
    })
    puppy_areas = ["Adoption and setup", "Food", "Veterinary care", "Training", "Grooming", "Equipment", "Contingency"]
    for enum in puppy["enum_contract"]:
        if enum.get("column") == "Area":
            enum["values"] = puppy_areas
    puppy["cases"][0]["name"] = "puppy setup spending updates the control panel"
    puppy["cases"][0]["purpose"] = "Comprueba que un gasto de preparación del cachorro actualiza la partida, el gasto acumulado y el presupuesto restante sin duplicarse."
    contracts[16] = puppy
    vehicle = deepcopy(puppy)
    vehicle.update({
        "contract_id": "vehicle-maintenance-cost-planner-v1",
        "title": "Vehicle Maintenance Spreadsheet & Car Cost Tracker Kit",
        "product_id": 17,
        "purpose": "Organizar costes introducidos por el usuario, presupuestos de talleres y eventos de servicio sin prescribir intervalos mecánicos ni sustituir asesoramiento profesional.",
        "relationships": [
            "Purchases is the operational source for vehicle ownership costs.",
            "Purchases.Area must match Budget.Area so each cost is attributed once.",
            "Quotes compare garage or parts estimates before a choice is made.",
            "Timeline records owner-entered dates and service events; it does not prescribe maintenance intervals.",
            "Rows labelled TOTAL remain outside editable and filterable table ranges.",
        ],
    })
    vehicle_areas = ["Fuel", "Insurance", "Routine service", "Repairs and parts", "Registration and fees", "Tyres", "Contingency"]
    for enum in vehicle["enum_contract"]:
        if enum.get("column") == "Area":
            enum["values"] = vehicle_areas
    vehicle["cases"][0]["name"] = "vehicle ownership cost updates the control panel"
    vehicle["cases"][0]["purpose"] = "Comprueba que un gasto de mantenimiento actualiza su partida, el gasto acumulado y el presupuesto restante sin duplicarse."
    contracts[17] = vehicle
    book = deepcopy(contracts[12])
    book.update({
        "contract_id": "community-garden-planner-v1",
        "title": "Community Garden Budget & Project Planner",
        "product_id": 1,
        "purpose": "Conectar costes de jardin comunitario, proveedores, decisiones compartidas y fechas de temporada desde una unica base de proyecto.",
    })
    book_areas = ["Site preparation", "Soil and compost", "Plants and seeds", "Irrigation", "Tools and storage", "Workshops", "Contingency"]
    for enum in book["enum_contract"]:
        if enum.get("column") == "Area":
            enum["values"] = book_areas
    book["relationships"] = [
        "Purchases is the operational source for community-garden spend.",
        "Purchases.Area must match Budget.Area so each garden cost is attributed once.",
        "Quotes compare suppliers before a materials, equipment or workshop decision is committed.",
        "Timeline keeps seasonal garden deadlines visible while Days left remains derived.",
        "Rows labelled TOTAL remain outside editable and filterable table ranges.",
    ]
    book["cases"][0]["name"] = "garden spend updates the control panel"
    book["cases"][0]["purpose"] = "Comprueba que un coste del jardin actualiza su partida, el gasto acumulado y el presupuesto disponible."
    contracts[1] = book
    move = deepcopy(book)
    move.update({
        "contract_id": "moving-home-planner-v1",
        "title": "Moving Home Planner & Budget",
        "product_id": 2,
        "purpose": "Conectar el presupuesto de mudanza, costes reales, presupuestos de empresas y tareas con fecha desde la preparación hasta la entrega de llaves.",
    })
    move_areas = ["Removal company", "Packing materials", "Cleaning", "Utility setup", "Storage", "Travel and meals", "Contingency"]
    for enum in move["enum_contract"]:
        if enum.get("column") == "Area":
            enum["values"] = move_areas
    move["relationships"] = [
        "Purchases is the single operational source for moving costs.",
        "Purchases.Area must match Budget.Area so each cost is attributed once.",
        "Quotes compare removal and setup services before a booking is committed.",
        "Timeline keeps handover and move dates visible while Days left remains derived.",
        "Rows labelled TOTAL remain outside editable and filterable table ranges.",
    ]
    move["cases"][0]["name"] = "moving costs update the control panel"
    move["cases"][0]["purpose"] = "Comprueba que un coste de mudanza actualiza su partida, el gasto acumulado y el presupuesto disponible."
    contracts[2] = move
    from project_model import extend_contract
    for contract in contracts.values():
        extend_contract(contract)
    return contracts


def contract_for_product(product_id: int) -> dict[str, Any]:
    """Return a defensive copy of the authoritative contract for a product."""
    contract = _product_contracts().get(int(product_id))
    if not contract:
        raise KeyError(f"No hay contrato de modelo declarado para product-{product_id}")
    return deepcopy(contract)


def contract_for_slug(slug: str) -> dict[str, Any] | None:
    normalized = str(slug).casefold()
    if "presupuesto-familiar" in normalized or "family-budget" in normalized:
        return family_base_contract()
    if "planificador-de-boda-base" in normalized:
        return wedding_base_contract()
    if "home-renovation" in normalized or "reforma" in normalized:
        return contract_for_product(12)
    if "wedding" in normalized or "boda" in normalized:
        return contract_for_product(8)
    return None


def validate_contract_schema(contract: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    for key in ("contract_id", "title", "purpose", "sheets", "tables", "relationships", "failure_criteria", "cases"):
        if not contract.get(key):
            errors.append(f"falta {key} en el contrato de modelo")
    if len(contract.get("sheets") or []) != len(set(contract.get("sheets") or [])):
        errors.append("sheets contiene duplicados")
    for index, table in enumerate(contract.get("tables") or [], 1):
        prefix = f"tabla declarada {index}"
        for key in ("sheet", "ref", "headers"):
            if not table.get(key):
                errors.append(f"{prefix}: falta {key}")
        headers = table.get("headers") or []
        names = [str(item.get("name", "")).strip() for item in headers]
        if not names or any(not name for name in names):
            errors.append(f"{prefix}: hay columnas sin nombre")
        if len(names) != len(set(name.casefold() for name in names)):
            errors.append(f"{prefix}: hay columnas duplicadas")
        for column in headers:
            if column.get("role") not in {"input", "derived"}:
                errors.append(f"{prefix}: role no válido en {column.get('name')!r}")
            if column.get("kind") not in {"text", "general", "amount", "count", "percent", "date"}:
                errors.append(f"{prefix}: kind no válido en {column.get('name')!r}")
        try:
            min_col, min_row, max_col, max_row = range_boundaries(str(table.get("ref")))
            if max_col - min_col + 1 != len(headers) or max_row <= min_row:
                errors.append(f"{prefix}: ref no coincide con el número de columnas o filas")
        except ValueError:
            errors.append(f"{prefix}: ref inválido {table.get('ref')!r}")
    if int(contract.get("product_id") or 0) >= 12:
        policy = contract.get("input_text_policy") or []
        if not policy:
            errors.append("los productos derivados del master deben clasificar sus campos de texto como enum o free_text")
        policy_keys: set[tuple[str, str]] = set()
        policy_modes: dict[tuple[str, str], str] = {}
        for index, item in enumerate(policy, 1):
            key = (str(item.get("sheet", "")), str(item.get("column", "")))
            mode = str(item.get("mode", ""))
            if not key[0] or not key[1] or mode not in {"enum", "free_text"}:
                errors.append(f"text_field_policy {index}: cada campo debe declarar sheet, column y mode enum/free_text")
            if key in policy_keys:
                errors.append(f"text_field_policy: campo duplicado {key[0]}!{key[1]}")
            policy_keys.add(key)
            policy_modes[key] = mode
        expected_text_keys: set[tuple[str, str]] = set()
        for table in contract.get("tables") or []:
            for header in table.get("headers") or []:
                if header.get("role") == "input" and header.get("kind") == "text":
                    expected_text_keys.add((str(table.get("sheet")), str(header.get("name"))))
        missing_policy = sorted(expected_text_keys - policy_keys)
        extra_policy = sorted(policy_keys - expected_text_keys)
        if missing_policy:
            errors.append(f"text_field_policy: faltan campos {missing_policy}")
        if extra_policy:
            errors.append(f"text_field_policy: campos no declarados en tablas {extra_policy}")
        enum_keys = {(str(item.get("sheet")), str(item.get("column"))) for item in contract.get("enum_contract") or []}
        for key in expected_text_keys & policy_keys:
            if policy_modes.get(key) == "enum" and key not in enum_keys:
                errors.append(f"text_field_policy: {key[0]}!{key[1]} está marcado como enum pero no tiene enum_contract")
            if policy_modes.get(key) == "free_text" and key in enum_keys:
                errors.append(f"text_field_policy: {key[0]}!{key[1]} está marcado como free_text pero tiene enum_contract")
    for index, case in enumerate(contract.get("cases") or [], 1):
        if not case.get("name") or not case.get("purpose"):
            errors.append(f"caso {index}: falta name o purpose")
        if not case.get("assertions"):
            errors.append(f"caso {index}: no tiene resultados esperados")
    for index, enum in enumerate(contract.get("enum_contract") or [], 1):
        prefix = f"enum declarado {index}"
        for key in ("sheet", "range", "column", "values"):
            if not enum.get(key):
                errors.append(f"{prefix}: falta {key}")
        if not isinstance(enum.get("values"), list) or not enum.get("values") or len(enum["values"]) != len(set(str(value) for value in enum["values"])):
            errors.append(f"{prefix}: values debe ser una lista no vacia y sin duplicados")
    return errors


def render_model_behavior_markdown(contract: dict[str, Any]) -> str:
    """Render the normative user/QA guide from the declared contract only."""
    errors = validate_contract_schema(contract)
    if errors:
        raise ValueError("Contrato inválido: " + "; ".join(errors))
    lines = [
        f"# Contrato funcional del modelo · {contract['title']}",
        "",
        "> Este documento es normativo: describe lo que la plantilla debe hacer, no lo que se observó en un XLSX ya construido.",
        "> Si el libro no cumple una regla, el libro falla la verificación y debe corregirse el diseño o la implementación.",
        "",
        "## Guía rápida para el usuario",
        "",
        f"**Objetivo:** {contract['purpose']}",
        "",
        "1. Introduce o modifica únicamente las columnas marcadas como `input`.",
        "2. No sobrescribas columnas `derived`: contienen fórmulas y deben actualizarse.",
        "3. Usa las listas y estados declarados por la ficha del producto; no escribas porcentajes en columnas de cantidades o importes.",
        "4. Comprueba el resultado en las hojas de resumen después de añadir un registro.",
        "",
        "## Hojas que debe contener el libro",
        "",
        " · ".join(f"`{sheet}`" for sheet in contract["sheets"]),
        "",
        "## Tablas y columnas esperadas",
        "",
        "Cada rango es una tabla de datos. Las filas de total indicadas después del rango deben quedar fuera del ListObject y fuera del filtro de entradas.",
        "",
        "| Hoja | Rango de datos | Columnas esperadas |",
        "|---|---|---|",
    ]
    for table in contract["tables"]:
        headers = "; ".join(
            f"`{column['name']}` ({column['role']}, {column['kind']})"
            for column in table["headers"]
        )
        total = ""
        if table.get("total_ref"):
            total = f" Total fuera: `{table['total_ref']}`"
            if table.get("total_label"):
                total += f" (`{table['total_label']}`)."
        lines.append(f"| `{table['sheet']}` | `{table['ref']}` | {headers}.{total} |")
    lines += ["", "## Reglas de cálculo", ""]
    for relationship in contract["relationships"]:
        lines.append(f"- {relationship}")
    lines += ["", "### Fórmulas que deben existir", ""]
    for rule in contract.get("formula_rules") or []:
        lines.append(f"- `{rule['sheet']}!{rule['cell']}` debe contener `{rule['contains']}`: {rule['description']}")
    lines += ["", "## Listas y enums permitidos", "", "Las validaciones desplegables del XLSX deben coincidir exactamente con estas listas; no se acepta una lista distinta para la misma columna.", ""]
    lines.append("| Hoja | Rango | Columna | Valores permitidos |")
    lines.append("|---|---|---|---|")
    for enum in contract.get("enum_contract") or []:
        lines.append(f"| `{enum['sheet']}` | `{enum['range']}` | `{enum['column']}` | " + ", ".join(f"`{value}`" for value in enum["values"]) + " |")
    lines += ["", "## Casos de verificación", ""]
    lines.append("Los casos se ejecutan sobre una copia temporal del libro. El archivo original no se modifica.")
    lines.append("")
    for index, case in enumerate(contract.get("cases") or [], 1):
        lines.append(f"### Caso {index}: {case['name']}")
        lines.append("")
        lines.append(case["purpose"])
        lines.append("")
        lines.append("**Entrada:** " + ", ".join(f"`{item['sheet']}!{item['cell']} = {item['value']}`" for item in case["inputs"]))
        lines.append("")
        lines.append("**Resultados esperados:**")
        for action in case.get("actions", []):
            lines.append(f"- Table operation: `{action}`")
        for assertion in case["assertions"]:
            lines.append(f"- `{assertion['sheet']}!{assertion['cell']}` {assertion.get('operator', 'equals')} `{assertion.get('value', '')}`")
        lines.append("")
    lines += ["## Criterios de fallo", ""]
    lines.extend(f"- {item}" for item in contract["failure_criteria"])
    lines += [
        "",
        "## Regla de independencia",
        "",
        "Este documento se actualiza desde el contrato de diseño declarado antes de generar el XLSX. No se debe redactar leyendo hacia atrás el libro generado ni cambiar el resultado esperado para que coincida con un error observado.",
        "",
    ]
    return "\n".join(lines)


def write_model_behavior_markdown(path: Path, contract: dict[str, Any]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(render_model_behavior_markdown(contract), encoding="utf-8")
    return path


def render_app_usage_markdown(spec: dict[str, Any], contract: dict[str, Any]) -> str:
    """Render the customer-facing workbook instructions from the declaration."""
    title = spec.get("title") or contract.get("title") or "Excel workbook"
    language = str(spec.get("design_language") or "en").casefold()
    language_label = {"en": "English", "es": "Español"}.get(language, language)
    lines = [
        f"# How to use · {title}",
        "",
        f"**Language:** `{language}` ({language_label})",
        "",
        "> This guide is generated from the expected product behaviour and is included with the XLSX. A reader-friendly PDF copy is provided alongside this Markdown file.",
        "",
        "## Quick start",
        "",
        "1. Open the `Start here` sheet and read the project notes.",
        "2. Go to `Settings` and enter the project assumptions, dates and lists.",
        "3. Enter new records only in cells marked as inputs or in input columns.",
        "4. Use the dropdowns instead of creating new spellings for statuses or categories.",
        "5. Review `Dashboard` after each change to confirm that connected totals update.",
        "",
        "## Workbook flow",
        "",
        " → ".join(f"`{sheet}`" for sheet in contract.get("sheets", [])),
        "",
        "## What can be edited",
        "",
    ]
    ebook = spec.get("ebook") or {}
    if ebook.get("filename"):
        lines[5:5] = [
            "## Companion e-book",
            "",
            f"The product package also includes `{ebook['filename']}` as a topic-specific PDF companion.",
            "",
        ]
    for table in contract.get("tables", []):
        editable = [column["name"] for column in table.get("headers", []) if column.get("role") == "input"]
        derived = [column["name"] for column in table.get("headers", []) if column.get("role") == "derived"]
        lines.append(f"### {table['sheet']}")
        lines.append("")
        lines.append(f"- Data range: `{table['ref']}`")
        lines.append(f"- Edit: {', '.join(f'`{item}`' for item in editable) or 'no direct inputs'}")
        lines.append(f"- Do not overwrite: {', '.join(f'`{item}`' for item in derived) or 'no derived columns'}")
        if table.get("total_ref"):
            lines.append(f"- Total row: `{table['total_ref']}` remains outside the data table.")
        lines.append("")
    lines += [
        "## Connected behaviour",
        "",
    ]
    lines.extend(f"- {item}" for item in contract.get("relationships", []))
    lines += [
        "",
        "## Dropdowns and data entry",
        "",
        "- Select an existing enum value from the dropdown; do not type a second variant of the same value.",
        "- Keep dates as real Excel dates and amounts as numbers.",
        "- Keep notes and descriptions in their designated free-text columns.",
        "- Start with the -blank.xlsx copy; the -excel-template.xlsx file contains worked examples.",
        "- Add purchases with Insert > Table Rows inside the Purchases table so totals and formatting expand together. Do not type records below the total row.",
        "- In Quotes, mark Selected = Yes and Decision = Preferred to include an offer in Budget.Forecast for the same Area. This is still an estimate: record a real payment separately in Purchases; Dashboard.Actual spend only includes that log.",
        "- Negative amounts represent refunds. Keep an existing area selected so costs remain assigned to the budget.",
        "",
        "## Verification",
        "",
        "The factory checks the formulas, table ranges, dropdowns, dates and declared functional cases before packaging the workbook.",
        "",
    ]
    for case in contract.get("cases", []):
        lines.append(f"- `{case['name']}`: {case['purpose']}")
    lines += ["", "If a calculated result does not update, check the source input and do not replace the formula with a manual value.", ""]
    return "\n".join(lines)


def write_app_usage_markdown(path: Path, spec: dict[str, Any], contract: dict[str, Any]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(render_app_usage_markdown(spec, contract), encoding="utf-8")
    return path


def write_model_contract_json(path: Path, contract: dict[str, Any]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(contract, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return path


def load_model_contract(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))

