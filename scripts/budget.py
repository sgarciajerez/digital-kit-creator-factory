from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from factory_core import initialize_database, managed_connection, record_expense


class BudgetExceededError(RuntimeError):
    """Señala que una tarea no puede iniciar una llamada de pago."""


@dataclass(frozen=True)
class BudgetSnapshot:
    limit_eur: float
    spent_eur: float
    remaining_eur: float
    percentage: float
    alerts: tuple[int, ...]


def _today_utc() -> str:
    return datetime.now(timezone.utc).date().isoformat()


class BudgetManager:
    def __init__(
        self,
        database_path: Path,
        monthly_limit_eur: float,
        daily_limit_eur: float | None = None,
        task_limit_eur: float | None = None,
        product_limit_eur: float | None = None,
        batch_limit_eur: float | None = None,
    ) -> None:
        limits = {
            "monthly": monthly_limit_eur,
            "daily": daily_limit_eur,
            "task": task_limit_eur,
            "product": product_limit_eur,
            "batch": batch_limit_eur,
        }
        for name, value in limits.items():
            if value is not None and value < 0:
                raise ValueError(f"El límite {name} no puede ser negativo")
        self.database_path = database_path
        initialize_database(database_path)
        self.monthly_limit_eur = float(monthly_limit_eur)
        self.daily_limit_eur = daily_limit_eur
        self.task_limit_eur = task_limit_eur
        self.product_limit_eur = product_limit_eur
        self.batch_limit_eur = batch_limit_eur

    def _spent(self, where: str = "", parameters: tuple[Any, ...] = ()) -> float:
        with managed_connection(self.database_path) as connection:
            query = (
                "SELECT COALESCE(SUM(amount), 0) FROM expenses "
                "WHERE category = 'ai' AND currency = 'EUR'"
            )
            if where:
                query += f" AND {where}"
            return float(connection.execute(query, parameters).fetchone()[0] or 0)

    def snapshot(self) -> BudgetSnapshot:
        spent = self._spent("date >= strftime('%Y-%m-01', 'now')")
        percentage = 0.0 if self.monthly_limit_eur == 0 else spent / self.monthly_limit_eur * 100
        alerts = tuple(level for level in (50, 75, 90, 100) if percentage >= level)
        return BudgetSnapshot(
            limit_eur=self.monthly_limit_eur,
            spent_eur=spent,
            remaining_eur=max(self.monthly_limit_eur - spent, 0.0),
            percentage=percentage,
            alerts=alerts,
        )

    def _scope_limit(self, scope: str | None) -> float | None:
        return {
            "task": self.task_limit_eur,
            "product": self.product_limit_eur,
            "batch": self.batch_limit_eur,
        }.get(scope)

    def check(
        self,
        estimated_cost_eur: float,
        scope: str | None = None,
        scope_id: str | None = None,
    ) -> BudgetSnapshot:
        estimated = float(estimated_cost_eur)
        if estimated < 0:
            raise ValueError("El coste estimado no puede ser negativo")
        if scope not in {None, "task", "product", "batch"}:
            raise ValueError("scope debe ser task, product, batch o None")
        if scope is not None and not scope_id:
            raise ValueError("scope_id es obligatorio cuando se usa scope")

        snapshot = self.snapshot()
        if snapshot.spent_eur + estimated > self.monthly_limit_eur:
            raise BudgetExceededError("El límite mensual de IA bloquearía esta llamada")

        if self.daily_limit_eur is not None:
            daily = self._spent("date = ?", (_today_utc(),))
            if daily + estimated > self.daily_limit_eur:
                raise BudgetExceededError("El límite diario de IA bloquearía esta llamada")

        scope_limit = self._scope_limit(scope)
        if scope_limit is not None:
            scoped = self._spent("scope = ? AND scope_id = ?", (scope, scope_id))
            if scoped + estimated > scope_limit:
                raise BudgetExceededError(f"El límite de IA para {scope} bloquearía esta llamada")
        return snapshot

    def record_cost(
        self,
        provider: str,
        model: str,
        units: float,
        estimated_cost_eur: float,
        actual_cost_eur: float,
        scope: str | None = None,
        scope_id: str | None = None,
        notes: str | None = None,
    ) -> int:
        self.check(estimated_cost_eur, scope=scope, scope_id=scope_id)
        actual = float(actual_cost_eur)
        if actual < 0:
            raise ValueError("El coste real no puede ser negativo")
        if actual > estimated_cost_eur:
            self.check(actual, scope=scope, scope_id=scope_id)
        return record_expense(
            self.database_path,
            amount=actual,
            category="ai",
            provider=provider,
            provider_model=model,
            units=units,
            cost_estimate=estimated_cost_eur,
            cost_actual=actual,
            scope=scope,
            scope_id=scope_id,
            notes=notes,
        )
