from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Literal, TypeVar

from pydantic import BaseModel, ConfigDict, Field

from budget import BudgetManager


class ProviderConfigurationError(RuntimeError):
    pass


class AIRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    task: str = Field(min_length=1)
    prompt: str = Field(min_length=1)
    model: str = "mock"
    max_output_tokens: int = Field(default=512, ge=1, le=4096)
    estimated_cost_eur: float = Field(default=0.0, ge=0)
    scope: Literal["task", "product", "batch"] | None = None
    scope_id: str | None = None


class AIResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")

    provider: str
    model: str
    text: str
    output: dict[str, Any]
    input_units: int = Field(ge=0)
    output_units: int = Field(ge=0)
    estimated_cost_eur: float = Field(ge=0)
    actual_cost_eur: float = Field(ge=0)


class AIProvider(ABC):
    @abstractmethod
    def complete(self, request: AIRequest) -> AIResponse:
        raise NotImplementedError


class MockProvider(AIProvider):
    """Proveedor local determinista; nunca realiza una llamada de red."""

    def __init__(self, budget: BudgetManager, model: str = "mock") -> None:
        self.budget = budget
        self.model = model or "mock"

    def complete(self, request: AIRequest) -> AIResponse:
        model = request.model or self.model
        self.budget.check(
            request.estimated_cost_eur,
            scope=request.scope,
            scope_id=request.scope_id,
        )
        text = f"Simulación local para la tarea: {request.task}"
        output = {
            "status": "simulated",
            "task": request.task,
            "prompt_received": True,
        }
        input_units = len(request.prompt.split())
        output_units = len(text.split())
        self.budget.record_cost(
            provider="mock",
            model=model,
            units=input_units + output_units,
            estimated_cost_eur=request.estimated_cost_eur,
            actual_cost_eur=0.0,
            scope=request.scope,
            scope_id=request.scope_id,
            notes="Proveedor mock local; coste real cero",
        )
        return AIResponse(
            provider="mock",
            model=model,
            text=text,
            output=output,
            input_units=input_units,
            output_units=output_units,
            estimated_cost_eur=request.estimated_cost_eur,
            actual_cost_eur=0.0,
        )


OutputModel = TypeVar("OutputModel", bound=BaseModel)


def validate_structured_output(response: AIResponse, output_model: type[OutputModel]) -> OutputModel:
    """Valida la parte estructurada con el modelo Pydantic solicitado."""
    return output_model.model_validate(response.output)


def build_provider(settings: Any, budget: BudgetManager) -> AIProvider:
    provider = settings.ai_provider.lower()
    if provider == "mock":
        return MockProvider(budget, model=settings.ai_model or "mock")
    raise ProviderConfigurationError(
        f"Proveedor no implementado: {settings.ai_provider}. "
        "No se realizan llamadas externas automáticamente."
    )
