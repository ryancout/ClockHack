"""Fronteira de publicação que isola o modelo Push do fluxo da aplicação."""

from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass
from enum import Enum
from typing import Any, Protocol, runtime_checkable

from app.integrations.powerbi_client import PowerBiClient


PUSH_NEW_MODEL_SUPPORT_END = "2027-10-31"


class PowerBiDestinationKind(str, Enum):
    PUSH_SEMANTIC_MODEL = "push_semantic_model"
    FABRIC_LAKEHOUSE = "fabric_lakehouse"


@dataclass(frozen=True, slots=True)
class PowerBiPublishResult:
    destination: PowerBiDestinationKind
    resource_id: str
    row_count: int


@runtime_checkable
class PowerBiDestination(Protocol):
    """Contrato que o futuro destino Fabric deverá implementar."""

    kind: PowerBiDestinationKind

    @property
    def authenticated(self) -> bool: ...

    def login(self) -> None: ...

    def publish(
        self,
        rows: Iterable[dict[str, Any]],
        *,
        on_progress: Callable[[int, int], None] | None = None,
    ) -> PowerBiPublishResult: ...


class _PushClient(Protocol):
    @property
    def autenticado(self) -> bool: ...

    def login_interativo(self) -> None: ...

    def obter_ou_criar_dataset(self) -> str: ...

    def enviar_linhas(
        self,
        dataset_id: str,
        linhas: Iterable[dict[str, Any]],
        *,
        ao_progresso: Callable[[int, int], None] | None = None,
    ) -> int: ...


class PushSemanticModelDestination:
    """Adaptador temporário para o modelo semântico Push atualmente usado."""

    kind = PowerBiDestinationKind.PUSH_SEMANTIC_MODEL

    def __init__(self, client: _PushClient | None = None) -> None:
        self.client = client or PowerBiClient()

    @property
    def authenticated(self) -> bool:
        return self.client.autenticado

    def login(self) -> None:
        self.client.login_interativo()

    def publish(
        self,
        rows: Iterable[dict[str, Any]],
        *,
        on_progress: Callable[[int, int], None] | None = None,
    ) -> PowerBiPublishResult:
        dataset_id = self.client.obter_ou_criar_dataset()
        row_count = self.client.enviar_linhas(
            dataset_id,
            rows,
            ao_progresso=on_progress,
        )
        return PowerBiPublishResult(self.kind, dataset_id, row_count)


__all__ = [
    "PUSH_NEW_MODEL_SUPPORT_END",
    "PowerBiDestination",
    "PowerBiDestinationKind",
    "PowerBiPublishResult",
    "PushSemanticModelDestination",
]
