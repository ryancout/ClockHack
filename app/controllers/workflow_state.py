"""Estado mínimo compartilhado pelos fluxos da aplicação."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class WorkflowState:
    preferences: dict[str, Any]
    selected_files: list[str] = field(default_factory=list)
    last_result: dict[str, Any] | None = None


__all__ = ["WorkflowState"]
