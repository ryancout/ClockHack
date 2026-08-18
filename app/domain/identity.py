"""Políticas de normalização da identificação funcional."""

from typing import Any


def normalizar_matricula(valor: Any) -> str:
    """Mantém a matrícula completa como texto, inclusive zeros à esquerda."""
    if valor is None or valor == "":
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()
