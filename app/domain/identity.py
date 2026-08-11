"""Politicas de identificacao que evitam expor a matricula completa."""

import re
from typing import Any


def obter_final_matricula(valor: Any) -> str:
    """Retorna os tres ultimos digitos da matricula, sempre como texto."""
    if valor is None or valor == "":
        return ""

    digitos = re.sub(r"\D", "", str(valor).strip())
    return digitos[-3:]
