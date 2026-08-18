"""Modelos independentes das camadas de interface e infraestrutura."""

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class RegistroFuncionario:
    """Dados de um funcionario usados para compor os relatorios."""

    nome: str
    matricula: str
    departamento: str | None
    banco_total_minutos: int
    banco_saldo_minutos: int
    faltas: str | None
