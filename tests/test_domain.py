from dataclasses import FrozenInstanceError

import pytest

from app.domain import RegistroFuncionario, normalizar_matricula


@pytest.mark.parametrize(
    ("matricula", "esperado"),
    [
        ("123001", "123001"),
        ("MAT-0002", "MAT-0002"),
        ("000", "000"),
        ("7", "7"),
        ("sem-digitos", "sem-digitos"),
        ("", ""),
        (None, ""),
    ],
)
def test_normalizar_matricula_preserva_identificador_completo(matricula, esperado):
    assert normalizar_matricula(matricula) == esperado


def test_registro_funcionario_armazena_minutos_sem_formatacao():
    registro = RegistroFuncionario(
        nome="Ana Teste",
        matricula="MAT-0001",
        departamento="Operacoes",
        banco_total_minutos=-90,
        banco_saldo_minutos=1530,
        faltas="",
    )

    assert registro.banco_total_minutos == -90
    assert registro.banco_saldo_minutos == 1530
    assert not hasattr(registro, "banco_total_fmt")


def test_registro_funcionario_e_imutavel_e_usa_slots():
    registro = RegistroFuncionario("Ana", "001", "Operacoes", 60, 30, "")

    with pytest.raises(FrozenInstanceError):
        registro.nome = "Outro nome"

    assert not hasattr(registro, "__dict__")
