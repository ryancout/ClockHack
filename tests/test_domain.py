from dataclasses import FrozenInstanceError

import pytest

from app.domain import RegistroFuncionario, obter_final_matricula


@pytest.mark.parametrize(
    ("matricula", "esperado"),
    [
        ("123001", "001"),
        ("MAT-0002", "002"),
        ("000", "000"),
        ("7", "7"),
        ("sem-digitos", ""),
        ("", ""),
        (None, ""),
    ],
)
def test_obter_final_matricula(matricula, esperado):
    assert obter_final_matricula(matricula) == esperado


def test_registro_funcionario_armazena_minutos_sem_formatacao():
    registro = RegistroFuncionario(
        nome="Ana Teste",
        final_matricula="001",
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
