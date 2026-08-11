from openpyxl import Workbook

from app.domain import RegistroFuncionario
from app.reports import criar_aba_ranking


def _registro(nome, saldo):
    return RegistroFuncionario(
        nome=nome,
        final_matricula="001",
        departamento="Teste",
        banco_total_minutos=0,
        banco_saldo_minutos=saldo,
        faltas="",
    )


def test_ranking_exclui_limites_exatos_de_oito_horas():
    wb = Workbook()
    criar_aba_ranking(
        wb,
        [
            _registro("Menos 8 exato", -480),
            _registro("Abaixo de menos 8", -481),
            _registro("Mais 8 exato", 480),
            _registro("Acima de mais 8", 481),
        ],
    )

    nomes = [cell.value for cell in wb["RANKING"]["A"]]
    assert "Abaixo de menos 8" in nomes
    assert "Acima de mais 8" in nomes
    assert "Menos 8 exato" not in nomes
    assert "Mais 8 exato" not in nomes
