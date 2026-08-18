from types import SimpleNamespace

import pytest

from app.ui.report_pages import (
    CsvPage,
    ProcessingPage,
    SuccessPage,
    _normalizar_progresso,
)


class WidgetFalso:
    def __init__(self, valor=None):
        self.valor = valor
        self.configuracao = {}

    def configure(self, **kwargs):
        self.configuracao.update(kwargs)

    def get(self):
        return self.valor

    def set(self, valor):
        self.valor = valor


@pytest.mark.parametrize(
    ("entrada", "esperado"),
    [
        (0.25, 0.25),
        (25, 0.25),
        (150, 1.0),
        (-1, 0.0),
        ("inválido", 0.0),
    ],
)
def test_normalizar_progresso_aceita_fracao_e_percentual(entrada, esperado):
    assert _normalizar_progresso(entrada) == pytest.approx(esperado)


def test_csv_processar_entrega_departamento_e_opcoes_ao_callback():
    chamadas = []
    pagina = SimpleNamespace(
        _ao_processar=lambda *args, **kwargs: chamadas.append((args, kwargs)),
        combo_departamento=WidgetFalso("Operações"),
        var_saldo=WidgetFalso(True),
        var_resumo=WidgetFalso(False),
        var_ranking=WidgetFalso(True),
    )

    CsvPage._processar(pagina)

    assert chamadas == [
        (
            ("Operações",),
            {
                "gerar_saldo": True,
                "gerar_resumo": False,
                "gerar_ranking": True,
            },
        )
    ]


def test_csv_departamentos_incluem_todos_e_preservam_selecao_valida():
    combo = WidgetFalso()
    pagina = SimpleNamespace(combo_departamento=combo)

    CsvPage.atualizar_departamentos(
        pagina,
        ["Financeiro", "Operações"],
        selecionado="Operações",
    )

    assert combo.configuracao["values"] == ["Todos", "Financeiro", "Operações"]
    assert combo.valor == "Operações"


def test_csv_definir_acoes_atualiza_controles_sem_tk():
    checks = [WidgetFalso(), WidgetFalso(), WidgetFalso()]
    pagina = SimpleNamespace(
        btn_selecionar=WidgetFalso(),
        btn_limpar=WidgetFalso(),
        btn_processar=WidgetFalso(),
        combo_departamento=WidgetFalso(),
        check_saldo=checks[0],
        check_resumo=checks[1],
        check_ranking=checks[2],
    )

    CsvPage.definir_acoes(
        pagina,
        selecionar_habilitado=False,
        limpar_habilitado=False,
        processar_habilitado=False,
        configuracao_habilitada=False,
        texto_selecionar="Processando...",
        texto_processar="Processando...",
    )

    assert pagina.btn_selecionar.configuracao == {
        "state": "disabled",
        "text": "Processando...",
    }
    assert pagina.btn_processar.configuracao == {
        "state": "disabled",
        "text": "Processando...",
    }
    assert pagina.combo_departamento.configuracao["state"] == "disabled"
    assert all(item.configuracao["state"] == "disabled" for item in checks)


def test_processing_page_atualiza_percentual_sem_janela_real():
    mensagens = []
    pagina = SimpleNamespace(
        progress=WidgetFalso(),
        label_percentual=WidgetFalso(),
        atualizar_status=mensagens.append,
    )

    ProcessingPage.atualizar_progresso(pagina, 37, "Consultando o RHiD...")

    assert pagina.progress.valor == pytest.approx(0.37)
    assert pagina.label_percentual.configuracao["text"] == "37%"
    assert mensagens == ["Consultando o RHiD..."]


def test_success_page_atualiza_metricas_e_caminho_sem_janela_real():
    pagina = SimpleNamespace(
        label_metric_funcionarios=WidgetFalso(),
        label_metric_banco_total=WidgetFalso(),
        label_metric_banco_saldo=WidgetFalso(),
        label_caminho=WidgetFalso(),
    )

    SuccessPage.atualizar_metricas(pagina, 47, "12:30", "-03:15")
    SuccessPage.atualizar_caminho(pagina, r"C:\Relatorios\jornada.xlsx")

    assert pagina.label_metric_funcionarios.configuracao["text"] == "47"
    assert pagina.label_metric_banco_total.configuracao["text"] == "12:30"
    assert pagina.label_metric_banco_saldo.configuracao["text"] == "-03:15"
    assert pagina.label_caminho.configuracao["text"].endswith("jornada.xlsx")
