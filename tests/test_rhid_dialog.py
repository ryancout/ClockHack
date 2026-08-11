from datetime import date
from types import SimpleNamespace

import pytest

from app.ui.rhid_dialog import (
    TODAS_AS_EMPRESAS,
    TODOS_OS_SETORES,
    RhidConnectionDialog,
    periodo_padrao,
    validar_periodo,
)


class WidgetFalso:
    def __init__(self, valor=""):
        self.valor = valor
        self.configuracao = {}

    def configure(self, **kwargs):
        self.configuracao.update(kwargs)

    def set(self, valor):
        self.valor = valor

    def get(self):
        return self.valor

    def grid(self):
        self.configuracao["visivel"] = True

    def grid_remove(self):
        self.configuracao["visivel"] = False


def test_periodo_padrao_usa_inicio_do_mes_ate_hoje():
    assert periodo_padrao(date(2026, 8, 11)) == ("2026-08-01", "2026-08-11")


@pytest.mark.parametrize(
    ("inicio", "fim", "mensagem"),
    [
        ("11/08/2026", "2026-08-11", "AAAA-MM-DD"),
        ("2026-08-12", "2026-08-11", "anterior"),
        ("2026-01-01", "2026-08-11", "31 dias"),
    ],
)
def test_validar_periodo_rejeita_valores_invalidos(inicio, fim, mensagem):
    with pytest.raises(ValueError, match=mensagem):
        validar_periodo(inicio, fim)


def test_validar_periodo_normaliza_datas_iso():
    assert validar_periodo(" 2026-08-01 ", "2026-08-11") == (
        "2026-08-01",
        "2026-08-11",
    )


def test_setores_com_mesmo_nome_sao_mapeados_pelo_id():
    combo = WidgetFalso()
    dialogo = SimpleNamespace(
        _empresa_id_por_rotulo={"Empresa — ID 10": 10},
        _departamentos=(
            SimpleNamespace(id=21, name="Operação", company_id="10"),
            SimpleNamespace(id=22, name="Operação", company_id=10),
            SimpleNamespace(id=31, name="Operação", company_id=11),
        ),
        combo_departamento=combo,
    )

    RhidConnectionDialog._empresa_alterada(dialogo, "Empresa — ID 10")

    assert dialogo._departamento_id_por_rotulo == {
        TODOS_OS_SETORES: None,
        "Operação — ID 21": 21,
        "Operação — ID 22": 22,
    }
    assert combo.valor == TODOS_OS_SETORES


def test_empresas_com_mesmo_nome_sao_mapeadas_pelo_id():
    dialogo = SimpleNamespace(
        definir_ocupado=lambda _ocupado: None,
        credenciais=WidgetFalso(),
        escopo=WidgetFalso(),
        combo_empresa=WidgetFalso(),
        combo_departamento=WidgetFalso(),
        label_status=WidgetFalso(),
        _atualizar_estados_controles=lambda: None,
    )
    dialogo._empresa_alterada = lambda rotulo: RhidConnectionDialog._empresa_alterada(
        dialogo, rotulo
    )
    empresas = (
        SimpleNamespace(id=10, label="Unidade Centro"),
        SimpleNamespace(id=11, label="Unidade Centro"),
        SimpleNamespace(id="10", label="Cadastro repetido"),
    )

    departamentos = (SimpleNamespace(id=21, name="Operação", company_id=10),)

    RhidConnectionDialog.exibir_catalogo(dialogo, empresas, departamentos)

    assert dialogo._empresa_id_por_rotulo == {
        TODAS_AS_EMPRESAS: None,
        "Unidade Centro — ID 10": 10,
        "Unidade Centro — ID 11": 11,
    }
    assert dialogo._catalogo_carregado
    assert dialogo.combo_empresa.valor == TODAS_AS_EMPRESAS
    assert dialogo._departamento_id_por_rotulo == {
        TODOS_OS_SETORES: None,
        "Operação — ID 21": 21,
    }


def test_botao_gerar_so_habilita_com_catalogo_e_callback():
    widgets = [WidgetFalso() for _ in range(8)]
    dialogo = SimpleNamespace(
        _conexao_em_andamento=False,
        _geracao_em_andamento=False,
        _catalogo_carregado=False,
        _ao_gerar=lambda *_args: None,
        _dominio_somente_leitura=False,
        entry_email=widgets[0],
        entry_senha=widgets[1],
        combo_dominio=widgets[2],
        btn_conectar=widgets[3],
        combo_empresa=widgets[4],
        combo_departamento=widgets[5],
        entry_data_inicial=widgets[6],
        entry_data_final=widgets[7],
        btn_gerar=WidgetFalso(),
        btn_fechar=WidgetFalso(),
    )

    RhidConnectionDialog._atualizar_estados_controles(dialogo)
    assert dialogo.btn_gerar.configuracao["state"] == "disabled"

    dialogo._catalogo_carregado = True
    RhidConnectionDialog._atualizar_estados_controles(dialogo)
    assert dialogo.btn_gerar.configuracao["state"] == "normal"


def test_progresso_aceita_percentual_do_controller():
    dialogo = SimpleNamespace(
        progress_geracao=WidgetFalso(),
        label_status=WidgetFalso(),
    )

    RhidConnectionDialog.atualizar_progresso(dialogo, 35, "Consultando pessoas...")

    assert dialogo.progress_geracao.valor == pytest.approx(0.35)
    assert dialogo.label_status.configuracao["text"] == "Consultando pessoas..."


def test_callback_de_geracao_recebe_ids_e_datas_sem_antecipar_estado_ocupado():
    chamadas = []
    dialogo = SimpleNamespace(
        _ao_gerar=lambda *parametros: chamadas.append(parametros),
        _geracao_em_andamento=False,
        obter_parametros_geracao=lambda: (10, 22, "2026-08-01", "2026-08-11"),
        exibir_erro_geracao=lambda mensagem: pytest.fail(mensagem),
    )

    RhidConnectionDialog._gerar_excel(dialogo)

    assert chamadas == [(10, 22, "2026-08-01", "2026-08-11")]
