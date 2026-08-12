from datetime import date
from types import SimpleNamespace

import pytest

from app.ui.rhid_page import (
    RHID_FORGOT_PASSWORD_URL,
    RhidPage,
    catalogar_por_id,
    departamentos_da_empresa,
    periodo_padrao,
    validar_periodo,
)


def test_recuperacao_de_senha_usa_a_rota_oficial_do_rhid():
    assert RHID_FORGOT_PASSWORD_URL == "https://www.rhid.com.br/v2/#/forgot_password"


class ValorFalso:
    def __init__(self, valor):
        self.valor = valor

    def get(self):
        return self.valor


def test_periodo_padrao_e_validacao_usam_formato_brasileiro_sem_limite_artificial():
    assert periodo_padrao(date(2026, 8, 12)) == ("01/08/2026", "12/08/2026")
    assert validar_periodo("01/01/2025", "12/08/2026") == (
        "2025-01-01",
        "2026-08-12",
    )


@pytest.mark.parametrize(
    ("inicio", "fim", "mensagem"),
    [
        ("2026-08-01", "12/08/2026", "DD/MM/AAAA"),
        ("13/08/2026", "12/08/2026", "anterior"),
        ("31/02/2026", "12/08/2026", "DD/MM/AAAA"),
    ],
)
def test_validar_periodo_rejeita_datas_invalidas(inicio, fim, mensagem):
    with pytest.raises(ValueError, match=mensagem):
        validar_periodo(inicio, fim)


def test_catalogo_preserva_homonimos_por_id_e_remove_so_id_repetido():
    itens = (
        SimpleNamespace(id=10, name="Operação"),
        SimpleNamespace(id=11, name="Operação"),
        SimpleNamespace(id="10", name="Duplicado"),
    )

    rotulos, por_id = catalogar_por_id(itens, lambda item: item.name)

    assert rotulos == {
        "Operação — ID 10": 10,
        "Operação — ID 11": 11,
    }
    assert tuple(por_id) == ("10", "11")


def test_departamentos_sao_filtrados_por_empresa_sem_unir_homonimos():
    departamentos = (
        SimpleNamespace(id=21, name="Campo", company_id="10"),
        SimpleNamespace(id=22, name="Campo", company_id=10),
        SimpleNamespace(id=31, name="Campo", company_id=11),
    )

    filtrados = departamentos_da_empresa(departamentos, 10)

    assert tuple(item.id for item in filtrados) == (21, 22)


@pytest.mark.parametrize("empresa_global", [None, "", 0, "0"])
def test_departamento_global_do_rhid_aparece_na_empresa_selecionada(empresa_global):
    departamentos = (
        SimpleNamespace(id=20, name="Global", company_id=empresa_global),
        SimpleNamespace(id=21, name="Projeto", company_id=10),
        SimpleNamespace(id=31, name="Outro projeto", company_id=11),
    )

    filtrados = departamentos_da_empresa(departamentos, 10)

    assert tuple(item.id for item in filtrados) == (20, 21)


def _pagina_parametros(*, todos, selecionados):
    variaveis = {
        "21": ValorFalso("21" in selecionados),
        "22": ValorFalso("22" in selecionados),
    }
    return SimpleNamespace(
        combo_empresa=ValorFalso("Projeto — ID 10"),
        _empresa_id_por_rotulo={"Projeto — ID 10": 10},
        _departamentos_visiveis=(object(), object()),
        var_todos_setores=ValorFalso(todos),
        _departamento_variaveis=variaveis,
        _departamento_id_por_chave={"21": 21, "22": 22},
        entry_data_inicial=ValorFalso("01/08/2026"),
        entry_data_final=ValorFalso("12/08/2026"),
        var_gerar_saldo=ValorFalso(True),
        var_gerar_resumo=ValorFalso(True),
        var_gerar_ranking=ValorFalso(True),
    )


def test_todos_os_setores_e_representado_por_none_no_callback():
    pagina = _pagina_parametros(todos=True, selecionados=())

    parametros = RhidPage.obter_parametros_geracao(pagina)

    assert parametros == (
        10,
        None,
        "2026-08-01",
        "2026-08-12",
        True,
        True,
        True,
    )


def test_selecao_multipla_preserva_os_ids_dos_setores():
    pagina = _pagina_parametros(todos=False, selecionados=("21", "22"))

    parametros = RhidPage.obter_parametros_geracao(pagina)

    assert parametros[1] == (21, 22)


def test_selecao_avulsa_exige_ao_menos_um_setor():
    pagina = _pagina_parametros(todos=False, selecionados=())

    with pytest.raises(ValueError, match="ao menos um setor"):
        RhidPage.obter_parametros_geracao(pagina)


def test_callback_de_geracao_recebe_contrato_completo():
    chamadas = []
    parametros = (10, (21, 22), "2026-08-01", "2026-08-12", True, False, True)
    pagina = SimpleNamespace(
        _geracao_em_andamento=False,
        _ao_gerar=lambda *args: chamadas.append(args),
        obter_parametros_geracao=lambda: parametros,
        exibir_erro=lambda mensagem: pytest.fail(mensagem),
    )

    RhidPage._gerar_relatorio(pagina)

    assert chamadas == [parametros]


def test_um_unico_dominio_nao_abre_etapa_intermediaria():
    chamadas = []
    pagina = SimpleNamespace(
        definir_ocupado=lambda _ocupado: None,
        _dominio_por_rotulo={"antigo": "antigo"},
        _dominio_preenchido="",
        obter_credenciais_digitadas=lambda: (
            "usuario@empresa.com",
            "segredo",
            pagina._dominio_preenchido,
            False,
        ),
        _ao_conectar=lambda *args: chamadas.append(args),
        _mostrar_etapa=lambda _etapa: pytest.fail("Não deveria mostrar outra etapa"),
    )
    tenant = SimpleNamespace(domain="cliente", name="Cliente")

    RhidPage.exibir_dominios(pagina, (tenant,))

    assert chamadas == [("usuario@empresa.com", "segredo", "cliente")]
    assert pagina._dominio_por_rotulo == {}
