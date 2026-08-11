from dataclasses import FrozenInstanceError

import pytest

from app.ui.view_state import (
    ConfiguracaoInterface,
    EstadoInterface,
    obter_configuracao_interface,
)


def test_estado_vazio_convida_selecao_e_bloqueia_processamento():
    config = obter_configuracao_interface(EstadoInterface.VAZIO)

    assert config.texto_selecionar == "Selecionar arquivo(s)"
    assert config.processar_habilitado is False
    assert config.limpar_habilitado is False
    assert config.selecionar_habilitado is True
    assert config.configuracao_habilitada is True


@pytest.mark.parametrize(
    ("total", "texto_selecionar", "texto_processar"),
    [
        (1, "1 arquivo selecionado ✓", "Processar arquivo"),
        (3, "3 arquivos selecionados ✓", "Processar arquivos"),
    ],
)
def test_estado_pronto_trata_singular_e_plural(
    total,
    texto_selecionar,
    texto_processar,
):
    config = obter_configuracao_interface(EstadoInterface.PRONTO, total)

    assert config.texto_selecionar == texto_selecionar
    assert config.texto_processar == texto_processar
    assert config.selecionar_habilitado is True
    assert config.limpar_habilitado is True
    assert config.processar_habilitado is True
    assert config.configuracao_habilitada is True


def test_estado_processando_bloqueia_todas_as_acoes():
    config = obter_configuracao_interface(EstadoInterface.PROCESSANDO, 2)

    assert config.texto_selecionar == "Processando..."
    assert config.texto_processar == "Processando..."
    assert config.selecionar_habilitado is False
    assert config.limpar_habilitado is False
    assert config.processar_habilitado is False
    assert config.cancelar_habilitado is True
    assert config.configuracao_habilitada is False


def test_estado_cancelando_bloqueia_todas_as_acoes():
    config = obter_configuracao_interface(EstadoInterface.CANCELANDO, 2)

    assert config.texto_cancelar == "Cancelando..."
    assert config.selecionar_habilitado is False
    assert config.limpar_habilitado is False
    assert config.processar_habilitado is False
    assert config.cancelar_habilitado is False


@pytest.mark.parametrize("total", [1, 2])
def test_estado_cancelado_permite_nova_tentativa(total):
    config = obter_configuracao_interface(EstadoInterface.CANCELADO, total)

    assert config.texto_processar == "Processar novamente"
    assert config.selecionar_habilitado is True
    assert config.limpar_habilitado is True
    assert config.processar_habilitado is True
    assert config.cancelar_habilitado is False
    assert config.configuracao_habilitada is True


@pytest.mark.parametrize(
    ("total", "texto_selecionar", "texto_processar"),
    [
        (1, "Selecionar novo arquivo", "Salvo ✓"),
        (2, "Selecionar novos arquivos", "Arquivos salvos ✓"),
    ],
)
def test_estado_concluido_trata_singular_e_plural(
    total,
    texto_selecionar,
    texto_processar,
):
    config = obter_configuracao_interface(EstadoInterface.CONCLUIDO, total)

    assert config.texto_selecionar == texto_selecionar
    assert config.texto_processar == texto_processar
    assert config.selecionar_habilitado is True
    assert config.processar_habilitado is False
    assert config.configuracao_habilitada is False


def test_estado_erro_permite_corrigir_e_tentar_novamente():
    config = obter_configuracao_interface(EstadoInterface.ERRO, 1)

    assert config.texto_processar == "Tentar novamente"
    assert config.selecionar_habilitado is True
    assert config.limpar_habilitado is True
    assert config.processar_habilitado is True
    assert config.configuracao_habilitada is True


def test_configuracao_e_imutavel():
    config = obter_configuracao_interface(EstadoInterface.VAZIO)

    assert isinstance(config, ConfiguracaoInterface)
    with pytest.raises(FrozenInstanceError):
        config.texto_selecionar = "Alterado"


def test_total_negativo_e_rejeitado():
    with pytest.raises(ValueError, match="nao pode ser negativo"):
        obter_configuracao_interface(EstadoInterface.PRONTO, -1)
