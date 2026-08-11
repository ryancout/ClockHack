"""Estado visual puro da tela principal.

Este modulo nao conhece widgets. Ele apenas traduz o estado do fluxo em
textos e permissoes que a camada CustomTkinter pode aplicar.
"""

from dataclasses import dataclass
from enum import Enum


class EstadoInterface(str, Enum):
    """Etapas possiveis do fluxo de selecao e processamento."""

    VAZIO = "vazio"
    PRONTO = "pronto"
    PROCESSANDO = "processando"
    CANCELANDO = "cancelando"
    CONCLUIDO = "concluido"
    CANCELADO = "cancelado"
    ERRO = "erro"


@dataclass(frozen=True, slots=True)
class ConfiguracaoInterface:
    """Textos e habilitacoes a serem refletidos pelos widgets da tela."""

    texto_selecionar: str
    texto_processar: str
    texto_cancelar: str
    selecionar_habilitado: bool
    limpar_habilitado: bool
    processar_habilitado: bool
    cancelar_habilitado: bool
    configuracao_habilitada: bool


def obter_configuracao_interface(
    estado: EstadoInterface,
    total_arquivos: int = 0,
) -> ConfiguracaoInterface:
    """Monta a configuracao visual sem consultar ou alterar a interface."""
    if total_arquivos < 0:
        raise ValueError("O total de arquivos nao pode ser negativo.")

    if estado is EstadoInterface.VAZIO:
        return ConfiguracaoInterface(
            texto_selecionar="Selecionar arquivo(s)",
            texto_processar="Processar arquivo(s)",
            texto_cancelar="Cancelar",
            selecionar_habilitado=True,
            limpar_habilitado=False,
            processar_habilitado=False,
            cancelar_habilitado=False,
            configuracao_habilitada=True,
        )

    if estado is EstadoInterface.PRONTO:
        singular = total_arquivos == 1
        return ConfiguracaoInterface(
            texto_selecionar=(
                "1 arquivo selecionado ✓"
                if singular
                else f"{total_arquivos} arquivos selecionados ✓"
            ),
            texto_processar=("Processar arquivo" if singular else "Processar arquivos"),
            texto_cancelar="Cancelar",
            selecionar_habilitado=True,
            limpar_habilitado=True,
            processar_habilitado=total_arquivos > 0,
            cancelar_habilitado=False,
            configuracao_habilitada=True,
        )

    if estado is EstadoInterface.PROCESSANDO:
        return ConfiguracaoInterface(
            texto_selecionar="Processando...",
            texto_processar="Processando...",
            texto_cancelar="Cancelar processamento",
            selecionar_habilitado=False,
            limpar_habilitado=False,
            processar_habilitado=False,
            cancelar_habilitado=True,
            configuracao_habilitada=False,
        )

    if estado is EstadoInterface.CANCELANDO:
        return ConfiguracaoInterface(
            texto_selecionar="Cancelando...",
            texto_processar="Cancelando...",
            texto_cancelar="Cancelando...",
            selecionar_habilitado=False,
            limpar_habilitado=False,
            processar_habilitado=False,
            cancelar_habilitado=False,
            configuracao_habilitada=False,
        )

    if estado is EstadoInterface.CONCLUIDO:
        plural = total_arquivos != 1
        return ConfiguracaoInterface(
            texto_selecionar=(
                "Selecionar novos arquivos" if plural else "Selecionar novo arquivo"
            ),
            texto_processar="Arquivos salvos ✓" if plural else "Salvo ✓",
            texto_cancelar="Cancelar",
            selecionar_habilitado=True,
            limpar_habilitado=True,
            processar_habilitado=False,
            cancelar_habilitado=False,
            configuracao_habilitada=False,
        )

    if estado is EstadoInterface.CANCELADO:
        plural = total_arquivos != 1
        return ConfiguracaoInterface(
            texto_selecionar=(
                "Selecionar novos arquivos" if plural else "Selecionar novo arquivo"
            ),
            texto_processar="Processar novamente",
            texto_cancelar="Cancelar",
            selecionar_habilitado=True,
            limpar_habilitado=True,
            processar_habilitado=total_arquivos > 0,
            cancelar_habilitado=False,
            configuracao_habilitada=True,
        )

    if estado is EstadoInterface.ERRO:
        return ConfiguracaoInterface(
            texto_selecionar="Selecionar arquivo(s)",
            texto_processar="Tentar novamente",
            texto_cancelar="Cancelar",
            selecionar_habilitado=True,
            limpar_habilitado=True,
            processar_habilitado=total_arquivos > 0,
            cancelar_habilitado=False,
            configuracao_habilitada=True,
        )

    raise ValueError(f"Estado de interface desconhecido: {estado!r}")


__all__ = [
    "ConfiguracaoInterface",
    "EstadoInterface",
    "obter_configuracao_interface",
]
