"""Estado puro da navegação entre as páginas do FAS Jornada."""

from __future__ import annotations

from enum import Enum


class PaginaInterface(str, Enum):
    INICIO = "inicio"
    DIAGNOSTICOS = "diagnosticos"
    CSV = "csv"
    RHID_LOGIN = "rhid_login"
    RHID_DOMINIO = "rhid_dominio"
    RHID_ESCOPO = "rhid_escopo"
    PROCESSAMENTO = "processamento"
    SUCESSO = "sucesso"


_VOLTAR_PARA = {
    PaginaInterface.DIAGNOSTICOS: PaginaInterface.INICIO,
    PaginaInterface.CSV: PaginaInterface.INICIO,
    PaginaInterface.RHID_LOGIN: PaginaInterface.INICIO,
    PaginaInterface.RHID_DOMINIO: PaginaInterface.RHID_LOGIN,
    PaginaInterface.RHID_ESCOPO: PaginaInterface.RHID_LOGIN,
}


def pagina_anterior(pagina: PaginaInterface) -> PaginaInterface | None:
    return _VOLTAR_PARA.get(PaginaInterface(pagina))


__all__ = ["PaginaInterface", "pagina_anterior"]
