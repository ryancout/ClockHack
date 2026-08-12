from app.ui.navigation import PaginaInterface, pagina_anterior


def test_paginas_de_entrada_voltam_ao_inicio():
    assert pagina_anterior(PaginaInterface.CSV) is PaginaInterface.INICIO
    assert pagina_anterior(PaginaInterface.RHID_LOGIN) is PaginaInterface.INICIO


def test_dominio_e_escopo_voltam_ao_login():
    assert pagina_anterior(PaginaInterface.RHID_DOMINIO) is PaginaInterface.RHID_LOGIN
    assert pagina_anterior(PaginaInterface.RHID_ESCOPO) is PaginaInterface.RHID_LOGIN


def test_inicio_nao_possui_pagina_anterior():
    assert pagina_anterior(PaginaInterface.INICIO) is None
