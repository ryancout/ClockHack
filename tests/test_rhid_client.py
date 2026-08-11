import json
import base64
from io import BytesIO
from datetime import date
from urllib.error import HTTPError

import pytest

from app.integrations.rhid_client import (
    RhidApiError,
    RhidClient,
    RhidDepartment,
    RhidTenantRequired,
)


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self):
        if isinstance(self.payload, bytes):
            return self.payload
        return json.dumps(self.payload, ensure_ascii=False).encode("utf-8")


class FakeApi:
    def __init__(self):
        self.requests = []

    def __call__(self, request, timeout):
        self.requests.append((request, timeout))
        path = request.full_url.split("?", 1)[0]
        if path.endswith("/api/auth/login"):
            return FakeResponse({"accessToken": "token-de-teste"})
        if path.endswith("/v2/customerdb/company.svc/a"):
            return FakeResponse(
                {
                    "data": [
                        {"id": 1, "name": "Projeto A", "tradingName": "A"},
                        {"id": 2, "name": "Projeto B", "tradingName": ""},
                    ]
                }
            )
        if path.endswith("/v2/customerdb/department.svc/a"):
            return FakeResponse(
                {
                    "data": [
                        {"id": 10, "name": "Operações", "idCompany": 1},
                        {"id": 20, "name": "Financeiro", "idCompany": 2},
                    ]
                }
            )
        if path.endswith("/v2/customerdb/person.svc/a_resumido"):
            return FakeResponse(
                {
                    "data": [
                        {
                            "id": 100,
                            "name": "Pessoa",
                            "registration": "0001",
                            "idCompany": 1,
                            "idDepartment": 10,
                            "status": 1,
                        }
                    ]
                }
            )
        if path.endswith("/v2/api.svc/apuracao_ponto"):
            return FakeResponse(
                json.dumps(
                    [{"idPerson": 100, "totalHorasTrabalhadas": 480}],
                    ensure_ascii=False,
                )
            )
        raise AssertionError(request.full_url)


def client_autenticado():
    api = FakeApi()
    client = RhidClient(opener=api)
    client.login("usuario@empresa.com", "segredo")
    return client, api


def test_login_e_catalogos_usam_token_sem_guardar_senha():
    client, api = client_autenticado()

    empresas = client.listar_empresas()
    departamentos = client.listar_departamentos()
    pessoas = client.listar_pessoas()

    assert [empresa.label for empresa in empresas] == ["A", "Projeto B"]
    assert departamentos[0].company_id == 1
    assert pessoas[0].registration == "0001"
    assert api.requests[1][0].headers["Authorization"] == "Bearer token-de-teste"
    assert not hasattr(client, "password")


def test_apuracao_decodifica_json_serializado_e_valida_periodo():
    client, _api = client_autenticado()

    apuracao = client.obter_apuracao(100, date(2026, 7, 1), "2026-07-31")

    assert apuracao[0]["totalHorasTrabalhadas"] == 480

    with pytest.raises(RhidApiError, match="90 dias"):
        client.obter_apuracao(100, "2026-01-01", "2026-05-01")


def test_cliente_exige_https_e_autenticacao():
    with pytest.raises(ValueError, match="HTTPS"):
        RhidClient("http://rhid.local")

    with pytest.raises(RhidApiError, match="Conecte-se"):
        RhidClient(opener=FakeApi()).listar_empresas()


def test_login_repete_automaticamente_quando_existe_um_unico_dominio():
    chamadas = []

    def api(request, timeout):
        chamadas.append(json.loads(request.data.decode("utf-8")))
        if len(chamadas) == 1:
            return FakeResponse(
                {
                    "Data": [
                        {
                            "systemId": 1,
                            "tenants": [
                                {
                                    "tenantId": 7,
                                    "domain": "cliente-a",
                                    "name": "Cliente A",
                                    "status": True,
                                }
                            ],
                        }
                    ]
                }
            )
        return FakeResponse({"Data": [{"token": "token"}]})

    client = RhidClient(opener=api)
    client.login("usuario@empresa.com", "segredo")

    assert client.autenticado is True
    assert chamadas[1]["domain"] == "cliente-a"
    assert chamadas[1]["tenantId"] == 7
    assert chamadas[1]["systemId"] == 1


def test_erro_http_exibe_mensagem_de_negocio_do_rhid():
    def api(_request, timeout=None):
        raise HTTPError(
            "https://rhid.com.br/v2/api.svc/login",
            500,
            "erro",
            {},
            BytesIO(json.dumps({"error": "Cliente sem acesso ativo."}).encode()),
        )

    with pytest.raises(RhidApiError, match="Cliente sem acesso ativo"):
        RhidClient(opener=api).login("usuario@empresa.com", "segredo")


def test_login_informa_lista_quando_existem_varios_dominios():
    def api(_request, timeout=None):
        return FakeResponse(
            {
                "Data": [
                    {
                        "systemId": 1,
                        "tenants": [
                            {
                                "tenantId": 1,
                                "domain": "PROJETO-A",
                                "name": "Projeto A",
                                "status": True,
                            },
                            {
                                "tenantId": 2,
                                "domain": "PROJETO-B",
                                "name": "Projeto B",
                                "status": True,
                            },
                        ],
                    }
                ]
            }
        )

    with pytest.raises(RhidTenantRequired) as erro:
        RhidClient(opener=api).login("usuario@empresa.com", "segredo")

    assert [tenant.domain for tenant in erro.value.tenants] == ["PROJETO-A", "PROJETO-B"]


def test_login_compativel_nao_envia_ids_nulos():
    chamadas = []

    def api(request, timeout=None):
        payload = json.loads(request.data.decode("utf-8"))
        chamadas.append(payload)
        if len(chamadas) == 1:
            return FakeResponse(
                {"listCustomer": [{"domain": "CLIENTE", "name": "Cliente"}]}
            )
        return FakeResponse({"accessToken": "token"})

    client = RhidClient(opener=api)
    client.login("usuario@empresa.com", "segredo")

    assert chamadas[1]["domain"] == "CLIENTE"
    assert "tenantId" not in chamadas[1]
    assert "systemId" not in chamadas[1]


def test_requisicoes_web_enviam_cliente_extraido_do_token():
    api = FakeApi()
    payload = base64.urlsafe_b64encode(
        json.dumps({"cidCustomerId": 321}).encode("utf-8")
    ).decode("ascii").rstrip("=")
    token = f"cabecalho.{payload}.assinatura"

    def api_com_token(request, timeout=None):
        if request.full_url.endswith("/api/auth/login"):
            return FakeResponse({"accessToken": token})
        return api(request, timeout)

    client = RhidClient(opener=api_com_token)
    client.login("usuario@empresa.com", "segredo")
    client.listar_empresas()

    assert api.requests[0][0].headers["X-cid-rhid"] == "321"


def _catalogo_colunas_clockhack():
    campos = [
        ("Person", "name", "Nome do funcionário"),
        ("Person", "registration", "Número de matrícula"),
        ("Department", "name", "Nome do departamento"),
        ("Person", "cpf", "CPF do funcionário"),
        ("Acjef", "horasTotalNaoExtra", "Total Normais"),
        ("Acjef", "totalHorasTrabalhadas", "Total Trabalhado"),
        ("Acjef", "horasUteis", "Horas Previstas"),
        ("Acjef", "faltasDiasInteiro", "Dia Falta"),
        ("Acjef", "horasFaltaAtraso", "Falta e Atraso"),
        ("Acjef", "minutosAbono", "Abono"),
        ("Acjef", "horasExtrasCalculadas", "Extras Total"),
        ("Acjef", "saldoBancoAjustadoMaisCredDeb", "Banco Total"),
        ("Acjef", "saldoBancoFinalDia", "Banco Saldo"),
    ]
    return [
        {
            "className": classe,
            "propertyName": propriedade,
            "headerText": cabecalho,
            "ordem": ordem,
            "mostra": True,
            "multiplica": False,
            "multiplicaNoturno": False,
        }
        for ordem, (classe, propriedade, cabecalho) in enumerate(campos)
    ]


def test_catalogo_prioriza_cabecalho_oficial_quando_nome_interno_colide():
    catalogo = _catalogo_colunas_clockhack()
    catalogo.extend(
        [
            {
                "className": "Acjef",
                "propertyName": "strHorarioContratualSimples",
                "headerText": "Previsto",
                "ordem": 90,
                "mostra": True,
                "multiplica": False,
                "multiplicaNoturno": False,
            },
            {
                "className": "Acjef",
                "propertyName": "horaExtraPorPercentual",
                "headerText": "Extra por Percentual",
                "ordem": 91,
                "mostra": True,
                "multiplica": False,
                "multiplicaNoturno": False,
            },
        ]
    )

    def api(request, timeout=None):
        if request.full_url.endswith("/api/auth/login"):
            return FakeResponse({"accessToken": "token"})
        if request.full_url.endswith(
            "/maindb/layoutproperty.svc/exportableAcjefAndPersonColumns"
        ):
            return FakeResponse(catalogo)
        raise AssertionError(request.full_url)

    cliente = RhidClient(opener=api)
    cliente.login("usuario@empresa.com", "segredo")

    selecionadas = cliente._propriedades_relatorio()

    assert selecionadas[6]["propertyName"] == "horasUteis"
    assert selecionadas[6]["headerText"] == "Horas Previstas"
    assert selecionadas[10]["propertyName"] == "horasExtrasCalculadas"
    assert selecionadas[10]["headerText"] == "Extras Total"


def test_gera_csv_oficial_com_filtros_colunas_polling_e_download():
    requisicoes = []
    percentuais = iter((25, 100))
    csv_esperado = b"Nome do funcion\xc3\xa1rio;N\xc3\xbamero de matr\xc3\xadcula\r\nPessoa;001\r\n"

    def api(request, timeout=None):
        requisicoes.append(request)
        url = request.full_url
        if url.endswith("/api/auth/login"):
            return FakeResponse({"accessToken": "token"})
        if url.endswith("/maindb/layoutproperty.svc/exportableAcjefAndPersonColumns"):
            return FakeResponse(_catalogo_colunas_clockhack())
        if url.endswith("/maindb/parameter.svc/parameters_global"):
            return FakeResponse([{"id": 118, "valueN": 1}])
        if "/customerdb/person.svc/a_ativo?" in url:
            return FakeResponse(
                {
                    "data": [
                        {
                            "id": 101,
                        },
                    ]
                }
            )
        if url.endswith("/report.svc/ponto"):
            return FakeResponse({"guid": "guid-123", "numPeople": 1})
        if "/customerdb/notify.svc/specificGuid/?guid=guid-123" in url:
            return FakeResponse({"percent": next(percentuais)})
        if "/customerdb/notify.svc/save_file/?format=CSV&guid=guid-123" in url:
            return FakeResponse(csv_esperado)
        raise AssertionError(url)

    progresso = []
    cliente = RhidClient(opener=api, sleeper=lambda _segundos: None)
    cliente.login("usuario@empresa.com", "segredo")
    conteudo = cliente.gerar_relatorio_csv(
        7,
        12,
        "2026-08-01",
        "2026-08-11",
        progresso.append,
    )

    assert conteudo == csv_esperado
    assert progresso == [0, 25, 100]
    requisicao_relatorio = next(
        item for item in requisicoes if item.full_url.endswith("/report.svc/ponto")
    )
    payload = json.loads(requisicao_relatorio.data.decode("utf-8"))
    assert payload["pdfCartaoPontoParameters"]["listIdStr"] == [101]
    assert payload["pdfCartaoPontoParameters"]["listCompanyStr"] == [7]
    assert payload["pdfCartaoPontoParameters"]["listDepartmentStr"] == [12]
    assert payload["listPersonInfo"] == [
        "Person.name",
        "Person.registration",
        "Department.name",
        "Person.cpf",
    ]
    assert payload["listColumns"][-2:] == [
        "saldoBancoAjustadoMaisCredDeb",
        "saldoBancoFinalDia",
    ]
    assert payload["ini"] == "20260801"
    assert payload["fim"] == "20260811"
    assert payload["formatoSaida"] == "CSV"
    assert payload["agrupamento"] == "person"
    assert payload["relatorio"] == "extrato"
    assert payload["status"] == 1
    requisicao_pessoas = next(
        item for item in requisicoes if "/customerdb/person.svc/a_ativo?" in item.full_url
    )
    assert "companies=7" in requisicao_pessoas.full_url
    assert "departments=12" in requisicao_pessoas.full_url


def test_relatorio_exige_as_colunas_do_contrato_e_limita_periodo():
    def api(request, timeout=None):
        if request.full_url.endswith("/api/auth/login"):
            return FakeResponse({"accessToken": "token"})
        if request.full_url.endswith(
            "/maindb/layoutproperty.svc/exportableAcjefAndPersonColumns"
        ):
            return FakeResponse(_catalogo_colunas_clockhack()[:-1])
        raise AssertionError(request.full_url)

    cliente = RhidClient(opener=api, sleeper=lambda _segundos: None)
    cliente.login("usuario@empresa.com", "segredo")

    with pytest.raises(RhidApiError, match="Banco Saldo"):
        cliente.gerar_relatorio_csv(1, None, "2026-08-01", "2026-08-10")
    with pytest.raises(RhidApiError, match="31 dias"):
        cliente.gerar_relatorio_csv(1, None, "2026-01-01", "2026-02-15")


def test_relatorio_interrompe_quando_rhid_informa_zero_pessoas():
    def api(request, timeout=None):
        url = request.full_url
        if url.endswith("/api/auth/login"):
            return FakeResponse({"accessToken": "token"})
        if url.endswith("/maindb/layoutproperty.svc/exportableAcjefAndPersonColumns"):
            return FakeResponse(_catalogo_colunas_clockhack())
        if url.endswith("/maindb/parameter.svc/parameters_global"):
            return FakeResponse([{"id": 118, "valueN": 1}])
        if "/customerdb/person.svc/a_ativo?" in url:
            return FakeResponse({"data": [{"id": 101}]})
        if url.endswith("/report.svc/ponto"):
            return FakeResponse({"guid": "", "numPeople": 0})
        raise AssertionError(url)

    cliente = RhidClient(opener=api, sleeper=lambda _segundos: None)
    cliente.login("usuario@empresa.com", "segredo")

    with pytest.raises(RhidApiError, match="não encontrou funcionários"):
        cliente.gerar_relatorio_csv(7, 12, "2026-08-01", "2026-08-11")


def test_ids_de_pessoas_sao_filtrados_no_servidor_e_deduplicados():
    requisicoes = []

    def api(request, timeout=None):
        requisicoes.append(request.full_url)
        if request.full_url.endswith("/api/auth/login"):
            return FakeResponse({"accessToken": "token"})
        if "/customerdb/person.svc/a_ativo?" in request.full_url:
            return FakeResponse({"data": [{"id": 10}, {"id": 10}, {"id": 11}]})
        raise AssertionError(request.full_url)

    cliente = RhidClient(opener=api)
    cliente.login("usuario@empresa.com", "segredo")

    assert cliente.listar_ids_pessoas_ativas(7) == (10, 11)
    assert "companies=7" in requisicoes[-1]
    assert "departments=" not in requisicoes[-1]


def test_catalogo_mantem_todos_os_setores_com_funcionarios_ativos():
    ids_setores_ativos = list(range(100, 118))
    pessoas = [
        {"id": indice + 1, "idDepartment": ids_setores_ativos[indice % 18]}
        for indice in range(47)
    ]

    def api(request, timeout=None):
        if request.full_url.endswith("/api/auth/login"):
            return FakeResponse({"accessToken": "token"})
        if request.full_url.endswith("/customerdb/person.svc/a_ativo"):
            return FakeResponse({"data": pessoas})
        raise AssertionError(request.full_url)

    cliente = RhidClient(opener=api)
    cliente.login("usuario@empresa.com", "segredo")
    departamentos = tuple(
        RhidDepartment(id_setor, f"Setor {id_setor}", 1)
        for id_setor in [*ids_setores_ativos, 999]
    )

    encontrados = cliente.filtrar_departamentos_com_pessoas_ativas(departamentos)

    assert len(encontrados) == 18
    assert {item.id for item in encontrados} == set(ids_setores_ativos)
    assert 999 not in {item.id for item in encontrados}
