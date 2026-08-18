import json

from app.core.config import POWER_BI_DATASET_NAME
from app.integrations.powerbi_client import JORNADA_SCHEMA, PowerBiClient


class Resposta:
    def __init__(self, payload=None):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self):
        if self.payload is None:
            return b""
        return json.dumps(self.payload).encode()


class AplicativoMsalFalso:
    def get_accounts(self):
        return []

    def acquire_token_interactive(self, **kwargs):
        self.kwargs = kwargs
        return {"access_token": "token-de-teste"}


def test_login_publico_usa_tenant_e_redirect_localhost():
    criado = {}
    aplicativo = AplicativoMsalFalso()

    def fabrica(client_id, authority):
        criado.update(client_id=client_id, authority=authority)
        return aplicativo

    cliente = PowerBiClient(
        client_id="app-id", tenant_id="tenant-id", app_factory=fabrica
    )
    cliente.login_interativo()

    assert cliente.autenticado
    assert criado["authority"].endswith("/tenant-id")
    assert "redirect_uri" not in aplicativo.kwargs
    assert aplicativo.kwargs["prompt"] == "select_account"


def test_cria_modelo_push_sem_coluna_cpf_e_envia_linhas():
    requisicoes = []
    respostas = iter(
        [
            Resposta({"value": []}),
            Resposta({"id": "dataset-1"}),
            Resposta(),
        ]
    )

    def opener(request, timeout):
        requisicoes.append((request, timeout))
        return next(respostas)

    cliente = PowerBiClient(opener=opener)
    cliente._access_token = "token"
    dataset_id = cliente.obter_ou_criar_dataset()
    enviados = cliente.enviar_linhas(
        dataset_id,
        [{"IDRelatorio": "r1", "Funcionario": "Ana"}],
    )

    assert enviados == 1
    payload_criacao = json.loads(requisicoes[1][0].data)
    nomes = [item["name"] for item in payload_criacao["tables"][0]["columns"]]
    assert nomes == [nome for nome, _tipo in JORNADA_SCHEMA]
    assert all("cpf" not in nome.casefold() for nome in nomes)
    payload_linhas = json.loads(requisicoes[2][0].data)
    assert payload_linhas == {"rows": [{"IDRelatorio": "r1", "Funcionario": "Ana"}]}


def test_reutiliza_modelo_push_existente():
    requisicoes = []

    respostas = iter(
        [
            Resposta(
                {
                    "value": [
                        {
                            "id": "existente",
                            "name": POWER_BI_DATASET_NAME,
                            "addRowsAPIEnabled": True,
                        }
                    ]
                }
            ),
            Resposta(
                {
                    "value": [
                        {
                            "name": "Jornada",
                            "columns": [
                                {"name": nome, "dataType": tipo}
                                for nome, tipo in JORNADA_SCHEMA
                            ],
                        }
                    ]
                }
            ),
        ]
    )

    def opener(request, timeout):
        requisicoes.append(request)
        return next(respostas)

    cliente = PowerBiClient(opener=opener)
    cliente._access_token = "token"

    assert cliente.obter_ou_criar_dataset() == "existente"
    assert len(requisicoes) == 2


def test_reutiliza_modelo_quando_api_retorna_apenas_nome_da_tabela():
    requisicoes = []
    respostas = iter(
        [
            Resposta(
                {
                    "value": [
                        {
                            "id": "existente",
                            "name": POWER_BI_DATASET_NAME,
                            "addRowsAPIEnabled": True,
                        }
                    ]
                }
            ),
            Resposta({"value": [{"name": "Jornada"}]}),
        ]
    )

    def opener(request, timeout):
        requisicoes.append(request)
        return next(respostas)

    cliente = PowerBiClient(opener=opener)
    cliente._access_token = "token"

    assert cliente.obter_ou_criar_dataset() == "existente"
    assert len(requisicoes) == 2


def test_diagnostico_confirma_workspace_e_modelo_sem_criar_recursos():
    requisicoes = []
    respostas = iter(
        [
            Resposta({"value": []}),
            Resposta(
                {
                    "value": [
                        {
                            "id": "modelo-1",
                            "name": POWER_BI_DATASET_NAME,
                            "addRowsAPIEnabled": True,
                        }
                    ]
                }
            ),
            Resposta({"value": [{"name": "Jornada"}]}),
        ]
    )

    def opener(request, timeout):
        requisicoes.append(request)
        return next(respostas)

    cliente = PowerBiClient(opener=opener)
    cliente._access_token = "token"

    workspace = cliente.verificar_workspace()
    modelo = cliente.verificar_modelo()

    assert workspace.name == "FAS Jornada Analytics"
    assert modelo is not None and modelo.id == "modelo-1"
    assert [request.method for request in requisicoes] == ["GET", "GET", "GET"]
    assert all(request.data is None for request in requisicoes)


def test_diagnostico_informa_modelo_ausente_sem_cria_lo():
    requisicoes = []

    def opener(request, timeout):
        requisicoes.append(request)
        return Resposta({"value": []})

    cliente = PowerBiClient(opener=opener)
    cliente._access_token = "token"

    assert cliente.verificar_modelo() is None
    assert len(requisicoes) == 1
    assert requisicoes[0].method == "GET"
