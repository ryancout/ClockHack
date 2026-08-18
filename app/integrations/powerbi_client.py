"""Cliente mínimo para envio de snapshots ao Power BI REST API."""

from __future__ import annotations

import json
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from app.core.config import (
    POWER_BI_CLIENT_ID,
    POWER_BI_DATASET_NAME,
    POWER_BI_TABLE_NAME,
    POWER_BI_TENANT_ID,
    POWER_BI_WORKSPACE_ID,
    POWER_BI_WORKSPACE_NAME,
)


POWER_BI_SCOPE = "https://analysis.windows.net/powerbi/api/Dataset.ReadWrite.All"
POWER_BI_API_URL = "https://api.powerbi.com/v1.0/myorg"


class PowerBiApiError(Exception):
    """Erro seguro e apresentável ocorrido na integração Power BI."""


@dataclass(frozen=True, slots=True)
class PowerBiWorkspaceInfo:
    id: str
    name: str


@dataclass(frozen=True, slots=True)
class PowerBiModelInfo:
    id: str
    name: str
    accepts_rows: bool


JORNADA_SCHEMA: tuple[tuple[str, str], ...] = (
    ("IDRelatorio", "string"),
    ("GeradoEm", "DateTime"),
    ("PeriodoInicial", "DateTime"),
    ("PeriodoFinal", "DateTime"),
    ("Competencia", "string"),
    ("Ano", "Int64"),
    ("Mes", "Int64"),
    ("Origem", "string"),
    ("Empresa", "string"),
    ("SetoresSelecionados", "string"),
    ("Arquivo", "string"),
    ("Matricula", "string"),
    ("Funcionario", "string"),
    ("Departamento", "string"),
    ("TotalNormaisMinutos", "Int64"),
    ("TotalTrabalhadoMinutos", "Int64"),
    ("HorasPrevistasMinutos", "Int64"),
    ("FaltaAtrasoMinutos", "Int64"),
    ("AbonoMinutos", "Int64"),
    ("ExtrasTotalMinutos", "Int64"),
    ("BancoTotalMinutos", "Int64"),
    ("BancoSaldoMinutos", "Int64"),
    ("FaltasDias", "Int64"),
    ("TotalNormaisHoras", "Double"),
    ("TotalTrabalhadoHoras", "Double"),
    ("HorasPrevistasHoras", "Double"),
    ("FaltaAtrasoHoras", "Double"),
    ("AbonoHoras", "Double"),
    ("ExtrasTotalHoras", "Double"),
    ("BancoTotalHoras", "Double"),
    ("BancoSaldoHoras", "Double"),
    ("DiferencaJornadaMinutos", "Int64"),
    ("DiferencaJornadaHoras", "Double"),
    ("PercentualTrabalhado", "Double"),
    ("ClassificacaoSaldo", "string"),
    ("ClassificacaoSaldoOrdem", "Int64"),
    ("IndicadorSaldoCritico", "Int64"),
)


class PowerBiClient:
    def __init__(
        self,
        *,
        client_id: str = POWER_BI_CLIENT_ID,
        tenant_id: str = POWER_BI_TENANT_ID,
        workspace_id: str = POWER_BI_WORKSPACE_ID,
        dataset_name: str = POWER_BI_DATASET_NAME,
        table_name: str = POWER_BI_TABLE_NAME,
        timeout: float = 45.0,
        opener: Callable[..., Any] | None = None,
        app_factory: Callable[..., Any] | None = None,
    ) -> None:
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.workspace_id = workspace_id
        self.dataset_name = dataset_name
        self.table_name = table_name
        self.timeout = timeout
        self._opener = opener or urlopen
        self._app_factory = app_factory
        self._access_token: str | None = None

    @property
    def autenticado(self) -> bool:
        return bool(self._access_token)

    def login_interativo(self) -> None:
        """Autentica no navegador sem senha ou segredo gravados pelo aplicativo."""

        fabrica = self._app_factory
        if fabrica is None:
            try:
                from msal import PublicClientApplication
            except ImportError as erro:
                raise PowerBiApiError(
                    "O componente de login Microsoft não está instalado."
                ) from erro
            fabrica = PublicClientApplication

        app = fabrica(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
        )
        scopes = [POWER_BI_SCOPE]
        resposta = None
        contas = app.get_accounts()
        if contas:
            resposta = app.acquire_token_silent(scopes, account=contas[0])
        if not resposta or not resposta.get("access_token"):
            try:
                # O MSAL desktop inicia o listener em http://localhost e passa
                # o redirect_uri internamente. Repeti-lo aqui quebra o MSAL 1.37.
                resposta = app.acquire_token_interactive(
                    scopes=scopes,
                    prompt="select_account",
                )
            except Exception as erro:
                raise PowerBiApiError(
                    "Não foi possível abrir o login Microsoft."
                ) from erro

        token = resposta.get("access_token") if isinstance(resposta, dict) else None
        if not token:
            detalhe = resposta.get("error_description") if isinstance(resposta, dict) else None
            if detalhe and "AADSTS65001" in detalhe:
                detalhe = "A permissão do Power BI ainda não recebeu consentimento."
            raise PowerBiApiError(detalhe or "Não foi possível entrar no Power BI.")
        self._access_token = str(token)

    def obter_ou_criar_dataset(self) -> str:
        self._exigir_login()
        resposta = self._request("GET", f"/groups/{self.workspace_id}/datasets")
        datasets = resposta.get("value", []) if isinstance(resposta, dict) else []
        encontrados = [
            item for item in datasets if str(item.get("name", "")) == self.dataset_name
        ]
        if len(encontrados) > 1:
            raise PowerBiApiError(
                f"Existem vários modelos chamados '{self.dataset_name}' no workspace."
            )
        if encontrados:
            dataset = encontrados[0]
            if dataset.get("addRowsAPIEnabled") is False:
                raise PowerBiApiError(
                    "O modelo existente não aceita inclusão de linhas pela API."
                )
            identificador = dataset.get("id")
            if identificador:
                identificador = str(identificador)
                self._validar_tabela_existente(identificador)
                return identificador

        payload = {
            "name": self.dataset_name,
            "defaultMode": "Push",
            "tables": [
                {
                    "name": self.table_name,
                    "columns": [
                        {"name": nome, "dataType": tipo}
                        for nome, tipo in JORNADA_SCHEMA
                    ],
                }
            ],
        }
        criado = self._request(
            "POST",
            f"/groups/{self.workspace_id}/datasets?defaultRetentionPolicy=None",
            payload,
        )
        identificador = criado.get("id") if isinstance(criado, dict) else None
        if not identificador:
            raise PowerBiApiError("O Power BI não retornou o ID do modelo criado.")
        return str(identificador)

    def verificar_workspace(self) -> PowerBiWorkspaceInfo:
        """Confirma acesso ao workspace sem criar ou alterar recursos."""

        self._exigir_login()
        resposta = self._request("GET", f"/groups/{self.workspace_id}/datasets")
        if not isinstance(resposta, dict) or not isinstance(
            resposta.get("value"), list
        ):
            raise PowerBiApiError("O Power BI não confirmou o acesso ao workspace.")
        return PowerBiWorkspaceInfo(
            id=self.workspace_id,
            name=POWER_BI_WORKSPACE_NAME,
        )

    def verificar_modelo(self) -> PowerBiModelInfo | None:
        """Localiza e valida o modelo atual sem criar um modelo novo."""

        self._exigir_login()
        resposta = self._request("GET", f"/groups/{self.workspace_id}/datasets")
        datasets = resposta.get("value", []) if isinstance(resposta, dict) else []
        encontrados = [
            item for item in datasets if str(item.get("name", "")) == self.dataset_name
        ]
        if len(encontrados) > 1:
            raise PowerBiApiError(
                f"Existem vários modelos chamados '{self.dataset_name}' no workspace."
            )
        if not encontrados:
            return None
        dataset = encontrados[0]
        identificador = dataset.get("id")
        if not identificador:
            raise PowerBiApiError("O modelo encontrado não possui um identificador.")
        accepts_rows = dataset.get("addRowsAPIEnabled") is not False
        if not accepts_rows:
            raise PowerBiApiError(
                "O modelo existente não aceita inclusão de linhas pela API."
            )
        self._validar_tabela_existente(str(identificador))
        return PowerBiModelInfo(
            id=str(identificador),
            name=str(dataset.get("name") or self.dataset_name),
            accepts_rows=accepts_rows,
        )

    def _validar_tabela_existente(self, dataset_id: str) -> None:
        resposta = self._request(
            "GET",
            f"/groups/{self.workspace_id}/datasets/{dataset_id}/tables",
        )
        tabelas = resposta.get("value", []) if isinstance(resposta, dict) else []
        tabela = next(
            (item for item in tabelas if item.get("name") == self.table_name),
            None,
        )
        if tabela is None:
            raise PowerBiApiError(
                f"O modelo existente não possui a tabela '{self.table_name}'."
            )
        colunas = tabela.get("columns")
        # O endpoint de tabelas pode devolver somente o nome da tabela. Nesse
        # caso, a ausência da propriedade ``columns`` não significa que o
        # modelo esteja sem colunas; o POST de linhas fará a validação real do
        # contrato pelo próprio Power BI.
        if not isinstance(colunas, list) or not colunas:
            return

        atuais = {str(item.get("name", "")) for item in colunas}
        ausentes = [nome for nome, _tipo in JORNADA_SCHEMA if nome not in atuais]
        if ausentes:
            raise PowerBiApiError(
                "O modelo existente usa uma estrutura antiga. Remova-o do "
                "workspace e envie novamente para recriá-lo. Colunas ausentes: "
                + ", ".join(ausentes)
            )

    def enviar_linhas(
        self,
        dataset_id: str,
        linhas: Iterable[dict[str, Any]],
        *,
        tamanho_lote: int = 5_000,
        ao_progresso: Callable[[int, int], None] | None = None,
    ) -> int:
        self._exigir_login()
        if tamanho_lote <= 0:
            raise ValueError("tamanho_lote deve ser positivo")
        todas = list(linhas)
        total = len(todas)
        if not total:
            raise PowerBiApiError("Não há dados para enviar ao Power BI.")

        enviados = 0
        endpoint = (
            f"/groups/{self.workspace_id}/datasets/{dataset_id}"
            f"/tables/{self.table_name}/rows"
        )
        for inicio in range(0, total, tamanho_lote):
            lote = todas[inicio : inicio + tamanho_lote]
            self._request("POST", endpoint, {"rows": lote})
            enviados += len(lote)
            if ao_progresso:
                ao_progresso(enviados, total)
        return enviados

    def _exigir_login(self) -> None:
        if not self._access_token:
            raise PowerBiApiError("Entre com sua conta Microsoft antes de enviar.")

    def _request(
        self,
        metodo: str,
        caminho: str,
        payload: dict[str, Any] | None = None,
    ) -> Any:
        dados = None
        headers = {
            "Authorization": f"Bearer {self._access_token}",
            "Accept": "application/json",
        }
        if payload is not None:
            dados = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"
        requisicao = Request(
            f"{POWER_BI_API_URL}{caminho}",
            data=dados,
            headers=headers,
            method=metodo,
        )
        try:
            with self._opener(requisicao, timeout=self.timeout) as resposta:
                conteudo = resposta.read()
        except HTTPError as erro:
            try:
                corpo = json.loads(erro.read().decode("utf-8", errors="replace"))
                mensagem = corpo.get("error", {}).get("message")
            except Exception:
                mensagem = None
            if erro.code in (401, 403):
                mensagem = "Sua conta não tem acesso ao workspace ou à API do Power BI."
            raise PowerBiApiError(
                mensagem or f"O Power BI respondeu com erro HTTP {erro.code}."
            ) from erro
        except (URLError, TimeoutError, OSError) as erro:
            raise PowerBiApiError(
                "Não foi possível comunicar com o Power BI. Verifique sua conexão."
            ) from erro

        if not conteudo:
            return {}
        try:
            return json.loads(conteudo.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as erro:
            raise PowerBiApiError("O Power BI retornou uma resposta inválida.") from erro


__all__ = [
    "JORNADA_SCHEMA",
    "PowerBiApiError",
    "PowerBiClient",
    "PowerBiModelInfo",
    "PowerBiWorkspaceInfo",
]
