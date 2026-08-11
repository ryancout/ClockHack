"""Cliente mínimo para a API oficial do RHiD."""

from __future__ import annotations

import base64
import json
import time
import unicodedata
from dataclasses import dataclass
from datetime import date
from typing import Any, Callable
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from app.core.logger import logger


DEFAULT_BASE_URL = "https://rhid.com.br/v2/api.svc"
DEFAULT_SSO_URL = "https://sso-backend.controlid.com.br:5000/api"
DEFAULT_WEB_URL = "https://rhid.com.br/v2"

# Contrato do CSV que já é aceito e testado pelo ClockHack. Os nomes internos
# são obtidos do catálogo oficial exportableAcjefAndPersonColumns.
_COLUNAS_CSV_CLOCKHACK = (
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
)


class RhidApiError(Exception):
    """Erro seguro e apresentável ocorrido na comunicação com o RHiD."""


@dataclass(frozen=True, slots=True)
class RhidTenant:
    domain: str
    name: str
    tenant_id: int | None = None
    system_id: int | None = None


class RhidTenantRequired(RhidApiError):
    def __init__(self, tenants: tuple[RhidTenant, ...]):
        self.tenants = tenants
        super().__init__("Selecione o cliente do RHiD e conecte novamente.")


@dataclass(frozen=True, slots=True)
class RhidCompany:
    id: int
    name: str
    trading_name: str = ""
    cnpj: str = ""

    @property
    def label(self) -> str:
        return self.trading_name or self.name or f"Empresa {self.id}"


@dataclass(frozen=True, slots=True)
class RhidDepartment:
    id: int
    name: str
    company_id: int


@dataclass(frozen=True, slots=True)
class RhidPerson:
    id: int
    name: str
    registration: str
    company_id: int
    department_id: int
    status: int | None = None


class RhidClient:
    """Consome somente endpoints de leitura necessários ao ClockHack."""

    def __init__(
        self,
        base_url: str = DEFAULT_BASE_URL,
        sso_url: str = DEFAULT_SSO_URL,
        web_url: str = DEFAULT_WEB_URL,
        timeout: float = 30.0,
        opener: Callable[..., Any] | None = None,
        sleeper: Callable[[float], None] | None = None,
    ):
        if not base_url.lower().startswith("https://"):
            raise ValueError("A API do RHiD deve ser acessada por HTTPS.")
        self.base_url = base_url.rstrip("/")
        self.sso_url = sso_url.rstrip("/")
        self.web_url = web_url.rstrip("/")
        self.timeout = timeout
        self._opener = opener or urlopen
        self._sleeper = sleeper or time.sleep
        self._access_token: str | None = None
        self._customer_id: str | None = None

    @property
    def autenticado(self) -> bool:
        return bool(self._access_token)

    def login(self, email: str, password: str, domain: str = "") -> None:
        email = (email or "").strip()
        if not email or not password:
            raise RhidApiError("Informe o e-mail e a senha do RHiD.")

        payload = {
            "email": email,
            "password": password,
            "domain": None,
            "cpf": None,
            "loginType": "operator",
            "supportedSystems": [1, 3],
        }
        resposta = self._login_sso(payload)
        token = self._token_sso(resposta)
        tenants = self._tenants_sso(resposta)

        if not token and tenants:
            dominio = (domain or "").strip()
            if dominio:
                tenant = next(
                    (item for item in tenants if item.domain.casefold() == dominio.casefold()),
                    None,
                )
                if tenant is None:
                    raise RhidApiError("O cliente selecionado não foi retornado pelo RHiD.")
            elif len(tenants) == 1:
                tenant = tenants[0]
            else:
                raise RhidTenantRequired(tenants)

            segundo_payload = payload.copy()
            segundo_payload["domain"] = tenant.domain
            if tenant.tenant_id is not None:
                segundo_payload["tenantId"] = tenant.tenant_id
            if tenant.system_id is not None:
                segundo_payload["systemId"] = tenant.system_id
            resposta = self._login_sso(segundo_payload)
            token = self._token_sso(resposta)

        if not token:
            detalhe = resposta.get("error") if isinstance(resposta, dict) else None
            raise RhidApiError(detalhe or "O RHiD não retornou um token de acesso.")
        self._access_token = str(token)
        self._customer_id = self._claim_token(self._access_token, "cidCustomerId")

    def _login_sso(self, payload: dict) -> Any:
        try:
            return self._request(
                "POST",
                "/auth/login",
                payload=payload,
                authenticated=False,
                base_url=self.sso_url,
                extra_headers={"Accept-Language": "pt_BR"},
            )
        except RhidApiError as exc:
            raise RhidApiError(f"Falha na autenticação SSO: {exc}") from exc

    @staticmethod
    def _token_sso(resposta: Any) -> str | None:
        if not isinstance(resposta, dict):
            return None
        if resposta.get("accessToken"):
            return str(resposta["accessToken"])
        sistemas = resposta.get("Data")
        if isinstance(sistemas, list) and len(sistemas) == 1:
            primeiro = sistemas[0]
            token = primeiro.get("token") if isinstance(primeiro, dict) else None
            return str(token) if token else None
        return None

    @staticmethod
    def _tenants_sso(resposta: Any) -> tuple[RhidTenant, ...]:
        if not isinstance(resposta, dict):
            return ()

        tenants = []
        sistemas = resposta.get("Data")
        if isinstance(sistemas, list):
            for sistema in sistemas:
                if not isinstance(sistema, dict):
                    continue
                system_id = sistema.get("systemId")
                for tenant in sistema.get("tenants") or []:
                    if not isinstance(tenant, dict) or not tenant.get("domain"):
                        continue
                    if tenant.get("status") is False:
                        continue
                    tenants.append(
                        RhidTenant(
                            domain=str(tenant["domain"]).strip(),
                            name=str(tenant.get("name") or "").strip(),
                            tenant_id=(
                                int(tenant["tenantId"])
                                if tenant.get("tenantId") is not None
                                else None
                            ),
                            system_id=(int(system_id) if system_id is not None else None),
                        )
                    )
        if tenants:
            return tuple(tenants)

        clientes = resposta.get("listCustomer")
        return tuple(
            RhidTenant(
                domain=str(cliente.get("domain") or "").strip(),
                name=str(cliente.get("name") or "").strip(),
                tenant_id=(
                    int(cliente.get("tenantId", cliente.get("id")))
                    if cliente.get("tenantId", cliente.get("id")) is not None
                    else None
                ),
                system_id=(
                    int(cliente["systemId"])
                    if cliente.get("systemId") is not None
                    else None
                ),
            )
            for cliente in (clientes or [])
            if isinstance(cliente, dict) and cliente.get("domain")
        )

    def listar_empresas(self) -> tuple[RhidCompany, ...]:
        registros = self._records(
            self._request(
                "GET",
                "/customerdb/company.svc/a",
                base_url=self.web_url,
            )
        )
        return tuple(
            RhidCompany(
                id=int(item["id"]),
                name=str(item.get("name") or ""),
                trading_name=str(item.get("tradingName") or ""),
                cnpj=str(item.get("cnpj") or ""),
            )
            for item in registros
            if item.get("id") is not None
        )

    def listar_departamentos(self) -> tuple[RhidDepartment, ...]:
        registros = self._records(
            self._request(
                "GET",
                "/customerdb/department.svc/a",
                base_url=self.web_url,
            )
        )
        return tuple(
            RhidDepartment(
                id=int(item["id"]),
                name=str(item.get("name") or f"Departamento {item['id']}"),
                company_id=int(item.get("idCompany") or 0),
            )
            for item in registros
            if item.get("id") is not None
        )

    def listar_pessoas(self) -> tuple[RhidPerson, ...]:
        registros = self._records(
            self._request(
                "GET",
                "/customerdb/person.svc/a_resumido",
                base_url=self.web_url,
            )
        )
        return tuple(
            RhidPerson(
                id=int(item["id"]),
                name=str(item.get("name") or ""),
                registration=str(item.get("registration") or ""),
                company_id=int(item.get("idCompany") or 0),
                department_id=int(item.get("idDepartment") or 0),
                status=(int(item["status"]) if item.get("status") is not None else None),
            )
            for item in registros
            if item.get("id") is not None
        )

    def listar_ids_pessoas_ativas(
        self,
        company_id: int | None = None,
        department_id: int | None = None,
    ) -> tuple[int, ...]:
        """Replica o filtro server-side usado pelo RHiD antes do CSV customizado."""

        registros = self._registros_pessoas_ativas(company_id, department_id)
        return tuple(
            dict.fromkeys(
                int(item["id"])
                for item in registros
                if item.get("id") is not None
            )
        )

    def filtrar_departamentos_com_pessoas_ativas(
        self,
        departamentos: tuple[RhidDepartment, ...] | list[RhidDepartment],
    ) -> tuple[RhidDepartment, ...]:
        """Mantém apenas setores que possuem ao menos um funcionário ativo."""

        departamentos = tuple(departamentos)
        registros = self._registros_pessoas_ativas()
        ids_departamentos = {
            self._id_relacionado(item, "idDepartment", "departmentId", "department")
            for item in registros
        }
        ids_departamentos.discard(None)
        if ids_departamentos:
            return tuple(
                departamento
                for departamento in departamentos
                if departamento.id in ids_departamentos
            )

        # Compatibilidade com respostas resumidas que omitem idDepartment:
        # o próprio servidor responde se há pessoas para cada setor.
        if registros:
            return tuple(
                departamento
                for departamento in departamentos
                if self._registros_pessoas_ativas(department_id=departamento.id)
            )
        return ()

    def _registros_pessoas_ativas(
        self,
        company_id: int | None = None,
        department_id: int | None = None,
    ) -> list[dict]:
        parametros = {}
        if company_id is not None:
            parametros["companies"] = int(company_id)
        if department_id is not None:
            parametros["departments"] = int(department_id)
        return self._records(
            self._request(
                "GET",
                "/customerdb/person.svc/a_ativo",
                params=parametros or None,
                base_url=self.web_url,
            )
        )

    @staticmethod
    def _id_relacionado(item: dict, campo: str, alias: str, objeto: str) -> int | None:
        valor = item.get(campo, item.get(alias))
        if valor is None and isinstance(item.get(objeto), dict):
            valor = item[objeto].get("id")
        try:
            return int(valor) if valor is not None else None
        except (TypeError, ValueError):
            return None

    def obter_apuracao(
        self,
        person_id: int,
        data_inicial: date | str,
        data_final: date | str,
    ) -> tuple[dict, ...]:
        inicio = self._data_iso(data_inicial)
        fim = self._data_iso(data_final)
        if fim < inicio:
            raise RhidApiError("A data final não pode ser anterior à data inicial.")
        if (fim - inicio).days > 90:
            raise RhidApiError("O período da consulta ao RHiD não pode ultrapassar 90 dias.")

        resposta = self._request(
            "GET",
            "/apuracao_ponto",
            params={
                "dataIni": inicio.isoformat(),
                "dataFinal": fim.isoformat(),
                "idPerson": int(person_id),
            },
        )
        if isinstance(resposta, str):
            try:
                resposta = json.loads(resposta)
            except json.JSONDecodeError as exc:
                raise RhidApiError("O RHiD retornou uma apuração em formato inválido.") from exc
        if not isinstance(resposta, list):
            raise RhidApiError("O RHiD retornou uma apuração em formato inesperado.")
        return tuple(item for item in resposta if isinstance(item, dict))

    def gerar_relatorio_csv(
        self,
        company_id: int | None,
        department_id: int | None,
        data_inicial: date | str,
        data_final: date | str,
        ao_progresso: Callable[[int], None] | None = None,
        limite_espera_segundos: int = 600,
    ) -> bytes:
        """Gera o extrato CSV consolidado por funcionário usado pelo RHiD."""

        inicio, fim = self._validar_periodo_relatorio(data_inicial, data_final)
        if company_id is not None and int(company_id) <= 0:
            raise RhidApiError("Selecione uma empresa válida do RHiD.")
        if department_id is not None and int(department_id) <= 0:
            raise RhidApiError("Selecione um setor válido do RHiD.")

        propriedades = self._propriedades_relatorio()
        colunas_acjef, informacoes_pessoa = self._separar_colunas(propriedades)
        formato_horas = self._formato_horas()
        reportar = ao_progresso or (lambda _valor: None)
        ids_pessoas = self.listar_ids_pessoas_ativas(company_id, department_id)
        if not ids_pessoas:
            raise RhidApiError("Nenhum funcionário ativo foi encontrado nesse escopo.")
        logger.info(
            "Relatório RHiD: solicitando extrato para %d funcionário(s).",
            len(ids_pessoas),
        )

        parametros = {
            "fontSizeTitle": 12,
            "fontSizeData": 8,
            "fontSizeHeader": 9,
            "fontSizeHeaderSmall": 8,
            "fontSizeFooter": 8,
            "fontName": "Helvetica",
            "listIdStr": list(ids_pessoas),
            "listCostCenterStr": [],
            "listDepartmentStr": [int(department_id)] if department_id is not None else [],
            "listPersonRoleStr": [],
            "listCompanyStr": [int(company_id)] if company_id is not None else [],
            "listShiftStr": [],
        }
        payload = {
            "pdfCartaoPontoParameters": parametros,
            "ini": inicio.strftime("%Y%m%d"),
            "fim": fim.strftime("%Y%m%d"),
            "listColumns": colunas_acjef,
            "listPersonInfo": informacoes_pessoa,
            "formatoSaida": "CSV",
            "agrupamento": "person",
            "relatorio": "extrato",
            "formatoHoras": formato_horas,
            "status": 1,
        }

        resposta = self._request(
            "POST",
            "/report.svc/ponto",
            payload=payload,
            base_url=self.web_url,
        )
        if not isinstance(resposta, dict):
            raise RhidApiError("O RHiD não iniciou a geração do relatório.")
        if resposta.get("error"):
            raise RhidApiError(str(resposta["error"]))
        guid = str(resposta.get("guid") or "").strip()
        logger.info(
            "Relatório RHiD iniciado: numPeople=%s, guid=%s.",
            resposta.get("numPeople"),
            "presente" if guid else "ausente",
        )
        try:
            nenhuma_pessoa = int(resposta.get("numPeople")) == 0
        except (TypeError, ValueError):
            nenhuma_pessoa = False
        if nenhuma_pessoa:
            raise RhidApiError(
                "O RHiD não encontrou funcionários com apuração nesse período."
            )
        if not guid:
            raise RhidApiError("O RHiD não retornou o identificador do relatório.")

        reportar(0)
        inicio_espera = time.monotonic()
        while True:
            situacao = self._request(
                "GET",
                "/customerdb/notify.svc/specificGuid/",
                params={"guid": guid},
                base_url=self.web_url,
            )
            if not isinstance(situacao, dict):
                raise RhidApiError("O RHiD retornou um status de relatório inválido.")
            try:
                percentual = int(situacao.get("percent", 0))
            except (TypeError, ValueError):
                percentual = 0
            if percentual == -1:
                raise RhidApiError(str(situacao.get("error") or "Falha ao gerar o relatório no RHiD."))
            reportar(max(0, min(percentual, 100)))
            if percentual >= 100:
                break
            if time.monotonic() - inicio_espera >= limite_espera_segundos:
                raise RhidApiError("O RHiD demorou demais para gerar o relatório.")
            self._sleeper(1.0)

        conteudo = self._request_bytes(
            "POST",
            "/customerdb/notify.svc/save_file/",
            params={"format": "CSV", "guid": guid},
            base_url=self.web_url,
        )
        logger.info(
            "Relatório RHiD concluído: percentual=%d, bytes=%d.",
            percentual,
            len(conteudo),
        )
        if not conteudo:
            raise RhidApiError(
                "O RHiD gerou um extrato vazio para o período selecionado."
            )
        return conteudo

    def _propriedades_relatorio(self) -> list[dict]:
        resposta = self._request(
            "GET",
            "/maindb/layoutproperty.svc/exportableAcjefAndPersonColumns",
            base_url=self.web_url,
        )
        if isinstance(resposta, dict):
            propriedades = resposta.get("data") or resposta.get("records")
        else:
            propriedades = resposta
        if not isinstance(propriedades, list):
            raise RhidApiError("O RHiD não retornou as colunas disponíveis do relatório.")
        propriedades = [item for item in propriedades if isinstance(item, dict)]
        if not propriedades:
            raise RhidApiError("O RHiD não disponibilizou colunas para o relatório.")

        por_propriedade = {
            (
                str(item.get("className") or "").casefold(),
                str(item.get("propertyName") or "").casefold(),
            ): item
            for item in propriedades
        }
        por_cabecalho = {
            (
                str(item.get("className") or "").casefold(),
                self._normalizar_texto(item.get("headerText")),
            ): item
            for item in propriedades
            if item.get("headerText")
        }

        selecionadas = []
        ausentes = []
        for classe, propriedade, cabecalho in _COLUNAS_CSV_CLOCKHACK:
            # O nome interno varia entre layouts do RHiD. O cabeçalho é o
            # contrato do CSV; o propertyName conhecido fica como fallback.
            item = por_cabecalho.get(
                (classe.casefold(), self._normalizar_texto(cabecalho))
            )
            if item is None:
                item = por_propriedade.get((classe.casefold(), propriedade.casefold()))
            if item is None:
                ausentes.append(cabecalho)
            else:
                selecionado = dict(item)
                selecionado["multiplica"] = False
                selecionado["multiplicaNoturno"] = False
                selecionadas.append(selecionado)
        if ausentes:
            raise RhidApiError(
                "O layout do RHiD não disponibilizou estas colunas: "
                + ", ".join(ausentes)
                + "."
            )
        logger.info(
            "Colunas do extrato RHiD: %s",
            [
                f"{item.get('className')}.{item.get('propertyName')}"
                for item in selecionadas
            ],
        )
        return selecionadas

    @staticmethod
    def _normalizar_texto(valor: Any) -> str:
        texto = unicodedata.normalize("NFKD", str(valor or ""))
        return " ".join(
            "".join(letra for letra in texto if not unicodedata.combining(letra))
            .casefold()
            .split()
        )

    @staticmethod
    def _separar_colunas(propriedades: list[dict]) -> tuple[list[str], list[str]]:
        colunas_acjef: list[str] = []
        informacoes_pessoa: list[str] = []
        for item in propriedades:
            propriedade = str(item.get("propertyName") or "").strip()
            classe = str(item.get("className") or "").strip()
            if not propriedade or not classe:
                continue
            if classe.casefold() == "acjef":
                for sufixo in (
                    "_MultiplicadoExNot",
                    "_MultiplicadoEx",
                    "_MultiplicadoNot",
                    "_Multiplicado",
                ):
                    propriedade = propriedade.replace(sufixo, "")
                if item.get("multiplica") and item.get("multiplicaNoturno"):
                    propriedade += "_MultiplicadoExNot"
                elif item.get("multiplica"):
                    propriedade += "_MultiplicadoEx"
                elif item.get("multiplicaNoturno"):
                    propriedade += "_MultiplicadoNot"
                colunas_acjef.append(propriedade)
            else:
                informacoes_pessoa.append(f"{classe}.{propriedade}")

        identificadores = {
            "Person.name",
            "Person.registration",
            "Person.numFolha",
            "Person.cpf",
            "Person.pis",
        }
        if not identificadores.intersection(informacoes_pessoa):
            informacoes_pessoa = [
                "Person.name",
                "Person.registration",
                "PersonRole.name",
                *informacoes_pessoa,
            ]
        if not colunas_acjef:
            raise RhidApiError("O layout padrão do RHiD não possui colunas de horas.")
        return colunas_acjef, list(dict.fromkeys(informacoes_pessoa))

    def _formato_horas(self) -> int:
        try:
            resposta = self._request(
                "GET",
                "/maindb/parameter.svc/parameters_global",
                base_url=self.web_url,
            )
        except RhidApiError:
            return 1
        parametros = resposta.get("data") if isinstance(resposta, dict) else resposta
        if not isinstance(parametros, list):
            return 1
        for parametro in parametros:
            if isinstance(parametro, dict) and str(parametro.get("id")) == "118":
                try:
                    return int(parametro.get("valueN", 1))
                except (TypeError, ValueError):
                    return 1
        return 1

    def _validar_periodo_relatorio(
        self,
        data_inicial: date | str,
        data_final: date | str,
    ) -> tuple[date, date]:
        inicio = self._data_iso(data_inicial)
        fim = self._data_iso(data_final)
        if fim < inicio:
            raise RhidApiError("A data final não pode ser anterior à data inicial.")
        if (fim - inicio).days > 31:
            raise RhidApiError("O relatório do RHiD permite no máximo 31 dias.")
        return inicio, fim

    @staticmethod
    def _data_iso(valor: date | str) -> date:
        if isinstance(valor, date):
            return valor
        try:
            return date.fromisoformat(str(valor))
        except ValueError as exc:
            raise RhidApiError("Use datas no formato AAAA-MM-DD.") from exc

    @staticmethod
    def _records(resposta: Any) -> list[dict]:
        if not isinstance(resposta, dict):
            raise RhidApiError("O RHiD retornou uma lista em formato inesperado.")
        registros = resposta.get("records")
        if not isinstance(registros, list):
            registros = resposta.get("data")
        if not isinstance(registros, list):
            raise RhidApiError("O RHiD retornou uma lista em formato inesperado.")
        return [item for item in registros if isinstance(item, dict)]

    def _request(
        self,
        method: str,
        path: str,
        payload: dict | None = None,
        params: dict | None = None,
        authenticated: bool = True,
        base_url: str | None = None,
        extra_headers: dict | None = None,
    ) -> Any:
        if authenticated and not self._access_token:
            raise RhidApiError("Conecte-se ao RHiD antes de consultar os dados.")

        url = f"{(base_url or self.base_url).rstrip('/')}{path}"
        if params:
            url = f"{url}?{urlencode(params)}"
        headers = {"Accept": "application/json"}
        headers.update(extra_headers or {})
        data = None
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json; charset=utf-8"
        if authenticated:
            headers["Authorization"] = f"Bearer {self._access_token}"
            if self._customer_id:
                headers["X-Cid-RHiD"] = self._customer_id

        request = Request(url, data=data, headers=headers, method=method)
        raw = self._abrir(request)
        if not raw:
            return {}
        try:
            return json.loads(raw.decode("utf-8-sig"))
        except (json.JSONDecodeError, UnicodeDecodeError) as exc:
            raise RhidApiError("O RHiD retornou uma resposta inválida.") from exc

    def _request_bytes(
        self,
        method: str,
        path: str,
        params: dict | None = None,
        base_url: str | None = None,
    ) -> bytes:
        if not self._access_token:
            raise RhidApiError("Conecte-se ao RHiD antes de consultar os dados.")
        url = f"{(base_url or self.base_url).rstrip('/')}{path}"
        if params:
            url = f"{url}?{urlencode(params)}"
        headers = {
            "Accept": "application/vnd.ms-excel, application/octet-stream, */*",
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json; charset=utf-8",
        }
        if self._customer_id:
            headers["X-Cid-RHiD"] = self._customer_id
        return self._abrir(Request(url, headers=headers, method=method))

    def _abrir(self, request: Request) -> bytes:
        try:
            with self._opener(request, timeout=self.timeout) as response:
                return response.read()
        except HTTPError as exc:
            detalhe = self._detalhe_erro_http(exc)
            if exc.code == 401:
                raise RhidApiError(
                    detalhe or "Acesso negado pelo RHiD. Verifique a conta e as permissões."
                ) from exc
            if exc.code == 400:
                raise RhidApiError(detalhe or "O RHiD recusou os dados enviados.") from exc
            raise RhidApiError(detalhe or f"O RHiD respondeu com erro HTTP {exc.code}.") from exc
        except (URLError, TimeoutError, OSError) as exc:
            raise RhidApiError("Não foi possível conectar ao RHiD.") from exc

    @staticmethod
    def _claim_token(token: str, claim: str) -> str | None:
        try:
            payload = token.split(".")[1]
            payload += "=" * (-len(payload) % 4)
            dados = json.loads(base64.urlsafe_b64decode(payload).decode("utf-8"))
            valor = dados.get(claim)
            return str(valor) if valor is not None else None
        except (IndexError, ValueError, UnicodeDecodeError, json.JSONDecodeError):
            return None

    @staticmethod
    def _detalhe_erro_http(erro: HTTPError) -> str:
        try:
            conteudo = erro.read().decode("utf-8-sig")
            resposta = json.loads(conteudo)
        except json.JSONDecodeError:
            texto = conteudo.strip()
            if texto and len(texto) <= 500 and not texto.startswith("<"):
                return texto
            return ""
        except (AttributeError, UnicodeDecodeError):
            return ""
        if isinstance(resposta, dict):
            codigos = resposta.get("errorCodes")
            if isinstance(codigos, list) and codigos:
                return ", ".join(str(codigo) for codigo in codigos)
            for campo in ("error", "message", "Message", "ExceptionMessage", "data"):
                valor = resposta.get(campo)
                if isinstance(valor, str) and valor.strip():
                    return valor.strip()
        if isinstance(resposta, str):
            return resposta.strip()
        return ""
