"""Diagnóstico de integrações sem gerar, criar ou enviar relatórios."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any, Callable

from app.core.logger import logger
from app.integrations.powerbi_client import PowerBiApiError, PowerBiClient
from app.integrations.rhid_client import RhidApiError, RhidClient, RhidTenantRequired
from app.services.background_task_runner import BackgroundTaskRunner


class ConnectionKey(str, Enum):
    RHID = "rhid"
    MICROSOFT = "microsoft"
    WORKSPACE = "workspace"
    POWERBI_MODEL = "powerbi_model"


class ConnectionStatus(str, Enum):
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    WARNING = "warning"
    ERROR = "error"
    SKIPPED = "skipped"


@dataclass(frozen=True, slots=True)
class ConnectionCheckResult:
    key: ConnectionKey
    status: ConnectionStatus
    message: str


@dataclass(frozen=True, slots=True)
class _DiagnosticCompletion:
    results: tuple[ConnectionCheckResult, ...]
    powerbi_client: PowerBiClient | None


class ConnectionDiagnosticsWorkflow:
    def __init__(
        self,
        view,
        *,
        operation_in_progress: Callable[[], bool],
        rhid_client_provider: Callable[[], RhidClient | None],
        powerbi_client_provider: Callable[[], PowerBiClient | None],
        cache_powerbi_client: Callable[[PowerBiClient], None],
        runner_factory: Callable[..., Any] = BackgroundTaskRunner,
        rhid_client_factory: Callable[[], RhidClient] = RhidClient,
        powerbi_client_factory: Callable[[], PowerBiClient] = PowerBiClient,
    ) -> None:
        self.view = view
        self._operation_in_progress = operation_in_progress
        self._rhid_client_provider = rhid_client_provider
        self._powerbi_client_provider = powerbi_client_provider
        self._cache_powerbi_client = cache_powerbi_client
        self._runner_factory = runner_factory
        self._rhid_client_factory = rhid_client_factory
        self._powerbi_client_factory = powerbi_client_factory
        self.runner = None

    @property
    def active(self) -> bool:
        return self.runner is not None and self.runner.ativo

    def run(self, email: str = "", password: str = "", domain: str = "") -> None:
        if self._operation_in_progress():
            self.view.finalizar_diagnostico(
                "Aguarde a operação atual terminar antes de verificar as conexões.",
                "warning",
            )
            return

        credentials = (str(email or "").strip(), password or "", str(domain or "").strip())
        self.runner = self._runner_factory(self.view.agendar_na_interface)
        self.view.reiniciar_diagnostico()
        self.view.definir_diagnostico_ocupado(True)

        def emit(report, key, status, message):
            result = ConnectionCheckResult(key, status, message)
            report(result)
            return result

        def execute(report):
            results: list[ConnectionCheckResult] = []

            report(ConnectionCheckResult(ConnectionKey.RHID, ConnectionStatus.RUNNING, "Validando acesso..."))
            rhid_client = self._rhid_client_provider()
            try:
                if rhid_client is None or not rhid_client.autenticado:
                    if not credentials[0] or not credentials[1]:
                        raise RhidApiError(
                            "Informe ou lembre o acesso na tela do RHiD para testar."
                        )
                    rhid_client = self._rhid_client_factory()
                    rhid_client.login(*credentials)
                companies = rhid_client.listar_empresas()
                results.append(
                    emit(
                        report,
                        ConnectionKey.RHID,
                        ConnectionStatus.SUCCESS,
                        f"RHiD acessível; {len(companies)} empresa(s) encontrada(s).",
                    )
                )
            except RhidTenantRequired:
                results.append(
                    emit(
                        report,
                        ConnectionKey.RHID,
                        ConnectionStatus.WARNING,
                        "Acesso válido, mas é necessário selecionar o cliente no RHiD.",
                    )
                )
            except RhidApiError as error:
                results.append(
                    emit(report, ConnectionKey.RHID, ConnectionStatus.ERROR, str(error))
                )

            report(ConnectionCheckResult(ConnectionKey.MICROSOFT, ConnectionStatus.RUNNING, "Abrindo autenticação Microsoft..."))
            powerbi_client = self._powerbi_client_provider() or self._powerbi_client_factory()
            try:
                if not powerbi_client.autenticado:
                    powerbi_client.login_interativo()
                results.append(
                    emit(
                        report,
                        ConnectionKey.MICROSOFT,
                        ConnectionStatus.SUCCESS,
                        "Conta Microsoft autenticada.",
                    )
                )
            except PowerBiApiError as error:
                results.append(
                    emit(report, ConnectionKey.MICROSOFT, ConnectionStatus.ERROR, str(error))
                )
                for key, message in (
                    (ConnectionKey.WORKSPACE, "Não verificado porque o login Microsoft falhou."),
                    (ConnectionKey.POWERBI_MODEL, "Não verificado porque o login Microsoft falhou."),
                ):
                    results.append(emit(report, key, ConnectionStatus.SKIPPED, message))
                return _DiagnosticCompletion(tuple(results), None)

            report(ConnectionCheckResult(ConnectionKey.WORKSPACE, ConnectionStatus.RUNNING, "Consultando workspace..."))
            try:
                workspace = powerbi_client.verificar_workspace()
                results.append(
                    emit(
                        report,
                        ConnectionKey.WORKSPACE,
                        ConnectionStatus.SUCCESS,
                        f"Workspace acessível: {workspace.name}.",
                    )
                )
            except PowerBiApiError as error:
                results.append(
                    emit(report, ConnectionKey.WORKSPACE, ConnectionStatus.ERROR, str(error))
                )
                results.append(
                    emit(
                        report,
                        ConnectionKey.POWERBI_MODEL,
                        ConnectionStatus.SKIPPED,
                        "Não verificado porque o workspace não está acessível.",
                    )
                )
                return _DiagnosticCompletion(tuple(results), powerbi_client)

            report(ConnectionCheckResult(ConnectionKey.POWERBI_MODEL, ConnectionStatus.RUNNING, "Validando modelo analítico..."))
            try:
                model = powerbi_client.verificar_modelo()
                if model is None:
                    results.append(
                        emit(
                            report,
                            ConnectionKey.POWERBI_MODEL,
                            ConnectionStatus.WARNING,
                            "Modelo ainda não existe; será criado no primeiro envio.",
                        )
                    )
                else:
                    results.append(
                        emit(
                            report,
                            ConnectionKey.POWERBI_MODEL,
                            ConnectionStatus.SUCCESS,
                            f"Modelo válido: {model.name}.",
                        )
                    )
            except PowerBiApiError as error:
                results.append(
                    emit(
                        report,
                        ConnectionKey.POWERBI_MODEL,
                        ConnectionStatus.ERROR,
                        str(error),
                    )
                )
            return _DiagnosticCompletion(tuple(results), powerbi_client)

        def progress(result):
            self.view.atualizar_diagnostico(
                result.key.value,
                result.status.value,
                result.message,
            )

        def complete(completion):
            if completion.powerbi_client is not None:
                self._cache_powerbi_client(completion.powerbi_client)
            self.view.definir_diagnostico_ocupado(False)
            errors = sum(item.status is ConnectionStatus.ERROR for item in completion.results)
            warnings = sum(item.status is ConnectionStatus.WARNING for item in completion.results)
            if errors:
                self.view.finalizar_diagnostico(
                    f"Verificação concluída com {errors} falha(s).", "error"
                )
            elif warnings:
                self.view.finalizar_diagnostico(
                    "Conexões principais válidas, com uma atenção pendente.", "warning"
                )
            else:
                self.view.finalizar_diagnostico(
                    "Todas as conexões estão funcionando.", "success"
                )

        def fail(error):
            logger.warning("Falha inesperada no diagnóstico das integrações: %s", error)
            self.view.definir_diagnostico_ocupado(False)
            self.view.finalizar_diagnostico(
                "Não foi possível concluir a verificação das conexões.", "error"
            )

        started = self.runner.executar(execute, progress, complete, fail)
        if not started:
            self.view.definir_diagnostico_ocupado(False)
            self.view.finalizar_diagnostico(
                "Já existe uma verificação em andamento.", "warning"
            )


__all__ = [
    "ConnectionCheckResult",
    "ConnectionDiagnosticsWorkflow",
    "ConnectionKey",
    "ConnectionStatus",
]
