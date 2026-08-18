"""Orquestra o envio e a abertura de relatórios no Power BI."""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Callable, cast

from app.controllers.workflow_state import WorkflowState
from app.core.logger import logger
from app.integrations.powerbi_client import PowerBiApiError, PowerBiClient
from app.integrations.powerbi_destination import (
    PowerBiDestination,
    PowerBiDestinationKind,
    PushSemanticModelDestination,
)
from app.services.analytics_service import (
    AnalyticsError,
    PowerBiSnapshot,
    preparar_snapshot_powerbi,
)
from app.services.background_task_runner import BackgroundTaskRunner
from app.services.file_service import nome_curto
from app.services.powerbi_desktop_service import (
    PowerBiDesktopError,
    abrir_relatorio_powerbi_desktop,
)
from app.services.powerbi_send_registry import (
    PowerBiSendRecord,
    PowerBiSendRegistry,
    calcular_fingerprint_snapshot,
)


@dataclass(frozen=True, slots=True)
class _DuplicateFound:
    snapshot: PowerBiSnapshot
    fingerprint: str
    previous: PowerBiSendRecord


@dataclass(frozen=True, slots=True)
class _SendCompleted:
    destination: PowerBiDestination
    snapshot: PowerBiSnapshot
    resource_id: str
    sent_rows: int
    fingerprint: str


class PowerBiWorkflow:
    def __init__(
        self,
        view,
        state: WorkflowState,
        *,
        operation_in_progress: Callable[[], bool],
        runner_factory: Callable[..., Any] = BackgroundTaskRunner,
        client_factory: Callable[[], PowerBiClient] = PowerBiClient,
        destination_factory: Callable[[], PowerBiDestination] | None = None,
        snapshot_factory: Callable[..., PowerBiSnapshot] = preparar_snapshot_powerbi,
        desktop_opener: Callable[[str], Any] = abrir_relatorio_powerbi_desktop,
        registry: PowerBiSendRegistry | None = None,
        confirm_duplicate: Callable[[str, str], bool] | None = None,
        audit: Callable[[str, dict[str, Any]], None] | None = None,
    ) -> None:
        self.view = view
        self.state = state
        self._operation_in_progress = operation_in_progress
        self._runner_factory = runner_factory
        self._client_factory = client_factory
        self._destination_factory = destination_factory or (
            lambda: PushSemanticModelDestination(self._client_factory())
        )
        self._snapshot_factory = snapshot_factory
        self._desktop_opener = desktop_opener
        self._registry = registry or PowerBiSendRegistry()
        self._confirm_duplicate = confirm_duplicate or (lambda _titulo, _texto: False)
        self._audit = audit or (lambda _acao, _detalhes: None)
        self.destination: PowerBiDestination | None = None
        self.runner = None

    @property
    def client(self) -> PowerBiClient | None:
        """Compatibilidade com o cliente Push enquanto a migração é preparada."""

        destination = self.destination
        if isinstance(destination, PushSemanticModelDestination):
            return cast(PowerBiClient, destination.client)
        return None

    @client.setter
    def client(self, value: PowerBiClient | None) -> None:
        self.destination = PushSemanticModelDestination(value) if value else None

    @property
    def active(self) -> bool:
        return self.runner is not None and self.runner.ativo

    def action(self) -> None:
        resultado = self.state.last_result
        if resultado and resultado.get("powerbi_dataset_id"):
            self.open_last()
            return
        self.send_last()

    def send_last(self) -> None:
        """Prepara e envia o último relatório como snapshot independente."""

        if self._operation_in_progress():
            self.view.atualizar_status(
                "Já existe uma operação em andamento.", "warning"
            )
            return
        resultado = self.state.last_result
        if not resultado:
            self.view.exibir_erro_powerbi("Gere um relatório antes de enviar.")
            return
        if resultado.get("powerbi_dataset_id"):
            self.open_last()
            return

        caminho = resultado.get("caminho_saida")
        if not caminho or not os.path.isfile(caminho):
            self.view.exibir_erro_powerbi(
                "O arquivo gerado não foi encontrado. Gere o relatório novamente."
            )
            return
        self._start_send(dict(resultado), str(caminho))

    def _start_send(
        self,
        resultado: dict[str, Any],
        caminho: str,
        *,
        prepared: tuple[PowerBiSnapshot, str] | None = None,
        allow_duplicate: bool = False,
    ) -> None:
        self.runner = self._runner_factory(self.view.agendar_na_interface)
        self.view.definir_powerbi_ocupado(True)
        self.view.atualizar_progresso_powerbi(0.05, "Preparando os indicadores...")

        def execute(report):
            if prepared is None:
                snapshot = self._snapshot_factory(caminho, resultado)
                fingerprint = calcular_fingerprint_snapshot(snapshot)
            else:
                snapshot, fingerprint = prepared

            previous = self._registry.find(fingerprint)
            if previous is not None and not allow_duplicate:
                return _DuplicateFound(snapshot, fingerprint, previous)

            report((0.2, "Abrindo o login Microsoft..."))
            destination = self.destination or self._destination_factory()
            if not destination.authenticated:
                destination.login()
            report((0.45, "Localizando o modelo analítico..."))

            def progress(sent, total):
                fraction = 0.5 + (sent / total) * 0.48
                report((fraction, f"Enviando dados: {sent}/{total}"))

            published = destination.publish(
                snapshot.rows,
                on_progress=progress,
            )
            try:
                self._registry.register(
                    fingerprint=fingerprint,
                    report_id=snapshot.report_id,
                    dataset_id=published.resource_id,
                    source_file=caminho,
                    row_count=published.row_count,
                )
            except Exception:
                logger.exception(
                    "Dados enviados, mas não foi possível registrar a prevenção de duplicidade"
                )
            return _SendCompleted(
                destination,
                snapshot,
                published.resource_id,
                published.row_count,
                fingerprint,
            )

        def progress(event):
            value, message = event
            self.view.atualizar_progresso_powerbi(value, message)

        def complete(returned):
            if isinstance(returned, _DuplicateFound):
                self.view.definir_powerbi_ocupado(False)
                try:
                    date_text = datetime.fromisoformat(
                        returned.previous.sent_at
                    ).strftime("%d/%m/%Y às %H:%M")
                except ValueError:
                    date_text = returned.previous.sent_at
                confirmed = self._confirm_duplicate(
                    "Relatório já enviado",
                    "Este mesmo conteúdo já foi enviado ao Power BI "
                    f"em {date_text} (ID {returned.previous.report_id}).\n\n"
                    "Deseja enviar novamente mesmo assim?",
                )
                if not confirmed:
                    self.view.atualizar_progresso_powerbi(
                        0, "Reenvio cancelado para evitar dados duplicados."
                    )
                    self.view.atualizar_status(
                        "O relatório já estava no Power BI; o reenvio foi cancelado.",
                        "warning",
                    )
                    return
                self._start_send(
                    resultado,
                    caminho,
                    prepared=(returned.snapshot, returned.fingerprint),
                    allow_duplicate=True,
                )
                return

            assert isinstance(returned, _SendCompleted)
            self.destination = returned.destination
            if self.state.last_result is not None:
                self.state.last_result["powerbi_report_id"] = returned.snapshot.report_id
                self.state.last_result["powerbi_resource_id"] = returned.resource_id
                self.state.last_result["powerbi_destination"] = (
                    returned.destination.kind.value
                )
                if (
                    returned.destination.kind
                    is PowerBiDestinationKind.PUSH_SEMANTIC_MODEL
                ):
                    self.state.last_result["powerbi_dataset_id"] = returned.resource_id
                self.state.last_result["powerbi_fingerprint"] = returned.fingerprint
            self.view.definir_powerbi_ocupado(False)
            self.view.definir_powerbi_enviado(True)
            self.view.atualizar_progresso_powerbi(1.0, "Envio concluído.")
            label = (
                "1 funcionário enviado"
                if returned.sent_rows == 1
                else f"{returned.sent_rows} funcionários enviados"
            )
            self.view.exibir_sucesso_powerbi(
                f"{label}. ID do relatório: {returned.snapshot.report_id}"
            )
            self._audit(
                "envio_powerbi",
                {
                    "id_relatorio": returned.snapshot.report_id,
                    "recurso_id": returned.resource_id,
                    "destino": returned.destination.kind.value,
                    "quantidade_linhas": returned.sent_rows,
                    "arquivo": nome_curto(caminho),
                    "fingerprint": returned.fingerprint[:12],
                },
            )
            self.open_last()

        def fail(error):
            logger.warning("Falha ao enviar relatório ao Power BI: %s", error)
            message = (
                str(error)
                if isinstance(error, (PowerBiApiError, AnalyticsError))
                else "Não foi possível enviar o relatório ao Power BI."
            )
            self.view.definir_powerbi_ocupado(False)
            self.view.exibir_erro_powerbi(message)

        try:
            started = self.runner.executar(execute, progress, complete, fail)
        except Exception:
            logger.exception("Não foi possível iniciar o envio ao Power BI")
            self.view.definir_powerbi_ocupado(False)
            self.view.exibir_erro_powerbi(
                "Não foi possível iniciar o envio ao Power BI."
            )
            return
        if not started:
            self.view.definir_powerbi_ocupado(False)
            self.view.exibir_erro_powerbi("Já existe um envio em andamento.")

    def open_last(self) -> None:
        """Abre um relatório fino conectado ao modelo sem duplicar o envio."""

        resultado = self.state.last_result
        if not resultado:
            self.view.exibir_erro_powerbi(
                "Gere e envie um relatório antes de abrir."
            )
            return
        dataset_id = resultado.get("powerbi_dataset_id")
        if not dataset_id:
            self.view.exibir_erro_powerbi(
                "Envie o relatório ao Power BI antes de abri-lo no Desktop."
            )
            return
        try:
            self._desktop_opener(str(dataset_id))
        except PowerBiDesktopError as error:
            logger.warning("Falha ao abrir o Power BI Desktop: %s", error)
            self.view.exibir_erro_powerbi(str(error))
            return
        self.view.definir_powerbi_enviado(True)
        self.view.exibir_sucesso_powerbi(
            "Power BI Desktop aberto. A tabela Jornada está pronta para os gráficos."
        )


__all__ = ["PowerBiWorkflow"]
