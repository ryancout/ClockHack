"""Orquestra autenticação, catálogo e geração de relatórios do RHiD."""

from __future__ import annotations

import os
import re
import time
from datetime import date
from typing import Any, Callable

from tkinter import filedialog

from app.controllers.workflow_state import WorkflowState
from app.core.exceptions import AppError, SobrescritaCanceladaError
from app.core.logger import logger
from app.integrations.rhid_client import RhidApiError, RhidClient, RhidTenantRequired
from app.integrations.rhid_report_service import RhidReportPlan, processar_relatorio_rhid
from app.services.background_task_runner import BackgroundTaskRunner
from app.services.file_service import garantir_extensao_xlsx
from app.ui.view_state import EstadoInterface


class RhidWorkflow:
    def __init__(
        self,
        view,
        state: WorkflowState,
        *,
        operation_in_progress: Callable[[], bool],
        runner_factory: Callable[..., Any] = BackgroundTaskRunner,
        client_factory: Callable[[], RhidClient] = RhidClient,
        report_processor: Callable[..., dict[str, Any]] = processar_relatorio_rhid,
        save_preferences: Callable[[dict[str, Any]], None],
        confirm_overwrite: Callable[[str], None],
        register_history: Callable[[str, dict[str, Any]], None],
        list_history: Callable[[], list[Any]],
        audit: Callable[[str, dict[str, Any]], None],
    ) -> None:
        self.view = view
        self.state = state
        self._operation_in_progress = operation_in_progress
        self._runner_factory = runner_factory
        self._client_factory = client_factory
        self._report_processor = report_processor
        self._save_preferences = save_preferences
        self._confirm_overwrite = confirm_overwrite
        self._register_history = register_history
        self._list_history = list_history
        self._audit = audit
        self.client: RhidClient | None = None
        self.companies = ()
        self.departments = ()
        self.runner = None

    @property
    def active(self) -> bool:
        return self.runner is not None and self.runner.ativo

    def connect(self, email, password, domain="") -> None:
        """Autentica e carrega o catálogo organizacional sem persistir a senha."""

        if self.active:
            return
        if not str(email).strip() or not password:
            self.view.exibir_erro_rhid("Informe o e-mail e a senha do RHiD.")
            return

        self.view.definir_conexao_rhid_ocupada(True)
        self.runner = self._runner_factory(self.view.agendar_na_interface)

        def load_catalog(_report):
            client = self._client_factory()
            client.login(email, password, domain)
            try:
                companies = client.listar_empresas()
            except RhidApiError as error:
                raise RhidApiError(
                    f"Login realizado, mas falhou ao carregar empresas: {error}"
                ) from error
            try:
                registered = client.listar_departamentos()
                departments = client.filtrar_departamentos_com_pessoas_ativas(
                    registered
                )
            except RhidApiError as error:
                raise RhidApiError(
                    f"Login realizado, mas falhou ao carregar setores ativos: {error}"
                ) from error
            return client, companies, departments

        def complete(result):
            client, companies, departments = result
            self.client = client
            self.companies = tuple(companies)
            self.departments = tuple(departments)
            self.view.definir_conexao_rhid_ocupada(False)
            self.view.exibir_catalogo_rhid(companies, departments)
            self.view.atualizar_status("Conectado ao RHiD.", "success")

        def fail(error):
            logger.warning("Falha ao conectar ao RHiD: %s", error)
            message = (
                str(error)
                if isinstance(error, RhidApiError)
                else "Não foi possível conectar ao RHiD."
            )
            self.view.definir_conexao_rhid_ocupada(False)
            self.view.exibir_erro_rhid(message)
            if isinstance(error, RhidTenantRequired):
                self.view.exibir_dominios_rhid(error.tenants)
            self.view.atualizar_status(message, "error")

        self.runner.executar(load_catalog, lambda _event: None, complete, fail)

    def generate(
        self,
        company_id,
        department_id,
        start_date,
        end_date,
        generate_balance=True,
        generate_summary=True,
        generate_ranking=True,
    ) -> None:
        """Gera e trata o CSV oficial do RHiD sem download manual."""

        if self._operation_in_progress():
            self.view.atualizar_status(
                "Já existe um processamento em andamento.", "warning"
            )
            return
        if self.client is None or not self.client.autenticado:
            self.view.exibir_erro_rhid(
                "Conecte-se ao RHiD antes de gerar o relatório."
            )
            return

        try:
            normalized_company_id = int(company_id) if company_id is not None else None
        except (TypeError, ValueError):
            self.view.exibir_erro_rhid("Selecione uma empresa válida do RHiD.")
            return

        try:
            if department_id is None:
                department_ids = ()
                department_filter = None
            elif isinstance(department_id, (list, tuple)):
                department_ids = tuple(
                    dict.fromkeys(int(item) for item in department_id)
                )
                department_filter = department_ids or None
            else:
                only_department = int(department_id)
                department_ids = (only_department,)
                department_filter = only_department
        except (TypeError, ValueError):
            self.view.exibir_erro_rhid("Selecione setores válidos do RHiD.")
            return

        try:
            start = date.fromisoformat(str(start_date))
            end = date.fromisoformat(str(end_date))
            if end < start:
                raise RhidApiError(
                    "A data final não pode ser anterior à data inicial."
                )

            company = None
            if normalized_company_id is not None:
                company = next(
                    (
                        item
                        for item in self.companies
                        if item.id == normalized_company_id
                    ),
                    None,
                )
                if company is None:
                    raise RhidApiError("Selecione uma empresa válida do RHiD.")

            departments = []
            for item_id in department_ids:
                department = next(
                    (
                        item
                        for item in self.departments
                        if item.id == item_id
                        and (
                            normalized_company_id is None
                            or item.company_id in (0, normalized_company_id)
                        )
                    ),
                    None,
                )
                if department is None:
                    raise RhidApiError(
                        "Selecione um setor válido para essa empresa."
                    )
                departments.append(department)
        except ValueError:
            self.view.exibir_erro_rhid(
                "Informe datas válidas no formato DD/MM/AAAA."
            )
            return
        except RhidApiError as error:
            self.view.exibir_erro_rhid(str(error))
            return

        company_label = company.label if company is not None else "Todas as empresas"
        if len(departments) == 1:
            department_label = departments[0].name
        elif departments:
            department_label = f"{len(departments)} setores selecionados"
        else:
            department_label = "Todos os setores"

        scope_name = department_label if departments else company_label
        safe_name = re.sub(r"[^\w.-]+", "_", scope_name, flags=re.UNICODE).strip(
            "_."
        )
        default_name = (
            f"relatorio_rhid_{safe_name or normalized_company_id}_"
            f"{start:%Y%m%d}_{end:%Y%m%d}.xlsx"
        )
        output_path = filedialog.asksaveasfilename(
            title="Salvar relatório do RHiD como",
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialdir=self.state.preferences.get("last_save_dir") or None,
            initialfile=default_name,
        )
        if not output_path:
            return
        output_path = garantir_extensao_xlsx(output_path)
        try:
            self._confirm_overwrite(output_path)
        except SobrescritaCanceladaError:
            return

        plan = RhidReportPlan(
            company_id=normalized_company_id,
            department_id=department_filter,
            company_label=company_label,
            department_label=department_label,
            data_inicial=start,
            data_final=end,
            caminho_saida=output_path,
            gerar_saldo=generate_balance,
            gerar_resumo=generate_summary,
            gerar_ranking=generate_ranking,
        )
        client = self.client
        self.runner = self._runner_factory(self.view.agendar_na_interface)
        self.view.definir_geracao_rhid_ocupada(True)
        self.view.atualizar_progresso_rhid(0, "Solicitando relatório ao RHiD...")
        self.view.atualizar_status(
            "Gerando relatório diretamente no RHiD...", "info"
        )
        self.view.atualizar_progresso(0.02)
        self.view.habilitar_botao_abrir(False)
        self.view.habilitar_botao_abrir_pasta(False)
        self.view.definir_estado(EstadoInterface.PROCESSANDO, total_arquivos=1)

        def execute(report):
            started_at = time.perf_counter()
            result = self._report_processor(client, plan, report)
            return result, time.perf_counter() - started_at

        def progress(value):
            value = max(0, min(int(value), 100))
            message = (
                "Tratando o CSV e montando o Excel..."
                if value >= 80
                else f"RHiD processando o relatório: {value}%"
            )
            self.view.atualizar_progresso_rhid(value, message)
            self.view.atualizar_progresso(max(0.02, value / 100))

        def complete(result_with_time):
            result, total_time = result_with_time
            result = dict(result)
            result.update(
                {
                    "empresa": company_label,
                    "setores": department_label,
                    "periodo_inicial": start.isoformat(),
                    "periodo_final": end.isoformat(),
                }
            )
            self.state.last_result = result
            self.state.preferences["last_save_dir"] = os.path.dirname(output_path)
            self._save_preferences(self.state.preferences)
            self.view.definir_geracao_rhid_ocupada(False)
            self.view.atualizar_progresso_rhid(100, "Excel gerado e salvo.")
            self.view.exibir_sucesso_rhid(f"Salvo em: {output_path}")
            self.view.atualizar_arquivo(
                f"RHiD: {company_label} / {department_label} / "
                f"{start:%d/%m/%Y} a {end:%d/%m/%Y}"
            )
            self.view.atualizar_pasta_saida(os.path.dirname(output_path))
            self.view.atualizar_metricas(
                result["quantidade_funcionarios"],
                result["banco_total"],
                result["banco_saldo"],
            )
            self.view.atualizar_tempo_execucao(total_time)
            self.view.atualizar_progresso(1.0)
            self.view.habilitar_botao_abrir(True)
            self.view.habilitar_botao_abrir_pasta(True)
            self.view.definir_estado(EstadoInterface.CONCLUIDO, total_arquivos=1)
            self.view.atualizar_status(
                "Relatório do RHiD processado e salvo.", "success"
            )
            self._register_history(f"RHiD — {company_label}", result)
            self.view.renderizar_historico(self._list_history())
            self._audit(
                "processamento_rhid",
                {
                    "empresa_id": normalized_company_id,
                    "departamento_ids": list(department_ids),
                    "periodo_inicial": start.isoformat(),
                    "periodo_final": end.isoformat(),
                    "caminho_saida": output_path,
                },
            )

        def fail(error):
            logger.warning("Falha ao gerar relatório do RHiD: %s", error)
            message = (
                str(error)
                if isinstance(error, (RhidApiError, AppError))
                else "Não foi possível gerar o relatório do RHiD."
            )
            self.view.definir_geracao_rhid_ocupada(False)
            self.view.exibir_erro_rhid(message)
            self.view.atualizar_status(message, "error")
            self.view.definir_estado(EstadoInterface.ERRO, total_arquivos=1)

        self.runner.executar(execute, progress, complete, fail)


__all__ = ["RhidWorkflow"]
