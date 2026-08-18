"""Fachada dos fluxos de CSV, RHiD e Power BI usados pela janela principal."""

from __future__ import annotations

import os
from tkinter import filedialog, messagebox

from app.controllers.csv_workflow import CsvWorkflow
from app.controllers.connection_diagnostics_workflow import (
    ConnectionDiagnosticsWorkflow,
)
from app.controllers.powerbi_workflow import PowerBiWorkflow
from app.controllers.rhid_workflow import RhidWorkflow
from app.controllers.workflow_state import WorkflowState
from app.core.logger import logger
from app.integrations.powerbi_client import PowerBiClient
from app.integrations.rhid_client import RhidClient
from app.integrations.rhid_report_service import processar_relatorio_rhid
from app.services.analytics_service import preparar_snapshot_powerbi
from app.services.audit_service import registrar_evento
from app.services.background_task_runner import BackgroundTaskRunner
from app.services.file_service import nome_curto
from app.services.history_service import registrar_historico, ultimos_processamentos
from app.services.powerbi_desktop_service import abrir_relatorio_powerbi_desktop
from app.services.powerbi_send_registry import PowerBiSendRegistry
from app.services.preferences_service import carregar_preferencias, salvar_preferencias
from app.services.validator_service import validar_arquivo_entrada
from app.services.workbook_pipeline_service import obter_departamentos, processar_arquivo
from app.ui.view_state import EstadoInterface


class MainController:
    """API estável da janela; cada integração vive em um fluxo especializado."""

    def __init__(self, view, task_runner=None, powerbi_registry=None):
        self.view = view
        self._state = WorkflowState(carregar_preferencias())

        self._csv = CsvWorkflow(
            view,
            self._state,
            operation_in_progress=lambda: self.processamento_em_andamento,
            task_runner=task_runner,
            validate_file=lambda path: validar_arquivo_entrada(path),
            list_departments=lambda path: obter_departamentos(path),
            process_file=lambda *args, **kwargs: processar_arquivo(*args, **kwargs),
            save_preferences=lambda data: salvar_preferencias(data),
            register_history_item=lambda item: registrar_historico(item),
            list_history=lambda: ultimos_processamentos(),
            audit=self._registrar_evento_seguro,
        )
        self._rhid = RhidWorkflow(
            view,
            self._state,
            operation_in_progress=lambda: self.processamento_em_andamento,
            runner_factory=lambda *args, **kwargs: BackgroundTaskRunner(
                *args, **kwargs
            ),
            client_factory=lambda: RhidClient(),
            report_processor=lambda *args, **kwargs: processar_relatorio_rhid(
                *args, **kwargs
            ),
            save_preferences=lambda data: salvar_preferencias(data),
            confirm_overwrite=self._csv.confirm_overwrite,
            register_history=self._registrar_historico_seguro,
            list_history=lambda: ultimos_processamentos(),
            audit=self._registrar_evento_seguro,
        )
        self._powerbi = PowerBiWorkflow(
            view,
            self._state,
            operation_in_progress=lambda: self.processamento_em_andamento,
            runner_factory=lambda *args, **kwargs: BackgroundTaskRunner(
                *args, **kwargs
            ),
            client_factory=lambda: PowerBiClient(),
            snapshot_factory=lambda *args, **kwargs: preparar_snapshot_powerbi(
                *args, **kwargs
            ),
            desktop_opener=lambda dataset_id: abrir_relatorio_powerbi_desktop(
                dataset_id
            ),
            registry=powerbi_registry or PowerBiSendRegistry(),
            confirm_duplicate=lambda title, text: messagebox.askyesno(title, text),
            audit=self._registrar_evento_seguro,
        )
        self._diagnostics = ConnectionDiagnosticsWorkflow(
            view,
            operation_in_progress=lambda: self.processamento_em_andamento,
            rhid_client_provider=lambda: self._rhid.client,
            powerbi_client_provider=lambda: self._powerbi.client,
            cache_powerbi_client=self._cache_powerbi_client,
            runner_factory=lambda *args, **kwargs: BackgroundTaskRunner(
                *args, **kwargs
            ),
            rhid_client_factory=lambda: RhidClient(),
            powerbi_client_factory=lambda: PowerBiClient(),
        )

    @property
    def processamento_em_andamento(self):
        return (
            self._csv.active
            or self._rhid.active
            or self._powerbi.active
            or self._diagnostics.active
        )

    @property
    def arquivos_selecionados(self):
        return self._state.selected_files

    @arquivos_selecionados.setter
    def arquivos_selecionados(self, value):
        self._state.selected_files = list(value)

    @property
    def ultimo_resultado(self):
        return self._state.last_result

    @ultimo_resultado.setter
    def ultimo_resultado(self, value):
        self._state.last_result = value

    @property
    def preferencias(self):
        return self._state.preferences

    # Propriedades de compatibilidade para testes e integrações anteriores.
    @property
    def _task_runner(self):
        return self._csv.runner

    @_task_runner.setter
    def _task_runner(self, value):
        self._csv.runner = value

    @property
    def _processando(self):
        return self._csv.processing

    @_processando.setter
    def _processando(self, value):
        self._csv.processing = bool(value)

    @property
    def _acumulador_lote(self):
        return self._csv.accumulator

    @_acumulador_lote.setter
    def _acumulador_lote(self, value):
        self._csv.accumulator = value

    @property
    def _rhid_client(self):
        return self._rhid.client

    @_rhid_client.setter
    def _rhid_client(self, value):
        self._rhid.client = value

    @property
    def _rhid_empresas(self):
        return self._rhid.companies

    @_rhid_empresas.setter
    def _rhid_empresas(self, value):
        self._rhid.companies = tuple(value)

    @property
    def _rhid_departamentos(self):
        return self._rhid.departments

    @_rhid_departamentos.setter
    def _rhid_departamentos(self, value):
        self._rhid.departments = tuple(value)

    @property
    def _rhid_runner(self):
        return self._rhid.runner

    @_rhid_runner.setter
    def _rhid_runner(self, value):
        self._rhid.runner = value

    @property
    def _powerbi_client(self):
        return self._powerbi.client

    @_powerbi_client.setter
    def _powerbi_client(self, value):
        self._powerbi.client = value

    @property
    def _powerbi_runner(self):
        return self._powerbi.runner

    @_powerbi_runner.setter
    def _powerbi_runner(self, value):
        self._powerbi.runner = value

    def iniciar(self):
        self.view.renderizar_historico(ultimos_processamentos())
        ultimo_depto = self.preferencias.get("last_department") or "Todos"
        self.view.atualizar_departamentos(["Todos"], selecionado=ultimo_depto)
        self.view.atualizar_pasta_saida(
            self.preferencias.get("last_save_dir")
            or "Nenhuma pasta selecionada ainda."
        )
        self.view.atualizar_versao()
        self.view.definir_estado(EstadoInterface.VAZIO)

    def conectar_rhid(self, email, senha, dominio=""):
        self._rhid.connect(email, senha, dominio)

    def gerar_relatorio_rhid(
        self,
        empresa_id,
        departamento_id,
        data_inicial,
        data_final,
        gerar_saldo=True,
        gerar_resumo=True,
        gerar_ranking=True,
    ):
        self._rhid.generate(
            empresa_id,
            departamento_id,
            data_inicial,
            data_final,
            gerar_saldo,
            gerar_resumo,
            gerar_ranking,
        )

    def limpar_selecao(self):
        self._csv.clear_selection()

    def selecionar_arquivos(self):
        self._csv.select_files()

    def processar(
        self,
        departamento,
        gerar_saldo=True,
        gerar_resumo=True,
        gerar_ranking=True,
    ):
        self._csv.process(
            departamento,
            gerar_saldo,
            gerar_resumo,
            gerar_ranking,
        )

    def cancelar_processamento(self):
        self._csv.cancel()

    def abrir_arquivo_gerado(self):
        if not self.ultimo_resultado:
            return
        try:
            os.startfile(self.ultimo_resultado["caminho_saida"])
        except Exception:
            messagebox.showerror("Erro", "Não foi possível abrir o arquivo gerado.")

    def abrir_pasta_gerada(self):
        if not self.ultimo_resultado:
            return
        try:
            os.startfile(os.path.dirname(self.ultimo_resultado["caminho_saida"]))
        except Exception:
            messagebox.showerror("Erro", "Não foi possível abrir a pasta de saída.")

    def enviar_ultimo_resultado_powerbi(self):
        self._powerbi.send_last()

    def acao_ultimo_resultado_powerbi(self):
        self._powerbi.action()

    def abrir_ultimo_resultado_powerbi(self):
        self._powerbi.open_last()

    def verificar_conexoes(self, email="", senha="", dominio=""):
        self._diagnostics.run(email, senha, dominio)

    def _cache_powerbi_client(self, client):
        self._powerbi.client = client

    def _confirmar_sobrescrita(self, caminho_saida):
        self._csv.confirm_overwrite(caminho_saida)

    @staticmethod
    def _registrar_historico_seguro(caminho_entrada, resultado):
        try:
            registrar_historico(
                {
                    "arquivo_origem": caminho_entrada,
                    "arquivo_saida": resultado["caminho_saida"],
                    "tipo_entrada": resultado["tipo_entrada"],
                    "quantidade_funcionarios": resultado["quantidade_funcionarios"],
                    "banco_total": resultado["banco_total"],
                    "banco_saldo": resultado["banco_saldo"],
                    "departamento": resultado["departamento"],
                    "gerou_saldo": resultado["gerou_saldo"],
                    "gerou_resumo": resultado["gerou_resumo"],
                    "gerou_ranking": resultado["gerou_ranking"],
                }
            )
        except Exception:
            logger.exception(
                "Arquivo salvo, mas não foi possível registrar o histórico de %s",
                nome_curto(caminho_entrada),
            )

    @staticmethod
    def _registrar_evento_seguro(acao, detalhes):
        try:
            registrar_evento(acao, detalhes)
        except Exception:
            logger.exception("Não foi possível registrar o evento %s", acao)


__all__ = ["MainController"]
