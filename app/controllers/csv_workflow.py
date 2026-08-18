"""Orquestra seleção e processamento em lote de relatórios CSV."""

from __future__ import annotations

import os
import time
from dataclasses import dataclass
from typing import Any, Callable

from tkinter import filedialog, messagebox

from app.controllers.workflow_state import WorkflowState
from app.core.config import TIPOS_ARQUIVO_ENTRADA
from app.core.exceptions import AppError, SobrescritaCanceladaError
from app.core.logger import logger
from app.domain import formatar_horas, para_minutos
from app.services.background_task_runner import BackgroundTaskRunner
from app.services.file_service import garantir_extensao_xlsx, nome_curto, sugerir_nome_saida
from app.services.validator_service import validar_arquivo_entrada
from app.services.workbook_pipeline_service import obter_departamentos, processar_arquivo
from app.ui.view_state import EstadoInterface


@dataclass(frozen=True, slots=True)
class BatchItem:
    input_path: str
    output_path: str


@dataclass(frozen=True, slots=True)
class BatchPlan:
    items: tuple[BatchItem, ...]
    output_dir: str
    department: str
    generate_balance: bool
    generate_summary: bool
    generate_ranking: bool


@dataclass(frozen=True, slots=True)
class FileStarted:
    index: int
    total: int
    item: BatchItem


@dataclass(frozen=True, slots=True)
class FileCompleted:
    index: int
    total: int
    item: BatchItem
    result: dict[str, Any]


@dataclass(slots=True)
class BatchAccumulator:
    employee_count: int = 0
    total_bank_minutes: int = 0
    balance_bank_minutes: int = 0
    processed: int = 0
    skipped: int = 0
    last_result: dict[str, Any] | None = None


class BatchProcessingFailure(Exception):
    def __init__(self, input_path, cause, expected):
        super().__init__(str(cause))
        self.input_path = input_path
        self.cause = cause
        self.expected = expected


class CancellationRequested(Exception):
    pass


class CsvWorkflow:
    def __init__(
        self,
        view,
        state: WorkflowState,
        *,
        operation_in_progress: Callable[[], bool],
        task_runner=None,
        validate_file: Callable[[str], Any] = validar_arquivo_entrada,
        list_departments: Callable[[str], list[str]] = obter_departamentos,
        process_file: Callable[..., dict[str, Any]] = processar_arquivo,
        save_preferences: Callable[[dict[str, Any]], None],
        register_history_item: Callable[[dict[str, Any]], None],
        list_history: Callable[[], list[Any]],
        audit: Callable[[str, dict[str, Any]], None],
    ) -> None:
        self.view = view
        self.state = state
        self._operation_in_progress = operation_in_progress
        self.runner = task_runner or BackgroundTaskRunner(view.agendar_na_interface)
        self._validate_file = validate_file
        self._list_departments = list_departments
        self._process_file = process_file
        self._save_preferences = save_preferences
        self._register_history_item = register_history_item
        self._list_history = list_history
        self._audit = audit
        self.processing = False
        self.accumulator: BatchAccumulator | None = None

    @property
    def active(self) -> bool:
        return self.processing or self.runner.ativo

    def clear_selection(self) -> None:
        if self._operation_in_progress():
            self.view.atualizar_status(
                "Aguarde o processamento terminar para limpar a seleção.",
                "warning",
            )
            return

        self.state.selected_files = []
        self.state.last_result = None
        self.view.atualizar_arquivo("Nenhum arquivo selecionado")
        self.view.atualizar_departamentos(["Todos"], selecionado="Todos")
        self.view.atualizar_metricas(0, "--:--", "--:--")
        self.view.atualizar_progresso(0)
        self.view.atualizar_status(
            "Seleção limpa. Escolha um novo arquivo para continuar.", "info"
        )
        self.view.atualizar_tempo_execucao(None)
        self.view.habilitar_botao_abrir(False)
        self.view.habilitar_botao_abrir_pasta(False)
        self.view.definir_estado(EstadoInterface.VAZIO)

    def select_files(self) -> None:
        if self._operation_in_progress():
            self.view.atualizar_status(
                "Aguarde o processamento terminar para selecionar novos arquivos.",
                "warning",
            )
            return

        try:
            initial_dir = self.state.preferences.get("last_open_dir")
            if not initial_dir or not os.path.exists(initial_dir):
                initial_dir = os.path.expanduser("~")

            paths = filedialog.askopenfilenames(
                title="Selecione o(s) arquivo(s) CSV",
                filetypes=TIPOS_ARQUIVO_ENTRADA,
                initialdir=initial_dir,
            )
            if not paths:
                self.view.atualizar_status("Nenhum arquivo selecionado.", "warning")
                return

            valid_files = []
            for path in paths:
                self._validate_file(path)
                valid_files.append(path)

            self.state.selected_files = valid_files
            self.state.last_result = None
            self.state.preferences["last_open_dir"] = os.path.dirname(valid_files[0])
            self._save_preferences(self.state.preferences)

            total = len(valid_files)
            names = [nome_curto(item) for item in valid_files[:3]]
            selection_label = (
                "1 arquivo selecionado"
                if total == 1
                else f"{total} arquivos selecionados"
            )
            text = f"{selection_label}: " + ", ".join(names)
            if total > 3:
                text += " ..."

            self.view.atualizar_arquivo(text)
            self.view.atualizar_metricas(0, "--:--", "--:--")
            self.view.atualizar_progresso(0)
            self.view.atualizar_tempo_execucao(None)
            self.view.habilitar_botao_abrir(False)
            self.view.habilitar_botao_abrir_pasta(False)

            departments = self._list_departments(valid_files[0])
            selected = self.state.preferences.get("last_department") or "Todos"
            self.view.atualizar_departamentos(departments, selecionado=selected)
            loading_message = "Arquivo carregado" if total == 1 else "Arquivos carregados"
            self.view.atualizar_status(
                f"{loading_message}. Ajuste as opções e clique em Processar.",
                "info",
            )
            self.view.definir_estado(EstadoInterface.PRONTO, total_arquivos=total)

            logger.info("Arquivos selecionados: %s", names)
            self._audit(
                "arquivos_selecionados",
                {"quantidade": total, "arquivos": [nome_curto(x) for x in valid_files]},
            )
        except AppError as error:
            logger.warning("Validação ao selecionar arquivos: %s", error)
            self.view.atualizar_status(str(error), "error")
            messagebox.showerror("Arquivo inválido", str(error))
        except Exception:
            logger.exception("Erro ao selecionar arquivos")
            self.view.atualizar_status(
                "Não foi possível carregar os arquivos selecionados.", "error"
            )
            messagebox.showerror(
                "Erro ao selecionar arquivos",
                "Não foi possível carregar os arquivos selecionados. "
                "Verifique se eles estão fechados e tente novamente.",
            )

    @staticmethod
    def confirm_overwrite(output_path: str) -> None:
        if os.path.exists(output_path):
            confirmed = messagebox.askyesno(
                "Confirmar substituição",
                f"Já existe um arquivo com este nome:\n\n"
                f"{nome_curto(output_path)}\n\nDeseja substituir?",
            )
            if not confirmed:
                raise SobrescritaCanceladaError(
                    "Gravação cancelada para evitar sobrescrita de arquivo existente."
                )

    @staticmethod
    def _validate_unique_destinations(items: tuple[BatchItem, ...]) -> None:
        destinations = {}
        for item in items:
            normalized = os.path.normcase(os.path.abspath(item.output_path))
            if normalized in destinations:
                first = destinations[normalized]
                raise AppError(
                    "Dois arquivos do lote gerariam a mesma saída: "
                    f"{nome_curto(first.input_path)} e {nome_curto(item.input_path)}. "
                    "Renomeie um dos arquivos de entrada e tente novamente."
                )
            destinations[normalized] = item

    def _prepare_plan(
        self,
        files,
        department,
        generate_balance,
        generate_summary,
        generate_ranking,
    ) -> BatchPlan | None:
        if len(files) == 1:
            default_name = sugerir_nome_saida(files[0], department)
            output_path = filedialog.asksaveasfilename(
                title="Salvar arquivo tratado como",
                defaultextension=".xlsx",
                filetypes=[("Arquivo Excel", "*.xlsx")],
                initialdir=self.state.preferences.get("last_save_dir")
                or self.state.preferences.get("last_open_dir")
                or None,
                initialfile=default_name,
            )
            if not output_path:
                return None
            output_path = garantir_extensao_xlsx(output_path)
            output_dir = os.path.dirname(output_path)
            items = (BatchItem(files[0], output_path),)
        else:
            output_dir = filedialog.askdirectory(
                title="Selecione a pasta onde os arquivos tratados serão salvos",
                initialdir=self.state.preferences.get("last_save_dir")
                or self.state.preferences.get("last_open_dir")
                or None,
            )
            if not output_dir:
                return None
            items = tuple(
                BatchItem(
                    file,
                    os.path.join(output_dir, sugerir_nome_saida(file, department)),
                )
                for file in files
            )

        self._validate_unique_destinations(items)
        for item in items:
            self.confirm_overwrite(item.output_path)
        return BatchPlan(
            items,
            output_dir,
            department,
            generate_balance,
            generate_summary,
            generate_ranking,
        )

    def process(
        self,
        department,
        generate_balance=True,
        generate_summary=True,
        generate_ranking=True,
    ) -> None:
        if self._operation_in_progress():
            self.view.atualizar_status(
                "Já existe um processamento em andamento.", "warning"
            )
            return
        if not self.state.selected_files:
            messagebox.showwarning("Aviso", "Selecione um ou mais arquivos primeiro.")
            return

        files = tuple(self.state.selected_files)
        total_files = len(files)
        try:
            plan = self._prepare_plan(
                files,
                department,
                generate_balance,
                generate_summary,
                generate_ranking,
            )
            if plan is None:
                self.view.atualizar_status("Operação cancelada pelo usuário.", "warning")
                self.view.definir_estado(
                    EstadoInterface.PRONTO, total_arquivos=total_files
                )
                return
        except SobrescritaCanceladaError as error:
            self.view.atualizar_status(str(error), "warning")
            self.view.definir_estado(EstadoInterface.PRONTO, total_arquivos=total_files)
            return
        except AppError as error:
            self.view.atualizar_status(str(error), "error")
            self.view.definir_estado(EstadoInterface.PRONTO, total_arquivos=total_files)
            messagebox.showerror("Não foi possível preparar o lote", str(error))
            return

        self.state.preferences["last_save_dir"] = plan.output_dir
        self.state.preferences["last_department"] = department or "Todos"
        self._save_preferences(self.state.preferences)
        self.view.atualizar_pasta_saida(plan.output_dir)
        self.view.atualizar_status("Processando arquivo(s)...", "info")
        self.view.atualizar_progresso(0.02)
        self.view.habilitar_botao_abrir(False)
        self.view.habilitar_botao_abrir_pasta(False)
        self.view.definir_estado(EstadoInterface.PROCESSANDO, total_arquivos=total_files)
        self.processing = True
        self.accumulator = BatchAccumulator()

        try:
            started = self.runner.executar(
                lambda report: self._execute_plan(plan, report),
                self._receive_event,
                lambda duration: self._finish_batch(plan, duration),
                lambda error: self._fail_batch(plan, error),
            )
        except Exception as error:
            self._fail_batch(plan, error)
            return
        if not started:
            self.processing = False
            self.accumulator = None
            self.view.atualizar_status(
                "Já existe um processamento em andamento.", "warning"
            )
            self.view.definir_estado(EstadoInterface.PRONTO, total_arquivos=total_files)

    def cancel(self) -> None:
        if not self.processing or not self.runner.cancelar():
            return
        self.view.atualizar_status(
            "Cancelando após concluir o arquivo atual...", "warning"
        )
        self.view.definir_estado(
            EstadoInterface.CANCELANDO,
            total_arquivos=len(self.state.selected_files),
        )

    def _execute_plan(self, plan: BatchPlan, report) -> float:
        started_at = time.perf_counter()
        total = len(plan.items)
        for index, item in enumerate(plan.items, start=1):
            if self.runner.cancelamento_solicitado:
                raise CancellationRequested()
            report(FileStarted(index, total, item))
            try:
                result = self._process_file(
                    item.input_path,
                    item.output_path,
                    plan.department,
                    gerar_saldo=plan.generate_balance,
                    gerar_ranking=plan.generate_ranking,
                    gerar_resumo=plan.generate_summary,
                )
            except AppError as error:
                logger.warning(
                    "Falha de validação/processamento em %s: %s",
                    nome_curto(item.input_path),
                    error,
                )
                raise BatchProcessingFailure(item.input_path, error, True) from error
            except Exception as error:
                logger.exception(
                    "Erro inesperado no processamento de %s",
                    nome_curto(item.input_path),
                )
                raise BatchProcessingFailure(item.input_path, error, False) from error
            report(FileCompleted(index, total, item, result))

        if self.runner.cancelamento_solicitado:
            raise CancellationRequested()
        return time.perf_counter() - started_at

    def _receive_event(self, event) -> None:
        if isinstance(event, FileStarted):
            progress = 0.05 + ((event.index - 1) / event.total) * 0.9
            self.view.atualizar_progresso(progress)
            self.view.atualizar_status(
                f"Processando {event.index}/{event.total}: "
                f"{nome_curto(event.item.input_path)}",
                "info",
            )
            return
        if not isinstance(event, FileCompleted) or self.accumulator is None:
            return

        result = event.result
        self.accumulator.employee_count += result["quantidade_funcionarios"]
        self.accumulator.total_bank_minutes += para_minutos(result["banco_total"])
        self.accumulator.balance_bank_minutes += para_minutos(result["banco_saldo"])
        self.accumulator.processed += 1
        self.accumulator.last_result = result
        progress = 0.05 + (event.index / event.total) * 0.9
        self.view.atualizar_progresso(progress)
        self._register_history_safely(event.item.input_path, result)

    def _register_history_safely(self, input_path, result) -> None:
        try:
            self._register_history_item(
                {
                    "arquivo_origem": input_path,
                    "arquivo_saida": result["caminho_saida"],
                    "tipo_entrada": result["tipo_entrada"],
                    "quantidade_funcionarios": result["quantidade_funcionarios"],
                    "banco_total": result["banco_total"],
                    "banco_saldo": result["banco_saldo"],
                    "departamento": result["departamento"],
                    "gerou_saldo": result["gerou_saldo"],
                    "gerou_resumo": result["gerou_resumo"],
                    "gerou_ranking": result["gerou_ranking"],
                }
            )
        except Exception:
            logger.exception(
                "Arquivo salvo, mas não foi possível registrar o histórico de %s",
                nome_curto(input_path),
            )

    @staticmethod
    def _safe_callback(description, callback) -> bool:
        try:
            callback()
        except Exception:
            logger.exception("Falha ao %s", description)
            return False
        return True

    def _expose_partial_results(self, accumulator: BatchAccumulator) -> None:
        if not accumulator.processed:
            return
        self.state.last_result = accumulator.last_result
        self._safe_callback(
            "habilitar a abertura do arquivo salvo",
            lambda: self.view.habilitar_botao_abrir(True),
        )
        self._safe_callback(
            "habilitar a abertura da pasta de saída",
            lambda: self.view.habilitar_botao_abrir_pasta(True),
        )
        self._safe_callback(
            "atualizar as métricas parciais",
            lambda: self.view.atualizar_metricas(
                accumulator.employee_count,
                formatar_horas(accumulator.total_bank_minutes),
                formatar_horas(accumulator.balance_bank_minutes),
            ),
        )
        self._safe_callback(
            "atualizar o histórico após processamento parcial",
            lambda: self.view.renderizar_historico(self._list_history()),
        )

    @staticmethod
    def _partial_summary(accumulator: BatchAccumulator, reason: str) -> str:
        if not accumulator.processed:
            return reason
        saved = (
            "1 arquivo foi salvo"
            if accumulator.processed == 1
            else f"{accumulator.processed} arquivos foram salvos"
        )
        return f"{reason} {saved} antes da interrupção."

    def _finish_cancellation(
        self, plan: BatchPlan, accumulator: BatchAccumulator
    ) -> None:
        self.processing = False
        self._expose_partial_results(accumulator)
        message = self._partial_summary(accumulator, "Processamento cancelado.")
        self._safe_callback(
            "atualizar o status de cancelamento",
            lambda: self.view.atualizar_status(message, "warning"),
        )
        self._safe_callback(
            "aplicar o estado cancelado",
            lambda: self.view.definir_estado(
                EstadoInterface.CANCELADO, total_arquivos=len(plan.items)
            ),
        )
        self.accumulator = None

    def _finish_batch(self, plan: BatchPlan, total_time: float) -> None:
        accumulator = self.accumulator or BatchAccumulator()
        self.processing = False
        self.view.atualizar_progresso(1.0)
        if accumulator.processed == 0:
            self.view.atualizar_status(
                "Nenhum arquivo foi processado com sucesso.", "error"
            )
            self.view.atualizar_tempo_execucao(None)
            self.view.definir_estado(
                EstadoInterface.ERRO, total_arquivos=len(plan.items)
            )
            messagebox.showerror(
                "Erro", "Nenhum arquivo foi processado com sucesso."
            )
            self.accumulator = None
            return

        self.state.last_result = accumulator.last_result
        self.view.habilitar_botao_abrir(True)
        self.view.habilitar_botao_abrir_pasta(True)
        self.view.definir_estado(
            EstadoInterface.CONCLUIDO, total_arquivos=len(plan.items)
        )
        self.view.atualizar_metricas(
            accumulator.employee_count,
            formatar_horas(accumulator.total_bank_minutes),
            formatar_horas(accumulator.balance_bank_minutes),
        )
        self.view.atualizar_tempo_execucao(total_time)
        saved = (
            "1 arquivo foi salvo"
            if accumulator.processed == 1
            else f"{accumulator.processed} arquivos foram salvos"
        )
        self.view.atualizar_status(
            f"Processamento concluído. {saved}. Ignorados: {accumulator.skipped} "
            f"| Filtro: {plan.department}",
            "success",
        )
        self.view.renderizar_historico(self._list_history())
        self._audit(
            "processamento_lote",
            {
                "processados": accumulator.processed,
                "ignorados": accumulator.skipped,
                "departamento": plan.department,
                "pasta_saida": plan.output_dir,
                "gerou_saldo": plan.generate_balance,
                "gerou_resumo": plan.generate_summary,
                "gerou_ranking": plan.generate_ranking,
                "tempo_execucao_segundos": round(total_time, 2),
            },
        )
        self.accumulator = None

    def _fail_batch(self, plan: BatchPlan, error: Exception) -> None:
        self.processing = False
        accumulator = self.accumulator or BatchAccumulator()
        if isinstance(error, CancellationRequested):
            self._finish_cancellation(plan, accumulator)
            return

        accumulator.skipped += 1
        if isinstance(error, BatchProcessingFailure):
            audit_error = str(error.cause) if error.expected else "erro_interno"
            self._audit(
                "arquivo_ignorado",
                {"arquivo": error.input_path, "erro": audit_error},
            )
            if error.expected:
                message = str(error.cause)
                title = "Não foi possível processar o arquivo"
            else:
                message = (
                    f"Não foi possível processar o arquivo "
                    f"{nome_curto(error.input_path)}.\n\n"
                    "Verifique se ele não está corrompido ou aberto em outro programa."
                )
                title = "Erro no processamento"
        else:
            logger.exception(
                "Falha inesperada fora do processamento de arquivo",
                exc_info=(type(error), error, error.__traceback__),
            )
            message = "O processamento foi interrompido por um erro inesperado."
            title = "Erro no processamento"

        self._expose_partial_results(accumulator)
        if accumulator.processed:
            saved = (
                "1 arquivo foi salvo antes da falha."
                if accumulator.processed == 1
                else f"{accumulator.processed} arquivos foram salvos antes da falha."
            )
            message = f"{saved}\n\n{message}"

        self._safe_callback(
            "atualizar o status de erro",
            lambda: self.view.atualizar_status(
                " ".join(message.splitlines()), "error"
            ),
        )
        self._safe_callback(
            "aplicar o estado de erro",
            lambda: self.view.definir_estado(
                EstadoInterface.ERRO, total_arquivos=len(plan.items)
            ),
        )
        self._safe_callback(
            "exibir a mensagem de erro",
            lambda: messagebox.showerror(title, message),
        )
        self.accumulator = None


__all__ = ["CsvWorkflow"]
