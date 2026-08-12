import os
import re
import time
from dataclasses import dataclass
from datetime import date
from tkinter import filedialog, messagebox

from app.core.config import TIPOS_ARQUIVO_ENTRADA
from app.core.exceptions import AppError, SobrescritaCanceladaError
from app.core.logger import logger
from app.integrations.rhid_client import RhidApiError, RhidClient, RhidTenantRequired
from app.integrations.rhid_report_service import RhidReportPlan, processar_relatorio_rhid
from app.services.audit_service import registrar_evento
from app.services.background_task_runner import BackgroundTaskRunner
from app.services.file_service import garantir_extensao_xlsx, nome_curto, sugerir_nome_saida
from app.services.history_service import registrar_historico, ultimos_processamentos
from app.services.preferences_service import carregar_preferencias, salvar_preferencias
from app.services.time_service import formatar_horas, para_minutos
from app.services.validator_service import validar_arquivo_entrada
from app.services.workbook_pipeline_service import obter_departamentos, processar_arquivo
from app.ui.view_state import EstadoInterface


@dataclass(frozen=True, slots=True)
class _ItemLote:
    caminho_entrada: str
    caminho_saida: str


@dataclass(frozen=True, slots=True)
class _PlanoProcessamento:
    itens: tuple[_ItemLote, ...]
    pasta_saida: str
    departamento: str
    gerar_saldo: bool
    gerar_resumo: bool
    gerar_ranking: bool


@dataclass(frozen=True, slots=True)
class _ArquivoIniciado:
    indice: int
    total: int
    item: _ItemLote


@dataclass(frozen=True, slots=True)
class _ArquivoConcluido:
    indice: int
    total: int
    item: _ItemLote
    resultado: dict


@dataclass(slots=True)
class _AcumuladorLote:
    total_funcionarios: int = 0
    total_bt_min: int = 0
    total_bs_min: int = 0
    processados: int = 0
    ignorados: int = 0
    ultimo_resultado: dict | None = None


class _FalhaProcessamentoLote(Exception):
    def __init__(self, caminho_entrada, causa, esperada):
        super().__init__(str(causa))
        self.caminho_entrada = caminho_entrada
        self.causa = causa
        self.esperada = esperada


class _CancelamentoSolicitado(Exception):
    pass


class MainController:
    def __init__(self, view, task_runner=None):
        self.view = view
        self.arquivos_selecionados = []
        self.ultimo_resultado = None
        self.preferencias = carregar_preferencias()
        self._processando = False
        self._acumulador_lote = None
        self._rhid_client = None
        self._rhid_empresas = ()
        self._rhid_departamentos = ()
        self._rhid_runner = None
        self._task_runner = (
            task_runner
            if task_runner is not None
            else BackgroundTaskRunner(self.view.agendar_na_interface)
        )

    @property
    def processamento_em_andamento(self):
        rhid_ativo = self._rhid_runner is not None and self._rhid_runner.ativo
        return self._processando or self._task_runner.ativo or rhid_ativo

    def iniciar(self):
        self.view.renderizar_historico(ultimos_processamentos())
        ultimo_depto = self.preferencias.get("last_department") or "Todos"
        self.view.atualizar_departamentos(["Todos"], selecionado=ultimo_depto)
        self.view.atualizar_pasta_saida(self.preferencias.get("last_save_dir") or "Nenhuma pasta selecionada ainda.")
        self.view.atualizar_versao()
        self.view.definir_estado(EstadoInterface.VAZIO)

    def conectar_rhid(self, email, senha, dominio=""):
        """Autentica e carrega o catálogo organizacional sem persistir a senha."""
        if self._rhid_runner is not None and self._rhid_runner.ativo:
            return

        if not email.strip() or not senha:
            self.view.exibir_erro_rhid("Informe o e-mail e a senha do RHiD.")
            return

        self.view.definir_conexao_rhid_ocupada(True)
        self._rhid_runner = BackgroundTaskRunner(self.view.agendar_na_interface)

        def carregar_catalogo(_reportar):
            cliente = RhidClient()
            cliente.login(email, senha, dominio)
            try:
                empresas = cliente.listar_empresas()
            except RhidApiError as erro:
                raise RhidApiError(
                    f"Login realizado, mas falhou ao carregar empresas: {erro}"
                ) from erro
            try:
                departamentos_cadastrados = cliente.listar_departamentos()
                departamentos = cliente.filtrar_departamentos_com_pessoas_ativas(
                    departamentos_cadastrados
                )
            except RhidApiError as erro:
                raise RhidApiError(
                    f"Login realizado, mas falhou ao carregar setores ativos: {erro}"
                ) from erro
            return cliente, empresas, departamentos

        def concluir(resultado):
            cliente, empresas, departamentos = resultado
            self._rhid_client = cliente
            self._rhid_empresas = tuple(empresas)
            self._rhid_departamentos = tuple(departamentos)
            self.view.definir_conexao_rhid_ocupada(False)
            self.view.exibir_catalogo_rhid(empresas, departamentos)
            self.view.atualizar_status("Conectado ao RHiD.", "success")

        def falhar(erro):
            logger.warning("Falha ao conectar ao RHiD: %s", erro)
            mensagem = str(erro) if isinstance(erro, RhidApiError) else "Não foi possível conectar ao RHiD."
            self.view.exibir_erro_rhid(mensagem)
            if isinstance(erro, RhidTenantRequired):
                self.view.exibir_dominios_rhid(erro.tenants)
            self.view.atualizar_status(mensagem, "error")

        self._rhid_runner.executar(carregar_catalogo, lambda _evento: None, concluir, falhar)

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
        """Gera e trata o CSV oficial do RHiD sem download manual."""
        if self.processamento_em_andamento:
            self.view.atualizar_status("Já existe um processamento em andamento.", "warning")
            return
        if self._rhid_client is None or not self._rhid_client.autenticado:
            self.view.exibir_erro_rhid("Conecte-se ao RHiD antes de gerar o relatório.")
            return

        try:
            empresa_id = int(empresa_id) if empresa_id is not None else None
        except (TypeError, ValueError):
            self.view.exibir_erro_rhid("Selecione uma empresa válida do RHiD.")
            return

        try:
            if departamento_id is None:
                departamento_ids = ()
                filtro_departamento = None
            elif isinstance(departamento_id, (list, tuple)):
                departamento_ids = tuple(
                    dict.fromkeys(int(item) for item in departamento_id)
                )
                filtro_departamento = departamento_ids or None
            else:
                departamento_unico = int(departamento_id)
                departamento_ids = (departamento_unico,)
                filtro_departamento = departamento_unico
        except (TypeError, ValueError):
            self.view.exibir_erro_rhid("Selecione setores válidos do RHiD.")
            return

        try:
            inicio = date.fromisoformat(str(data_inicial))
            fim = date.fromisoformat(str(data_final))
            if fim < inicio:
                raise RhidApiError("A data final não pode ser anterior à data inicial.")

            empresa = None
            if empresa_id is not None:
                empresa = next(
                    (item for item in self._rhid_empresas if item.id == empresa_id),
                    None,
                )
                if empresa is None:
                    raise RhidApiError("Selecione uma empresa válida do RHiD.")
            departamentos = []
            for item_id in departamento_ids:
                departamento = next(
                    (
                        item
                        for item in self._rhid_departamentos
                        if item.id == item_id
                        and (
                            empresa_id is None
                            or item.company_id in (0, empresa_id)
                        )
                    ),
                    None,
                )
                if departamento is None:
                    raise RhidApiError("Selecione um setor válido para essa empresa.")
                departamentos.append(departamento)
        except ValueError:
            self.view.exibir_erro_rhid(
                "Informe datas válidas no formato DD/MM/AAAA."
            )
            return
        except RhidApiError as erro:
            self.view.exibir_erro_rhid(str(erro))
            return

        rotulo_empresa = empresa.label if empresa is not None else "Todas as empresas"
        if len(departamentos) == 1:
            rotulo_departamento = departamentos[0].name
        elif departamentos:
            rotulo_departamento = f"{len(departamentos)} setores selecionados"
        else:
            rotulo_departamento = "Todos os setores"

        nome_escopo = rotulo_departamento if departamentos else rotulo_empresa
        nome_seguro = re.sub(r'[^\w.-]+', "_", nome_escopo, flags=re.UNICODE).strip("_.")
        nome_padrao = (
            f"relatorio_rhid_{nome_seguro or empresa_id}_"
            f"{inicio:%Y%m%d}_{fim:%Y%m%d}.xlsx"
        )
        caminho_saida = filedialog.asksaveasfilename(
            title="Salvar relatório do RHiD como",
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialdir=self.preferencias.get("last_save_dir") or None,
            initialfile=nome_padrao,
        )
        if not caminho_saida:
            return
        caminho_saida = garantir_extensao_xlsx(caminho_saida)
        try:
            self._confirmar_sobrescrita(caminho_saida)
        except SobrescritaCanceladaError:
            return

        plano = RhidReportPlan(
            company_id=empresa_id,
            department_id=filtro_departamento,
            company_label=rotulo_empresa,
            department_label=rotulo_departamento,
            data_inicial=inicio,
            data_final=fim,
            caminho_saida=caminho_saida,
            gerar_saldo=gerar_saldo,
            gerar_resumo=gerar_resumo,
            gerar_ranking=gerar_ranking,
        )
        cliente = self._rhid_client
        self._rhid_runner = BackgroundTaskRunner(self.view.agendar_na_interface)
        self.view.definir_geracao_rhid_ocupada(True)
        self.view.atualizar_progresso_rhid(0, "Solicitando relatório ao RHiD...")
        self.view.atualizar_status("Gerando relatório diretamente no RHiD...", "info")
        self.view.atualizar_progresso(0.02)
        self.view.habilitar_botao_abrir(False)
        self.view.habilitar_botao_abrir_pasta(False)
        self.view.definir_estado(EstadoInterface.PROCESSANDO, total_arquivos=1)

        def executar(reportar):
            inicio_execucao = time.perf_counter()
            resultado = processar_relatorio_rhid(cliente, plano, reportar)
            return resultado, time.perf_counter() - inicio_execucao

        def progresso(valor):
            valor = max(0, min(int(valor), 100))
            mensagem = (
                "Tratando o CSV e montando o Excel..."
                if valor >= 80
                else f"RHiD processando o relatório: {valor}%"
            )
            self.view.atualizar_progresso_rhid(valor, mensagem)
            self.view.atualizar_progresso(max(0.02, valor / 100))

        def concluir(resultado_com_tempo):
            resultado, tempo_total = resultado_com_tempo
            self.ultimo_resultado = resultado
            self.preferencias["last_save_dir"] = os.path.dirname(caminho_saida)
            salvar_preferencias(self.preferencias)
            self.view.definir_geracao_rhid_ocupada(False)
            self.view.atualizar_progresso_rhid(100, "Excel gerado e salvo.")
            self.view.exibir_sucesso_rhid(f"Salvo em: {caminho_saida}")
            self.view.atualizar_arquivo(
                f"RHiD: {rotulo_empresa} / {rotulo_departamento} / "
                f"{inicio:%d/%m/%Y} a {fim:%d/%m/%Y}"
            )
            self.view.atualizar_pasta_saida(os.path.dirname(caminho_saida))
            self.view.atualizar_metricas(
                resultado["quantidade_funcionarios"],
                resultado["banco_total"],
                resultado["banco_saldo"],
            )
            self.view.atualizar_tempo_execucao(tempo_total)
            self.view.atualizar_progresso(1.0)
            self.view.habilitar_botao_abrir(True)
            self.view.habilitar_botao_abrir_pasta(True)
            self.view.definir_estado(EstadoInterface.CONCLUIDO, total_arquivos=1)
            self.view.atualizar_status("Relatório do RHiD processado e salvo.", "success")
            self._registrar_historico_seguro(f"RHiD — {rotulo_empresa}", resultado)
            self.view.renderizar_historico(ultimos_processamentos())
            self._registrar_evento_seguro(
                "processamento_rhid",
                {
                    "empresa_id": empresa_id,
                    "departamento_ids": list(departamento_ids),
                    "periodo_inicial": inicio.isoformat(),
                    "periodo_final": fim.isoformat(),
                    "caminho_saida": caminho_saida,
                },
            )

        def falhar(erro):
            logger.warning("Falha ao gerar relatório do RHiD: %s", erro)
            mensagem = (
                str(erro)
                if isinstance(erro, (RhidApiError, AppError))
                else "Não foi possível gerar o relatório do RHiD."
            )
            self.view.definir_geracao_rhid_ocupada(False)
            self.view.exibir_erro_rhid(mensagem)
            self.view.atualizar_status(mensagem, "error")
            self.view.definir_estado(EstadoInterface.ERRO, total_arquivos=1)

        self._rhid_runner.executar(executar, progresso, concluir, falhar)

    def limpar_selecao(self):
        if self.processamento_em_andamento:
            self.view.atualizar_status(
                "Aguarde o processamento terminar para limpar a seleção.",
                "warning",
            )
            return

        self.arquivos_selecionados = []
        self.ultimo_resultado = None
        self.view.atualizar_arquivo("Nenhum arquivo selecionado")
        self.view.atualizar_departamentos(["Todos"], selecionado="Todos")
        self.view.atualizar_metricas(0, "--:--", "--:--")
        self.view.atualizar_progresso(0)
        self.view.atualizar_status("Seleção limpa. Escolha um novo arquivo para continuar.", "info")
        self.view.atualizar_tempo_execucao(None)
        self.view.habilitar_botao_abrir(False)
        self.view.habilitar_botao_abrir_pasta(False)
        self.view.definir_estado(EstadoInterface.VAZIO)

    def selecionar_arquivos(self):
        if self.processamento_em_andamento:
            self.view.atualizar_status(
                "Aguarde o processamento terminar para selecionar novos arquivos.",
                "warning",
            )
            return

        try:
            initial_dir = self.preferencias.get("last_open_dir")
            if not initial_dir or not os.path.exists(initial_dir):
                initial_dir = os.path.expanduser("~")

            caminhos = filedialog.askopenfilenames(
                title="Selecione o(s) arquivo(s) CSV",
                filetypes=TIPOS_ARQUIVO_ENTRADA,
                initialdir=initial_dir,
            )
            if not caminhos:
                self.view.atualizar_status("Nenhum arquivo selecionado.", "warning")
                return

            arquivos_validos = []
            for caminho in caminhos:
                validar_arquivo_entrada(caminho)
                arquivos_validos.append(caminho)

            self.arquivos_selecionados = arquivos_validos
            self.ultimo_resultado = None
            self.preferencias["last_open_dir"] = os.path.dirname(self.arquivos_selecionados[0])
            salvar_preferencias(self.preferencias)

            total = len(self.arquivos_selecionados)
            nomes = [nome_curto(x) for x in self.arquivos_selecionados[:3]]
            rotulo_selecao = (
                "1 arquivo selecionado"
                if total == 1
                else f"{total} arquivos selecionados"
            )
            texto = f"{rotulo_selecao}: " + ", ".join(nomes)
            if total > 3:
                texto += " ..."

            self.view.atualizar_arquivo(texto)
            self.view.atualizar_metricas(0, "--:--", "--:--")
            self.view.atualizar_progresso(0)
            self.view.atualizar_tempo_execucao(None)
            self.view.habilitar_botao_abrir(False)
            self.view.habilitar_botao_abrir_pasta(False)

            departamentos = obter_departamentos(self.arquivos_selecionados[0])
            selecionado = self.preferencias.get("last_department") or "Todos"
            self.view.atualizar_departamentos(departamentos, selecionado=selecionado)
            mensagem_carregamento = (
                "Arquivo carregado" if total == 1 else "Arquivos carregados"
            )
            self.view.atualizar_status(
                f"{mensagem_carregamento}. Ajuste as opções e clique em Processar.",
                "info",
            )
            self.view.definir_estado(EstadoInterface.PRONTO, total_arquivos=total)

            logger.info("Arquivos selecionados: %s", [nome_curto(x) for x in self.arquivos_selecionados])
            self._registrar_evento_seguro(
                "arquivos_selecionados",
                {"quantidade": total, "arquivos": [nome_curto(x) for x in self.arquivos_selecionados]},
            )
        except AppError as e:
            logger.warning("Validação ao selecionar arquivos: %s", e)
            self.view.atualizar_status(str(e), "error")
            messagebox.showerror("Arquivo inválido", str(e))
        except Exception:
            logger.exception("Erro em selecionar_arquivos")
            self.view.atualizar_status("Não foi possível carregar os arquivos selecionados.", "error")
            messagebox.showerror(
                "Erro ao selecionar arquivos",
                "Não foi possível carregar os arquivos selecionados. Verifique se eles estão fechados e tente novamente.",
            )

    def _confirmar_sobrescrita(self, caminho_saida):
        if os.path.exists(caminho_saida):
            confirmar = messagebox.askyesno(
                "Confirmar substituição",
                f"Já existe um arquivo com este nome:\n\n{nome_curto(caminho_saida)}\n\nDeseja substituir?",
            )
            if not confirmar:
                raise SobrescritaCanceladaError("Gravação cancelada para evitar sobrescrita de arquivo existente.")

    @staticmethod
    def _validar_destinos_unicos(itens):
        destinos = {}
        for item in itens:
            destino_normalizado = os.path.normcase(os.path.abspath(item.caminho_saida))
            if destino_normalizado in destinos:
                primeiro = destinos[destino_normalizado]
                raise AppError(
                    "Dois arquivos do lote gerariam a mesma saída: "
                    f"{nome_curto(primeiro.caminho_entrada)} e "
                    f"{nome_curto(item.caminho_entrada)}. "
                    "Renomeie um dos arquivos de entrada e tente novamente."
                )
            destinos[destino_normalizado] = item

    def _preparar_plano(
        self,
        arquivos,
        departamento,
        gerar_saldo,
        gerar_resumo,
        gerar_ranking,
    ):
        total_arquivos = len(arquivos)

        if total_arquivos == 1:
            nome_padrao = sugerir_nome_saida(arquivos[0], departamento)
            caminho_saida = filedialog.asksaveasfilename(
                title="Salvar arquivo tratado como",
                defaultextension=".xlsx",
                filetypes=[("Arquivo Excel", "*.xlsx")],
                initialdir=self.preferencias.get("last_save_dir")
                or self.preferencias.get("last_open_dir")
                or None,
                initialfile=nome_padrao,
            )
            if not caminho_saida:
                return None

            caminho_saida = garantir_extensao_xlsx(caminho_saida)
            pasta_saida = os.path.dirname(caminho_saida)
            itens = (_ItemLote(arquivos[0], caminho_saida),)
        else:
            pasta_saida = filedialog.askdirectory(
                title="Selecione a pasta onde os arquivos tratados serão salvos",
                initialdir=self.preferencias.get("last_save_dir")
                or self.preferencias.get("last_open_dir")
                or None,
            )
            if not pasta_saida:
                return None

            itens = tuple(
                _ItemLote(
                    arquivo,
                    os.path.join(pasta_saida, sugerir_nome_saida(arquivo, departamento)),
                )
                for arquivo in arquivos
            )

        self._validar_destinos_unicos(itens)
        for item in itens:
            self._confirmar_sobrescrita(item.caminho_saida)

        return _PlanoProcessamento(
            itens=itens,
            pasta_saida=pasta_saida,
            departamento=departamento,
            gerar_saldo=gerar_saldo,
            gerar_resumo=gerar_resumo,
            gerar_ranking=gerar_ranking,
        )

    def processar(self, departamento, gerar_saldo=True, gerar_resumo=True, gerar_ranking=True):
        if self.processamento_em_andamento:
            self.view.atualizar_status(
                "Já existe um processamento em andamento.",
                "warning",
            )
            return

        if not self.arquivos_selecionados:
            messagebox.showwarning("Aviso", "Selecione um ou mais arquivos primeiro.")
            return

        arquivos = tuple(self.arquivos_selecionados)
        total_arquivos = len(arquivos)

        try:
            plano = self._preparar_plano(
                arquivos,
                departamento,
                gerar_saldo,
                gerar_resumo,
                gerar_ranking,
            )
            if plano is None:
                self.view.atualizar_status("Operação cancelada pelo usuário.", "warning")
                self.view.definir_estado(
                    EstadoInterface.PRONTO,
                    total_arquivos=total_arquivos,
                )
                return
        except SobrescritaCanceladaError as e:
            self.view.atualizar_status(str(e), "warning")
            self.view.definir_estado(EstadoInterface.PRONTO, total_arquivos=total_arquivos)
            return
        except AppError as e:
            self.view.atualizar_status(str(e), "error")
            self.view.definir_estado(EstadoInterface.PRONTO, total_arquivos=total_arquivos)
            messagebox.showerror("Não foi possível preparar o lote", str(e))
            return

        self.preferencias["last_save_dir"] = plano.pasta_saida
        self.preferencias["last_department"] = departamento or "Todos"
        salvar_preferencias(self.preferencias)
        self.view.atualizar_pasta_saida(plano.pasta_saida)
        self.view.atualizar_status("Processando arquivo(s)...", "info")
        self.view.atualizar_progresso(0.02)
        self.view.habilitar_botao_abrir(False)
        self.view.habilitar_botao_abrir_pasta(False)
        self.view.definir_estado(EstadoInterface.PROCESSANDO, total_arquivos=total_arquivos)
        self._processando = True
        self._acumulador_lote = _AcumuladorLote()

        try:
            iniciado = self._task_runner.executar(
                lambda reportar: self._executar_plano(plano, reportar),
                self._receber_evento_lote,
                lambda tempo_total: self._finalizar_lote(
                    plano,
                    tempo_total,
                ),
                lambda erro: self._falhar_lote(
                    plano,
                    erro,
                ),
            )
        except Exception as erro:
            self._falhar_lote(plano, erro)
            return

        if not iniciado:
            self._processando = False
            self._acumulador_lote = None
            self.view.atualizar_status("Já existe um processamento em andamento.", "warning")
            self.view.definir_estado(EstadoInterface.PRONTO, total_arquivos=total_arquivos)

    def cancelar_processamento(self):
        if not self._processando:
            return
        if not self._task_runner.cancelar():
            return

        self.view.atualizar_status(
            "Cancelando após concluir o arquivo atual...",
            "warning",
        )
        self.view.definir_estado(
            EstadoInterface.CANCELANDO,
            total_arquivos=len(self.arquivos_selecionados),
        )

    def _executar_plano(self, plano, reportar):
        inicio = time.perf_counter()
        total = len(plano.itens)

        for indice, item in enumerate(plano.itens, start=1):
            if self._task_runner.cancelamento_solicitado:
                raise _CancelamentoSolicitado()

            reportar(_ArquivoIniciado(indice, total, item))
            try:
                resultado = processar_arquivo(
                    item.caminho_entrada,
                    item.caminho_saida,
                    plano.departamento,
                    gerar_saldo=plano.gerar_saldo,
                    gerar_ranking=plano.gerar_ranking,
                    gerar_resumo=plano.gerar_resumo,
                )
            except AppError as e:
                logger.warning(
                    "Falha de validação/processamento em %s: %s",
                    nome_curto(item.caminho_entrada),
                    e,
                )
                raise _FalhaProcessamentoLote(
                    item.caminho_entrada,
                    e,
                    esperada=True,
                ) from e
            except Exception as e:
                logger.exception(
                    "Erro inesperado no processamento de %s",
                    nome_curto(item.caminho_entrada),
                )
                raise _FalhaProcessamentoLote(
                    item.caminho_entrada,
                    e,
                    esperada=False,
                ) from e

            reportar(_ArquivoConcluido(indice, total, item, resultado))

        if self._task_runner.cancelamento_solicitado:
            raise _CancelamentoSolicitado()

        return time.perf_counter() - inicio

    def _receber_evento_lote(self, evento):
        if isinstance(evento, _ArquivoIniciado):
            progresso = 0.05 + ((evento.indice - 1) / evento.total) * 0.9
            self.view.atualizar_progresso(progresso)
            self.view.atualizar_status(
                f"Processando {evento.indice}/{evento.total}: "
                f"{nome_curto(evento.item.caminho_entrada)}",
                "info",
            )
            return

        if not isinstance(evento, _ArquivoConcluido):
            return

        acumulador = self._acumulador_lote
        if acumulador is None:
            return

        resultado = evento.resultado
        acumulador.total_funcionarios += resultado["quantidade_funcionarios"]
        acumulador.total_bt_min += para_minutos(resultado["banco_total"])
        acumulador.total_bs_min += para_minutos(resultado["banco_saldo"])
        acumulador.processados += 1
        acumulador.ultimo_resultado = resultado
        progresso = 0.05 + (evento.indice / evento.total) * 0.9
        self.view.atualizar_progresso(progresso)
        self._registrar_historico_seguro(evento.item.caminho_entrada, resultado)

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

    @staticmethod
    def _executar_callback_seguro(descricao, callback):
        try:
            callback()
        except Exception:
            logger.exception("Falha ao %s", descricao)
            return False
        return True

    def _expor_resultados_parciais(self, acumulador):
        if not acumulador.processados:
            return

        self.ultimo_resultado = acumulador.ultimo_resultado
        self._executar_callback_seguro(
            "habilitar a abertura do arquivo salvo",
            lambda: self.view.habilitar_botao_abrir(True),
        )
        self._executar_callback_seguro(
            "habilitar a abertura da pasta de saída",
            lambda: self.view.habilitar_botao_abrir_pasta(True),
        )
        self._executar_callback_seguro(
            "atualizar as métricas parciais",
            lambda: self.view.atualizar_metricas(
                acumulador.total_funcionarios,
                formatar_horas(acumulador.total_bt_min),
                formatar_horas(acumulador.total_bs_min),
            ),
        )
        self._executar_callback_seguro(
            "atualizar o histórico após processamento parcial",
            lambda: self.view.renderizar_historico(ultimos_processamentos()),
        )

    @staticmethod
    def _resumo_resultados_parciais(acumulador, motivo):
        if not acumulador.processados:
            return motivo
        quantidade = acumulador.processados
        salvo = "1 arquivo foi salvo" if quantidade == 1 else f"{quantidade} arquivos foram salvos"
        return f"{motivo} {salvo} antes da interrupção."

    def _finalizar_cancelamento(self, plano, acumulador):
        self._processando = False
        self._expor_resultados_parciais(acumulador)
        mensagem = self._resumo_resultados_parciais(
            acumulador,
            "Processamento cancelado.",
        )
        self._executar_callback_seguro(
            "atualizar o status de cancelamento",
            lambda: self.view.atualizar_status(mensagem, "warning"),
        )
        self._executar_callback_seguro(
            "aplicar o estado cancelado",
            lambda: self.view.definir_estado(
                EstadoInterface.CANCELADO,
                total_arquivos=len(plano.itens),
            ),
        )
        self._acumulador_lote = None

    def _finalizar_lote(self, plano, tempo_total):
        acumulador = self._acumulador_lote or _AcumuladorLote()
        self._processando = False
        self.view.atualizar_progresso(1.0)

        if acumulador.processados == 0:
            self.view.atualizar_status("Nenhum arquivo foi processado com sucesso.", "error")
            self.view.atualizar_tempo_execucao(None)
            self.view.definir_estado(
                EstadoInterface.ERRO,
                total_arquivos=len(plano.itens),
            )
            messagebox.showerror("Erro", "Nenhum arquivo foi processado com sucesso.")
            self._acumulador_lote = None
            return

        self.ultimo_resultado = acumulador.ultimo_resultado
        self.view.habilitar_botao_abrir(True)
        self.view.habilitar_botao_abrir_pasta(True)
        self.view.definir_estado(
            EstadoInterface.CONCLUIDO,
            total_arquivos=len(plano.itens),
        )
        self.view.atualizar_metricas(
            acumulador.total_funcionarios,
            formatar_horas(acumulador.total_bt_min),
            formatar_horas(acumulador.total_bs_min),
        )
        self.view.atualizar_tempo_execucao(tempo_total)
        mensagem_salvamento = (
            "1 arquivo foi salvo"
            if acumulador.processados == 1
            else f"{acumulador.processados} arquivos foram salvos"
        )
        self.view.atualizar_status(
            f"Processamento concluído. {mensagem_salvamento}. "
            f"Ignorados: {acumulador.ignorados} | Filtro: {plano.departamento}",
            "success",
        )
        self.view.renderizar_historico(ultimos_processamentos())

        self._registrar_evento_seguro(
            "processamento_lote",
            {
                "processados": acumulador.processados,
                "ignorados": acumulador.ignorados,
                "departamento": plano.departamento,
                "pasta_saida": plano.pasta_saida,
                "gerou_saldo": plano.gerar_saldo,
                "gerou_resumo": plano.gerar_resumo,
                "gerou_ranking": plano.gerar_ranking,
                "tempo_execucao_segundos": round(tempo_total, 2),
            },
        )

        self._acumulador_lote = None

    def _falhar_lote(self, plano, erro):
        self._processando = False
        acumulador = self._acumulador_lote or _AcumuladorLote()
        if isinstance(erro, _CancelamentoSolicitado):
            self._finalizar_cancelamento(plano, acumulador)
            return

        acumulador.ignorados += 1
        total_arquivos = len(plano.itens)

        if isinstance(erro, _FalhaProcessamentoLote):
            arquivo = erro.caminho_entrada
            detalhe_auditoria = str(erro.causa) if erro.esperada else "erro_interno"
            self._registrar_evento_seguro(
                "arquivo_ignorado",
                {"arquivo": arquivo, "erro": detalhe_auditoria},
            )

            if erro.esperada:
                mensagem = str(erro.causa)
                titulo = "Não foi possível processar o arquivo"
            else:
                mensagem = (
                    f"Não foi possível processar o arquivo {nome_curto(arquivo)}.\n\n"
                    "Verifique se ele não está corrompido ou aberto em outro programa."
                )
                titulo = "Erro no processamento"
        else:
            logger.exception(
                "Falha inesperada fora do processamento de arquivo",
                exc_info=(type(erro), erro, erro.__traceback__),
            )
            mensagem = "O processamento foi interrompido por um erro inesperado."
            titulo = "Erro no processamento"

        self._expor_resultados_parciais(acumulador)
        if acumulador.processados:
            quantidade = acumulador.processados
            resumo = (
                "1 arquivo foi salvo antes da falha."
                if quantidade == 1
                else f"{quantidade} arquivos foram salvos antes da falha."
            )
            mensagem = (
                f"{resumo}\n\n{mensagem}"
            )

        self._executar_callback_seguro(
            "atualizar o status de erro",
            lambda: self.view.atualizar_status(
                " ".join(mensagem.splitlines()),
                "error",
            ),
        )
        self._executar_callback_seguro(
            "aplicar o estado de erro",
            lambda: self.view.definir_estado(
                EstadoInterface.ERRO,
                total_arquivos=total_arquivos,
            ),
        )
        self._executar_callback_seguro(
            "exibir a mensagem de erro",
            lambda: messagebox.showerror(titulo, mensagem),
        )
        self._acumulador_lote = None

    def abrir_arquivo_gerado(self):
        if not self.ultimo_resultado:
            return
        caminho = self.ultimo_resultado["caminho_saida"]
        try:
            os.startfile(caminho)
        except Exception:
            messagebox.showerror("Erro", "Não foi possível abrir o arquivo gerado.")

    def abrir_pasta_gerada(self):
        if not self.ultimo_resultado:
            return
        pasta = os.path.dirname(self.ultimo_resultado["caminho_saida"])
        try:
            os.startfile(pasta)
        except Exception:
            messagebox.showerror("Erro", "Não foi possível abrir a pasta de saída.")
