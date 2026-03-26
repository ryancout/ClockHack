import os
from tkinter import filedialog, messagebox

from app.core.config import EXTENSOES_ACEITAS
from app.core.logger import logger
from app.services.audit_service import registrar_evento
from app.services.file_service import nome_curto, sugerir_nome_saida
from app.services.history_service import registrar_historico, ultimos_processamentos
from app.services.preferences_service import carregar_preferencias, salvar_preferencias
from app.services.workbook_pipeline_service import obter_departamentos, processar_arquivo
from app.services.time_service import para_minutos, formatar_horas

class MainController:
    def __init__(self, view):
        self.view = view
        self.arquivos_selecionados = []
        self.ultimo_resultado = None
        self.preferencias = carregar_preferencias()

    def iniciar(self):
        self.view.renderizar_historico(ultimos_processamentos())

    def selecionar_arquivos(self):
        try:
            initial_dir = self.preferencias.get("last_open_dir")

            if not initial_dir or not os.path.exists(initial_dir):
                initial_dir = os.path.expanduser("~")

            caminhos = filedialog.askopenfilenames(
                title="Selecione o(s) arquivo(s)",
                filetypes=EXTENSOES_ACEITAS,
                initialdir=initial_dir
            )

            if not caminhos:
                self.view.atualizar_status("Nenhum arquivo selecionado.", "warning")
                return

            self.arquivos_selecionados = list(caminhos)
            self.preferencias["last_open_dir"] = os.path.dirname(self.arquivos_selecionados[0])
            salvar_preferencias(self.preferencias)

            total = len(self.arquivos_selecionados)
            nomes = [nome_curto(x) for x in self.arquivos_selecionados[:3]]
            texto = f"{total} arquivo(s) selecionado(s): " + ", ".join(nomes)
            if total > 3:
                texto += " ..."

            self.view.atualizar_arquivo(texto)

            try:
                departamentos = obter_departamentos(self.arquivos_selecionados[0])
                self.view.atualizar_departamentos(departamentos)
            except Exception as e:
                self.view.atualizar_status(f"Arquivo selecionado, mas houve erro ao ler os departamentos: {e}", "error")
                logger.exception("Erro ao obter departamentos")
                return

            self.view.atualizar_status("Arquivos carregados. Escolha o departamento e processe.", "info")

            logger.info(f"Arquivos selecionados: {self.arquivos_selecionados}")
            registrar_evento("arquivos_selecionados", {"quantidade": total})

        except Exception as e:
            self.view.atualizar_status(f"Erro ao selecionar arquivos: {e}", "error")
            logger.exception("Erro em selecionar_arquivos")

    def processar(self, departamento):
        if not self.arquivos_selecionados:
            messagebox.showwarning("Aviso", "Selecione um ou mais arquivos primeiro.")
            return

        pasta_saida = filedialog.askdirectory(
            title="Selecione a pasta onde os arquivos tratados serão salvos",
            initialdir=self.preferencias.get("last_save_dir") or self.preferencias.get("last_open_dir") or None
        )
        if not pasta_saida:
            self.view.atualizar_status("Operação cancelada pelo usuário.", "warning")
            return

        self.preferencias["last_save_dir"] = pasta_saida
        salvar_preferencias(self.preferencias)

        total_arquivos = len(self.arquivos_selecionados)
        total_funcionarios = 0
        total_bt_min = 0
        total_bs_min = 0
        processados = 0
        ignorados = 0
        ultimo_resultado = None

        for idx, arquivo in enumerate(self.arquivos_selecionados, start=1):
            self.view.atualizar_progresso(idx / total_arquivos * 0.95)
            try:
                nome_saida = sugerir_nome_saida(arquivo, departamento)
                caminho_saida = os.path.join(pasta_saida, nome_saida)
                resultado = processar_arquivo(arquivo, caminho_saida, departamento)

                total_funcionarios += resultado["quantidade_funcionarios"]
                total_bt_min += para_minutos(resultado["banco_total"])
                total_bs_min += para_minutos(resultado["banco_saldo"])
                processados += 1
                ultimo_resultado = resultado

                registrar_historico({
                    "arquivo_origem": arquivo,
                    "arquivo_saida": resultado["caminho_saida"],
                    "tipo_entrada": resultado["tipo_entrada"],
                    "quantidade_funcionarios": resultado["quantidade_funcionarios"],
                    "banco_total": resultado["banco_total"],
                    "banco_saldo": resultado["banco_saldo"],
                    "departamento": resultado["departamento"],
                })

            except Exception as e:
                ignorados += 1
                erro_msg = f"Erro no arquivo {os.path.basename(arquivo)}: {e}"
                logger.exception(erro_msg)
                registrar_evento("arquivo_ignorado", {"arquivo": arquivo, "erro": str(e)})
                self.view.atualizar_status(erro_msg, "error")
                messagebox.showerror("Erro no processamento", erro_msg)
                return

        self.view.atualizar_progresso(1.0)

        if processados == 0:
            self.view.atualizar_status("Nenhum arquivo foi processado com sucesso.", "error")
            messagebox.showerror("Erro", "Nenhum arquivo foi processado com sucesso.")
            return

        self.ultimo_resultado = ultimo_resultado
        self.view.habilitar_botao_abrir(True)
        self.view.habilitar_botao_abrir_pasta(True)
        self.view.atualizar_metricas(total_funcionarios, formatar_horas(total_bt_min), formatar_horas(total_bs_min))
        self.view.atualizar_status(
            f"Processamento concluído. Processados: {processados} | Ignorados: {ignorados} | Filtro: {departamento}",
            "success"
        )
        self.view.renderizar_historico(ultimos_processamentos())

        registrar_evento("processamento_lote", {
            "processados": processados,
            "ignorados": ignorados,
            "departamento": departamento,
            "pasta_saida": pasta_saida
        })

        messagebox.showinfo(
            "Sucesso",
            f"Processamento concluído.\n\n"
            f"Arquivos processados: {processados}\n"
            f"Arquivos ignorados: {ignorados}\n"
            f"Filtro aplicado: {departamento}\n"
            f"Funcionários: {total_funcionarios}\n"
            f"Banco Total: {formatar_horas(total_bt_min)}\n"
            f"Banco Saldo: {formatar_horas(total_bs_min)}\n\n"
            f"Pasta de saída:\n{pasta_saida}"
        )

    def abrir_arquivo_gerado(self):
        if not self.ultimo_resultado:
            return
        caminho = self.ultimo_resultado["caminho_saida"]
        try:
            os.startfile(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{e}")

    def abrir_pasta_gerada(self):
        if not self.ultimo_resultado:
            return
        pasta = os.path.dirname(self.ultimo_resultado["caminho_saida"])
        try:
            os.startfile(pasta)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta:\n{e}")
