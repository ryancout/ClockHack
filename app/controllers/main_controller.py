import os
import time
from tkinter import filedialog, messagebox

from app.core.exceptions import AppError, SobrescritaCanceladaError
from app.core.logger import logger
from app.services.audit_service import registrar_evento
from app.services.file_service import garantir_extensao_xlsx, nome_curto, sugerir_nome_saida
from app.services.history_service import registrar_historico, ultimos_processamentos
from app.services.preferences_service import carregar_preferencias, salvar_preferencias
from app.services.validator_service import validar_arquivo_entrada
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
        ultimo_depto = self.preferencias.get("last_department") or "Todos"
        self.view.atualizar_departamentos(["Todos"], selecionado=ultimo_depto)
        self.view.atualizar_pasta_saida(self.preferencias.get("last_save_dir") or "Nenhuma pasta selecionada ainda.")
        self.view.atualizar_versao()

    def limpar_selecao(self):
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

    def selecionar_arquivos(self):
        try:
            initial_dir = self.preferencias.get("last_open_dir")
            if not initial_dir or not os.path.exists(initial_dir):
                initial_dir = os.path.expanduser("~")

            caminhos = filedialog.askopenfilenames(
                title="Selecione o(s) arquivo(s)",
                filetypes=[("Arquivos Excel e CSV", "*.xlsx *.csv"), ("Arquivos Excel", "*.xlsx"), ("Arquivos CSV", "*.csv")],
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
            self.preferencias["last_open_dir"] = os.path.dirname(self.arquivos_selecionados[0])
            salvar_preferencias(self.preferencias)

            total = len(self.arquivos_selecionados)
            nomes = [nome_curto(x) for x in self.arquivos_selecionados[:3]]
            texto = f"{total} arquivo(s) selecionado(s): " + ", ".join(nomes)
            if total > 3:
                texto += " ..."

            self.view.atualizar_arquivo(texto)
            self.view.atualizar_progresso(0)
            self.view.atualizar_tempo_execucao(None)
            self.view.habilitar_botao_abrir(False)
            self.view.habilitar_botao_abrir_pasta(False)

            departamentos = obter_departamentos(self.arquivos_selecionados[0])
            selecionado = self.preferencias.get("last_department") or "Todos"
            self.view.atualizar_departamentos(departamentos, selecionado=selecionado)
            self.view.atualizar_status("Arquivos carregados. Ajuste as opções e clique em Processar.", "info")

            logger.info("Arquivos selecionados: %s", [nome_curto(x) for x in self.arquivos_selecionados])
            registrar_evento(
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

    def processar(self, departamento, gerar_resumo=True, gerar_ranking=True):
        if not self.arquivos_selecionados:
            messagebox.showwarning("Aviso", "Selecione um ou mais arquivos primeiro.")
            return

        total_arquivos = len(self.arquivos_selecionados)
        caminho_saida_unico = None

        try:
            if total_arquivos == 1:
                nome_padrao = sugerir_nome_saida(self.arquivos_selecionados[0], departamento)
                caminho_saida_unico = filedialog.asksaveasfilename(
                    title="Salvar arquivo tratado como",
                    defaultextension=".xlsx",
                    filetypes=[("Arquivo Excel", "*.xlsx")],
                    initialdir=self.preferencias.get("last_save_dir") or self.preferencias.get("last_open_dir") or None,
                    initialfile=nome_padrao,
                )
                if not caminho_saida_unico:
                    self.view.atualizar_status("Operação cancelada pelo usuário.", "warning")
                    return
                caminho_saida_unico = garantir_extensao_xlsx(caminho_saida_unico)
                self._confirmar_sobrescrita(caminho_saida_unico)
                pasta_saida = os.path.dirname(caminho_saida_unico)
            else:
                pasta_saida = filedialog.askdirectory(
                    title="Selecione a pasta onde os arquivos tratados serão salvos",
                    initialdir=self.preferencias.get("last_save_dir") or self.preferencias.get("last_open_dir") or None,
                )
                if not pasta_saida:
                    self.view.atualizar_status("Operação cancelada pelo usuário.", "warning")
                    return
        except SobrescritaCanceladaError as e:
            self.view.atualizar_status(str(e), "warning")
            return

        self.preferencias["last_save_dir"] = pasta_saida
        self.preferencias["last_department"] = departamento or "Todos"
        salvar_preferencias(self.preferencias)
        self.view.atualizar_pasta_saida(pasta_saida)
        self.view.atualizar_status("Processando arquivo(s)...", "info")
        self.view.atualizar_progresso(0.02)
        self.view.habilitar_botao_abrir(False)
        self.view.habilitar_botao_abrir_pasta(False)

        total_funcionarios = 0
        total_bt_min = 0
        total_bs_min = 0
        processados = 0
        ignorados = 0
        ultimo_resultado = None
        inicio = time.perf_counter()

        for idx, arquivo in enumerate(self.arquivos_selecionados, start=1):
            progresso = 0.05 + ((idx - 1) / total_arquivos) * 0.9
            self.view.atualizar_progresso(progresso)
            self.view.atualizar_status(f"Processando {idx}/{total_arquivos}: {nome_curto(arquivo)}", "info")
            try:
                if total_arquivos == 1 and caminho_saida_unico:
                    caminho_saida = caminho_saida_unico
                else:
                    nome_saida = sugerir_nome_saida(arquivo, departamento)
                    caminho_saida = os.path.join(pasta_saida, nome_saida)
                    self._confirmar_sobrescrita(caminho_saida)

                resultado = processar_arquivo(
                    arquivo,
                    caminho_saida,
                    departamento,
                    gerar_ranking=gerar_ranking,
                    gerar_resumo=gerar_resumo,
                )

                total_funcionarios += resultado["quantidade_funcionarios"]
                total_bt_min += para_minutos(resultado["banco_total"])
                total_bs_min += para_minutos(resultado["banco_saldo"])
                processados += 1
                ultimo_resultado = resultado

                registrar_historico(
                    {
                        "arquivo_origem": arquivo,
                        "arquivo_saida": resultado["caminho_saida"],
                        "tipo_entrada": resultado["tipo_entrada"],
                        "quantidade_funcionarios": resultado["quantidade_funcionarios"],
                        "banco_total": resultado["banco_total"],
                        "banco_saldo": resultado["banco_saldo"],
                        "departamento": resultado["departamento"],
                        "gerou_resumo": resultado["gerou_resumo"],
                        "gerou_ranking": resultado["gerou_ranking"],
                    }
                )
            except SobrescritaCanceladaError as e:
                logger.info("Processamento interrompido por sobrescrita cancelada: %s", e)
                self.view.atualizar_status(str(e), "warning")
                return
            except AppError as e:
                ignorados += 1
                logger.warning("Falha de validação/processamento em %s: %s", nome_curto(arquivo), e)
                registrar_evento("arquivo_ignorado", {"arquivo": arquivo, "erro": str(e)})
                self.view.atualizar_status(str(e), "error")
                messagebox.showerror("Não foi possível processar o arquivo", str(e))
                return
            except Exception:
                ignorados += 1
                logger.exception("Erro inesperado no processamento de %s", nome_curto(arquivo))
                registrar_evento("arquivo_ignorado", {"arquivo": arquivo, "erro": "erro_interno"})
                self.view.atualizar_status(
                    f"Falha ao processar {nome_curto(arquivo)}. Verifique se o arquivo está íntegro e fechado.",
                    "error",
                )
                messagebox.showerror(
                    "Erro no processamento",
                    f"Não foi possível processar o arquivo {nome_curto(arquivo)}.\n\nVerifique se ele não está corrompido ou aberto em outro programa.",
                )
                return

        self.view.atualizar_progresso(1.0)

        if processados == 0:
            self.view.atualizar_status("Nenhum arquivo foi processado com sucesso.", "error")
            self.view.atualizar_tempo_execucao(None)
            messagebox.showerror("Erro", "Nenhum arquivo foi processado com sucesso.")
            return

        tempo_total = time.perf_counter() - inicio
        self.ultimo_resultado = ultimo_resultado
        self.view.habilitar_botao_abrir(True)
        self.view.habilitar_botao_abrir_pasta(True)
        self.view.atualizar_metricas(total_funcionarios, formatar_horas(total_bt_min), formatar_horas(total_bs_min))
        self.view.atualizar_tempo_execucao(tempo_total)
        self.view.atualizar_status(
            f"Processamento concluído. Processados: {processados} | Ignorados: {ignorados} | Filtro: {departamento}",
            "success",
        )
        self.view.renderizar_historico(ultimos_processamentos())

        registrar_evento(
            "processamento_lote",
            {
                "processados": processados,
                "ignorados": ignorados,
                "departamento": departamento,
                "pasta_saida": pasta_saida,
                "gerou_resumo": gerar_resumo,
                "gerou_ranking": gerar_ranking,
                "tempo_execucao_segundos": round(tempo_total, 2),
            },
        )

        messagebox.showinfo(
            "Sucesso",
            f"Processamento concluído.\n\n"
            f"Arquivos processados: {processados}\n"
            f"Arquivos ignorados: {ignorados}\n"
            f"Filtro aplicado: {departamento}\n"
            f"Funcionários: {total_funcionarios}\n"
            f"Banco Total: {formatar_horas(total_bt_min)}\n"
            f"Banco Saldo: {formatar_horas(total_bs_min)}\n"
            f"Tempo de execução: {tempo_total:.1f}s\n"
            f"Resumo: {'Sim' if gerar_resumo else 'Não'}\n"
            f"Ranking: {'Sim' if gerar_ranking else 'Não'}\n\n"
            f"Pasta de saída:\n{pasta_saida}",
        )

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
