"""Janela principal e navegação em páginas do FAS Jornada."""

from __future__ import annotations

import ctypes
import os
import sys
import webbrowser
from pathlib import Path
from tkinter import Canvas, messagebox

import customtkinter as ctk
from PIL import Image, ImageOps, ImageTk

from app.controllers.main_controller import MainController
from app.core.config import (
    APP_GEOMETRY,
    APP_TITLE,
    BG_APP,
    MIN_HEIGHT,
    MIN_WIDTH,
)
from app.core.logger import logger
from app.core.version import APP_VERSION
from app.services.rhid_credentials_service import (
    CredentialStorageError,
    RhidCredentialService,
)
from app.ui.navigation import PaginaInterface
from app.ui.report_pages import CsvPage, HomePage, ProcessingPage, SuccessPage
from app.ui.rhid_page import (
    ETAPA_DOMINIO,
    ETAPA_ESCOPO,
    ETAPA_LOGIN,
    RHID_FORGOT_PASSWORD_URL,
    RhidPage,
)
from app.ui.view_state import EstadoInterface, obter_configuracao_interface


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


def resource_path(relative_path: str) -> str:
    """Resolve recursos tanto no código-fonte quanto no executável empacotado."""

    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    return str(base_path / relative_path)


class MainWindow:
    """Coordena as páginas sem misturar regras de relatório com widgets."""

    def __init__(self):
        self.root = ctk.CTk()
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "fas.jornada"
            )
        except Exception:
            pass

        self._configurar_janela()
        self.estado_interface = EstadoInterface.VAZIO
        self.pagina_atual = PaginaInterface.INICIO
        self._fluxo_atual: str | None = None
        self._valor_progresso = 0.0
        self._status_atual = "Aguardando uma origem de dados."
        self._status_tipo = "info"
        self._metricas = (0, "--:--", "--:--")
        self._pasta_saida = "Nenhuma pasta selecionada ainda."
        self._tempo_execucao = None
        self._abrir_arquivo_habilitado = False
        self._abrir_pasta_habilitado = False
        self._credencial_pendente: tuple[str, str, str, bool] | None = None
        self._erro_credencial: str | None = None
        self._imagens = []
        self._footer_source = None
        self._footer_photo = None

        self._credenciais = RhidCredentialService()
        self.controller = MainController(self)
        self._montar_layout()
        self._configurar_atalhos()
        self._carregar_credenciais_salvas()
        self.controller.iniciar()

    def _configurar_janela(self) -> None:
        icone = resource_path("app/assets/icon.ico")
        if os.path.exists(icone):
            try:
                self.root.iconbitmap(icone)
            except Exception:
                logger.debug("Não foi possível aplicar o ícone da janela.")
        self.root.title(APP_TITLE)
        self._aplicar_geometria_inicial()
        self.root.configure(fg_color=BG_APP)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.protocol("WM_DELETE_WINDOW", self._ao_tentar_fechar)

    def _aplicar_geometria_inicial(self) -> None:
        try:
            largura_texto, altura_texto = APP_GEOMETRY.split("+", 1)[0].split("x")
            largura_desejada = max(int(largura_texto), MIN_WIDTH)
            altura_desejada = max(int(altura_texto), MIN_HEIGHT)
        except (AttributeError, TypeError, ValueError):
            largura_desejada, altura_desejada = MIN_WIDTH, MIN_HEIGHT

        largura_tela = max(1, self.root.winfo_screenwidth())
        altura_tela = max(1, self.root.winfo_screenheight())
        escala = max(0.5, float(ctk.ScalingTracker.get_window_scaling(self.root)))
        # O CustomTkinter recebe geometria lógica e a converte para pixels.
        # Reservamos barra de tarefas/moldura antes dessa conversão.
        largura_disponivel = max(1, int((largura_tela - 32) / escala))
        altura_disponivel = max(1, int((altura_tela - 80) / escala))
        largura_minima = min(MIN_WIDTH, largura_disponivel)
        altura_minima = min(MIN_HEIGHT, altura_disponivel)
        self.root.minsize(largura_minima, altura_minima)

        largura = min(max(largura_desejada, largura_minima), largura_disponivel)
        altura = min(max(altura_desejada, altura_minima), altura_disponivel)
        largura_fisica = round(largura * escala)
        altura_fisica = round(altura * escala)
        posicao_x = max(0, (largura_tela - largura_fisica) // 2)
        posicao_y = max(0, (altura_tela - altura_fisica) // 2)
        self.root.geometry(f"{largura}x{altura}+{posicao_x}+{posicao_y}")

    def _montar_layout(self) -> None:
        self._montar_cabecalho()

        self.content_host = ctk.CTkFrame(
            self.root, fg_color=BG_APP, corner_radius=0
        )
        self.content_host.grid(row=1, column=0, sticky="nsew", padx=18)
        self.content_host.grid_rowconfigure(0, weight=1)
        self.content_host.grid_columnconfigure(0, weight=1)
        # Páginas roláveis não podem aumentar a geometria da janela e ocultar
        # o cabeçalho/rodapé; elas devem se adaptar ao espaço disponível.
        self.content_host.grid_propagate(False)

        self.home_page = HomePage(
            self.content_host,
            self._abrir_fluxo_csv,
            self._abrir_fluxo_rhid,
        )
        self.csv_page = CsvPage(
            self.content_host,
            self._voltar_para_inicio,
            self.controller.selecionar_arquivos,
            self.controller.limpar_selecao,
            self.controller.processar,
        )
        self.rhid_page = RhidPage(
            self.content_host,
            self._conectar_rhid,
            self.controller.gerar_relatorio_rhid,
            self._voltar_para_inicio,
            self._abrir_recuperacao_senha,
        )
        self.processing_page = ProcessingPage(
            self.content_host,
            self.controller.cancelar_processamento,
        )
        self.success_page = SuccessPage(
            self.content_host,
            self.controller.abrir_arquivo_gerado,
            self.controller.abrir_pasta_gerada,
            self._gerar_outro_relatorio,
            self._voltar_para_inicio,
        )
        self._paginas = {
            PaginaInterface.INICIO: self.home_page,
            PaginaInterface.CSV: self.csv_page,
            PaginaInterface.RHID_LOGIN: self.rhid_page,
            PaginaInterface.RHID_DOMINIO: self.rhid_page,
            PaginaInterface.RHID_ESCOPO: self.rhid_page,
            PaginaInterface.PROCESSAMENTO: self.processing_page,
            PaginaInterface.SUCESSO: self.success_page,
        }

        # Nomes públicos mantidos durante a migração da interface.
        self.btn_selecionar = self.csv_page.btn_selecionar
        self.btn_limpar = self.csv_page.btn_limpar
        self.btn_processar = self.csv_page.btn_processar
        self.btn_cancelar = self.processing_page.btn_cancelar
        self.btn_abrir = self.success_page.btn_abrir
        self.btn_abrir_pasta = self.success_page.btn_abrir_pasta
        self.combo_departamento = self.csv_page.combo_departamento
        self.var_saldo = self.csv_page.var_saldo
        self.var_resumo = self.csv_page.var_resumo
        self.var_ranking = self.csv_page.var_ranking
        self.check_saldo = self.csv_page.check_saldo
        self.check_resumo = self.csv_page.check_resumo
        self.check_ranking = self.csv_page.check_ranking
        self.checkbox_saldo = self.check_saldo
        self.checkbox_resumo = self.check_resumo
        self.checkbox_ranking = self.check_ranking
        self.progress = self.processing_page.progress
        self.label_status = self.csv_page.label_status
        self.label_arquivo = self.csv_page.label_arquivo

        self._montar_rodape()
        self.mostrar_pagina(PaginaInterface.INICIO)

    def _montar_cabecalho(self) -> None:
        cabecalho = ctk.CTkFrame(self.root, fg_color=BG_APP, corner_radius=0)
        cabecalho.grid(row=0, column=0, sticky="ew", padx=24, pady=(2, 0))
        cabecalho.grid_columnconfigure((0, 2), weight=1, uniform="cabecalho")

        logo_path = resource_path("app/assets/logo_white.png")
        if os.path.exists(logo_path):
            try:
                imagem = Image.open(logo_path).convert("RGBA")
                # O arquivo fornecido contém margem da própria captura. O
                # recorte mantém o logotipo legível em telas com escala alta.
                imagem = imagem.crop((43, 25, 227, 105))
                logo = ctk.CTkImage(
                    light_image=imagem,
                    dark_image=imagem,
                    size=(125, 54),
                )
                self._imagens.append(logo)
                ctk.CTkLabel(cabecalho, image=logo, text="").grid(
                    row=0, column=1
                )
            except Exception:
                logger.exception("Falha ao carregar o logotipo da aplicação.")
        else:
            ctk.CTkLabel(
                cabecalho,
                text="FAS JORNADA",
                text_color="#ffffff",
                font=("Segoe UI", 20, "bold"),
            ).grid(row=0, column=1, pady=20)

        self.label_versao = ctk.CTkLabel(
            cabecalho,
            text=f"Versão {APP_VERSION}",
            text_color="#dce7ec",
            font=("Segoe UI", 10, "bold"),
        )
        self.label_versao.grid(row=0, column=2, sticky="e", padx=(0, 10))

    def _montar_rodape(self) -> None:
        rodape = Canvas(
            self.root,
            height=68,
            background="#034a5c",
            borderwidth=0,
            highlightthickness=0,
        )
        rodape.grid(row=2, column=0, sticky="ew")
        footer_path = resource_path("app/assets/footer_pattern.png")
        if os.path.exists(footer_path):
            try:
                self._footer_source = Image.open(footer_path).convert("RGB")
                self._footer_canvas = rodape
                rodape.bind("<Configure>", self._redesenhar_rodape, add="+")
                return
            except Exception:
                logger.exception("Falha ao carregar a faixa visual da aplicação.")

    def _redesenhar_rodape(self, evento) -> None:
        if self._footer_source is None or evento.width < 2 or evento.height < 2:
            return
        imagem = ImageOps.fit(
            self._footer_source,
            (evento.width, evento.height),
            method=Image.Resampling.LANCZOS,
            centering=(0.5, 0.5),
        )
        self._footer_photo = ImageTk.PhotoImage(imagem)
        self._footer_canvas.delete("all")
        self._footer_canvas.create_image(
            0, 0, image=self._footer_photo, anchor="nw"
        )

    def mostrar_pagina(self, pagina: PaginaInterface) -> None:
        pagina = PaginaInterface(pagina)
        vistos = set()
        for widget in self._paginas.values():
            if id(widget) not in vistos:
                widget.grid_remove()
                vistos.add(id(widget))
        self._paginas[pagina].grid(row=0, column=0, sticky="nsew")
        self.pagina_atual = pagina

    def _abrir_fluxo_csv(self) -> None:
        if self.controller.processamento_em_andamento:
            return
        if self.estado_interface is EstadoInterface.CONCLUIDO:
            self.controller.limpar_selecao()
            self.csv_page.redefinir_opcoes()
        self._fluxo_atual = "csv"
        self.mostrar_pagina(PaginaInterface.CSV)

    def _abrir_fluxo_rhid(self) -> None:
        if self.controller.processamento_em_andamento:
            return
        if (
            self.estado_interface is EstadoInterface.CONCLUIDO
            and self._fluxo_atual == "rhid"
        ):
            self.rhid_page.preparar_novo_relatorio()
        self._fluxo_atual = "rhid"
        if self._erro_credencial:
            self.rhid_page.exibir_erro(self._erro_credencial)
            self._erro_credencial = None
        pagina = {
            ETAPA_LOGIN: PaginaInterface.RHID_LOGIN,
            ETAPA_DOMINIO: PaginaInterface.RHID_DOMINIO,
            ETAPA_ESCOPO: PaginaInterface.RHID_ESCOPO,
        }.get(self.rhid_page.etapa_atual, PaginaInterface.RHID_LOGIN)
        self.mostrar_pagina(pagina)

    def _voltar_para_inicio(self) -> None:
        if self.controller.processamento_em_andamento:
            return
        self.mostrar_pagina(PaginaInterface.INICIO)

    def _gerar_outro_relatorio(self) -> None:
        if self._fluxo_atual == "rhid":
            self.rhid_page.preparar_novo_relatorio()
            self.mostrar_pagina(PaginaInterface.RHID_ESCOPO)
            return

        self.controller.limpar_selecao()
        self.csv_page.redefinir_opcoes()
        self._fluxo_atual = "csv"
        self.mostrar_pagina(PaginaInterface.CSV)

    def _configurar_atalhos(self) -> None:
        self.root.bind_all("<Control-o>", self._atalho_selecionar, add="+")
        self.root.bind_all("<Control-O>", self._atalho_selecionar, add="+")
        self.root.bind_all("<Control-Return>", self._atalho_processar, add="+")

    @staticmethod
    def _esta_habilitado(widget) -> bool:
        return str(widget.cget("state")) != "disabled"

    def _atalho_selecionar(self, _evento=None):
        if not self.controller.processamento_em_andamento:
            self._abrir_fluxo_csv()
            if self._esta_habilitado(self.btn_selecionar):
                self.controller.selecionar_arquivos()
        return "break"

    def _atalho_processar(self, _evento=None):
        if self.pagina_atual is PaginaInterface.CSV and self._esta_habilitado(
            self.btn_processar
        ):
            self.csv_page._processar()
        return "break"

    def definir_estado(self, estado, total_arquivos=0) -> None:
        if not isinstance(estado, EstadoInterface):
            estado = EstadoInterface(estado)
        configuracao = obter_configuracao_interface(estado, total_arquivos)
        self.estado_interface = estado
        self.csv_page.definir_acoes(
            selecionar_habilitado=configuracao.selecionar_habilitado,
            limpar_habilitado=configuracao.limpar_habilitado,
            processar_habilitado=configuracao.processar_habilitado,
            configuracao_habilitada=configuracao.configuracao_habilitada,
            texto_selecionar=configuracao.texto_selecionar,
            texto_processar=configuracao.texto_processar,
        )

        if estado in (EstadoInterface.PROCESSANDO, EstadoInterface.CANCELANDO):
            origem = "RHiD" if self._fluxo_atual == "rhid" else "arquivo CSV"
            self.processing_page.definir_origem(origem)
            self.processing_page.parar_indeterminado()
            self.processing_page.atualizar_progresso(
                self._valor_progresso, self._status_atual
            )
            self.processing_page.definir_cancelamento_habilitado(
                estado is EstadoInterface.PROCESSANDO and self._fluxo_atual == "csv",
                configuracao.texto_cancelar,
            )
            self.mostrar_pagina(PaginaInterface.PROCESSAMENTO)
            return

        if estado is EstadoInterface.CONCLUIDO:
            self._atualizar_pagina_sucesso()
            self.mostrar_pagina(PaginaInterface.SUCESSO)
            return

        if estado in (EstadoInterface.ERRO, EstadoInterface.CANCELADO):
            if self._abrir_arquivo_habilitado or self._abrir_pasta_habilitado:
                self._atualizar_pagina_sucesso()
                tipo = "error" if estado is EstadoInterface.ERRO else "warning"
                self.success_page.atualizar_status(self._status_atual, tipo)
                self.mostrar_pagina(PaginaInterface.SUCESSO)
                return
            if self._fluxo_atual == "rhid":
                self.mostrar_pagina(PaginaInterface.RHID_ESCOPO)
            elif self._fluxo_atual == "csv":
                self.mostrar_pagina(PaginaInterface.CSV)

    def _atualizar_pagina_sucesso(self) -> None:
        caminho = self._pasta_saida
        resultado = getattr(self.controller, "ultimo_resultado", None)
        if isinstance(resultado, dict) and resultado.get("caminho_saida"):
            caminho = resultado["caminho_saida"]
        self.success_page.atualizar_resultado(
            *self._metricas,
            caminho,
            status="Relatório salvo com sucesso.",
        )
        self.success_page.definir_abertura_habilitada(
            self._abrir_arquivo_habilitado,
            self._abrir_pasta_habilitado,
        )

    def _ao_tentar_fechar(self) -> None:
        if self.controller.processamento_em_andamento:
            messagebox.showwarning(
                "Processamento em andamento",
                "Aguarde o processamento terminar antes de fechar a janela.",
                parent=self.root,
            )
            return
        self.root.destroy()

    def agendar_na_interface(self, atraso_ms, callback):
        return self.root.after(atraso_ms, callback)

    # --- Integração RHiD -------------------------------------------------
    def _abrir_dialogo_rhid(self) -> None:
        """Compatibilidade: o antigo diálogo agora é uma página da janela."""

        self._abrir_fluxo_rhid()

    def _conectar_rhid(self, email, senha, dominio="") -> None:
        _email, _senha, _dominio, lembrar = self.rhid_page.obter_credenciais_digitadas()
        self._credencial_pendente = (email, senha, dominio, lembrar)
        if not lembrar:
            try:
                self._credenciais.delete()
            except CredentialStorageError:
                logger.warning("Não foi possível remover a credencial RHiD anterior.")
        self.controller.conectar_rhid(email, senha, dominio)

    def _abrir_recuperacao_senha(self) -> None:
        try:
            aberto = webbrowser.open_new_tab(RHID_FORGOT_PASSWORD_URL)
            if aberto is False:
                raise OSError("O navegador não aceitou a solicitação.")
        except Exception:
            logger.exception("Não foi possível abrir a recuperação de senha do RHiD.")
            self.rhid_page.exibir_erro(
                "Não foi possível abrir o site do RHiD no navegador."
            )

    def _carregar_credenciais_salvas(self) -> None:
        try:
            credencial = self._credenciais.load()
        except CredentialStorageError as erro:
            logger.warning("Falha ao carregar credencial segura do RHiD: %s", erro)
            self._erro_credencial = (
                "Não foi possível carregar o acesso lembrado. Digite suas credenciais novamente."
            )
            return
        if credencial is not None:
            self.rhid_page.preencher_credenciais(
                credencial.email, credencial.password, credencial.domain
            )

    def _salvar_credencial_apos_login(self) -> None:
        pendente = self._credencial_pendente
        if pendente is None:
            return
        email, senha, dominio, lembrar = pendente
        self._credencial_pendente = None
        if not lembrar:
            return
        try:
            self._credenciais.save(email, senha, dominio)
        except (CredentialStorageError, ValueError) as erro:
            logger.warning("Falha ao salvar credencial segura do RHiD: %s", erro)
            self.rhid_page.exibir_erro(
                "Conectado, mas não foi possível lembrar o acesso com segurança."
            )

    def definir_conexao_rhid_ocupada(self, ocupado):
        self.rhid_page.definir_ocupado(bool(ocupado))

    def exibir_catalogo_rhid(self, empresas, departamentos):
        self.rhid_page.exibir_catalogo(empresas, departamentos)
        self._salvar_credencial_apos_login()
        self._fluxo_atual = "rhid"
        self.mostrar_pagina(PaginaInterface.RHID_ESCOPO)

    def exibir_dominios_rhid(self, tenants):
        self.rhid_page.exibir_dominios(tenants)
        pagina = (
            PaginaInterface.RHID_DOMINIO
            if self.rhid_page.etapa_atual == ETAPA_DOMINIO
            else PaginaInterface.RHID_LOGIN
        )
        self.mostrar_pagina(pagina)

    def definir_geracao_rhid_ocupada(self, ocupado):
        self.rhid_page.definir_geracao_ocupada(bool(ocupado))

    def atualizar_progresso_rhid(self, valor, mensagem=""):
        self.rhid_page.atualizar_progresso(valor, mensagem)
        self.processing_page.atualizar_progresso(valor, mensagem)

    def exibir_sucesso_rhid(self, mensagem):
        self.rhid_page.exibir_sucesso_geracao(mensagem)

    def exibir_erro_rhid(self, mensagem):
        self.rhid_page.exibir_erro(mensagem)

    # --- Fachada usada pelo controller ----------------------------------
    def atualizar_departamentos(self, departamentos, selecionado="Todos"):
        self.csv_page.atualizar_departamentos(departamentos, selecionado)

    def atualizar_arquivo(self, texto):
        self.csv_page.atualizar_arquivo(texto)

    def atualizar_pasta_saida(self, texto):
        self._pasta_saida = str(texto)
        self.success_page.atualizar_caminho(texto)

    def atualizar_metricas(self, funcionarios, banco_total, banco_saldo):
        self._metricas = (funcionarios, banco_total, banco_saldo)
        self.success_page.atualizar_metricas(*self._metricas)

    def atualizar_status(self, texto, tipo="info"):
        self._status_atual = str(texto)
        self._status_tipo = tipo
        if self.pagina_atual is PaginaInterface.PROCESSAMENTO:
            self.processing_page.atualizar_status(texto, tipo)
        elif self.pagina_atual is PaginaInterface.SUCESSO:
            self.success_page.atualizar_status(texto, tipo)
        elif self._fluxo_atual == "csv":
            self.csv_page.atualizar_status(texto, tipo)
        self.root.update_idletasks()

    def atualizar_progresso(self, valor):
        try:
            self._valor_progresso = max(0.0, min(1.0, float(valor)))
        except (TypeError, ValueError):
            self._valor_progresso = 0.0
        self.processing_page.atualizar_progresso(self._valor_progresso)
        self.root.update_idletasks()

    def atualizar_tempo_execucao(self, segundos):
        self._tempo_execucao = segundos

    def atualizar_versao(self):
        self.label_versao.configure(text=f"Versão {APP_VERSION}")

    def habilitar_botao_abrir(self, habilitar):
        self._abrir_arquivo_habilitado = bool(habilitar)
        self.success_page.definir_abertura_habilitada(
            self._abrir_arquivo_habilitado,
            self._abrir_pasta_habilitado,
        )

    def habilitar_botao_abrir_pasta(self, habilitar):
        self._abrir_pasta_habilitado = bool(habilitar)
        self.success_page.definir_abertura_habilitada(
            self._abrir_arquivo_habilitado,
            self._abrir_pasta_habilitado,
        )

    @staticmethod
    def renderizar_historico(_itens):
        """O histórico continua persistido, mas não ocupa mais a interface."""

    def run(self):
        self.root.mainloop()


def iniciar_app():
    MainWindow().run()
