import os
import sys
import ctypes
from tkinter import messagebox

import customtkinter as ctk
from PIL import Image

from app.controllers.main_controller import MainController
from app.core.config import (
    APP_GEOMETRY,
    APP_TITLE,
    BG_APP,
    BG_BOX,
    BG_CARD,
    BORDER,
    ERROR,
    FG_MUTED,
    FG_TEXT,
    FG_TITLE,
    FONT_BUTTON,
    FONT_METRIC_TITLE,
    FONT_METRIC_VALUE,
    FONT_STATUS,
    FONT_SUBTITLE,
    FONT_TITLE,
    MIN_HEIGHT,
    MIN_WIDTH,
    PRIMARY,
    SUCCESS,
    WARNING,
)
from app.core.version import APP_VERSION
from app.ui.view_state import EstadoInterface, obter_configuracao_interface
from app.ui.rhid_dialog import RhidConnectionDialog

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class MainWindow:
    def __init__(self):
        self.root = ctk.CTk()

        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("fas.processador.planilhas")
        except Exception:
            pass

        icone = resource_path("app/assets/icon.ico")
        if os.path.exists(icone):
            try:
                self.root.iconbitmap(icone)
            except Exception:
                try:
                    self.root.wm_iconbitmap(icone)
                except Exception:
                    pass
        self.root.title(APP_TITLE)
        self.root.minsize(MIN_WIDTH, MIN_HEIGHT)
        self._aplicar_geometria_inicial()
        self.root.configure(fg_color=BG_APP)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.protocol("WM_DELETE_WINDOW", self._ao_tentar_fechar)

        self.controller = MainController(self)
        self.logo_ref = None
        self.estado_interface = EstadoInterface.VAZIO
        self._valor_progresso = 0.0
        self._rhid_dialog = None

        self._montar_layout()
        self._configurar_atalhos()
        self.controller.iniciar()

    def _aplicar_geometria_inicial(self):
        """Centraliza a janela e limita sua geometria inicial à tela disponível."""
        try:
            geometria_base = APP_GEOMETRY.split("+", 1)[0]
            largura_texto, altura_texto = geometria_base.lower().split("x", 1)
            largura_desejada = max(int(largura_texto), MIN_WIDTH)
            altura_desejada = max(int(altura_texto), MIN_HEIGHT)
        except (AttributeError, TypeError, ValueError):
            largura_desejada = MIN_WIDTH
            altura_desejada = MIN_HEIGHT

        largura_tela = max(1, self.root.winfo_screenwidth())
        altura_tela = max(1, self.root.winfo_screenheight())
        largura = min(largura_desejada, largura_tela)
        altura = min(altura_desejada, altura_tela)
        posicao_x = max(0, (largura_tela - largura) // 2)
        posicao_y = max(0, (altura_tela - altura) // 2)

        self.root.geometry(f"{largura}x{altura}+{posicao_x}+{posicao_y}")

    def _montar_layout(self):
        container = ctk.CTkFrame(self.root, fg_color=BG_APP, corner_radius=0)
        container.grid(row=0, column=0, sticky="nsew", padx=18, pady=18)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=3, minsize=560)
        container.grid_columnconfigure(1, weight=1, minsize=280)

        esquerdo = ctk.CTkScrollableFrame(
            container,
            fg_color=BG_APP,
            corner_radius=0,
            scrollbar_button_color=BORDER,
            scrollbar_button_hover_color=FG_MUTED,
        )
        esquerdo.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        esquerdo.grid_columnconfigure(0, weight=1)

        direito = ctk.CTkFrame(container, fg_color=BG_APP, corner_radius=0)
        direito.grid(row=0, column=1, sticky="nsew")

        self._montar_card_principal(esquerdo)
        self._montar_lateral(direito)

    def _montar_card_principal(self, parent):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=20, border_width=1, border_color=BORDER)
        card.grid(row=0, column=0, sticky="nsew")

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(20, 12))
        header.grid_columnconfigure(1, weight=1)

        logo_path = resource_path("app/assets/logo.png")
        if os.path.exists(logo_path):
            try:
                logo = ctk.CTkImage(light_image=Image.open(logo_path), size=(112, 57))
                ctk.CTkLabel(header, image=logo, text="").grid(
                    row=0,
                    column=0,
                    sticky="nw",
                    padx=(0, 18),
                )
                self.logo_ref = logo
            except Exception:
                pass

        textos_header = ctk.CTkFrame(header, fg_color="transparent")
        textos_header.grid(row=0, column=1, sticky="nsew")
        textos_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            textos_header,
            text="Processador de Planilhas",
            font=FONT_TITLE,
            text_color=FG_TITLE,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        label_subtitulo = ctk.CTkLabel(
            textos_header,
            text="Selecione arquivo(s) CSV, escolha o departamento, marque as abas desejadas, processe e salve a saída em Excel.",
            font=FONT_SUBTITLE,
            text_color=FG_MUTED,
            wraplength=360,
            justify="left",
            anchor="w",
        )
        label_subtitulo.grid(row=1, column=0, sticky="ew", pady=(2, 3))
        self._vincular_quebra_texto(label_subtitulo, textos_header, margem=4)

        self.label_versao = ctk.CTkLabel(
            textos_header,
            text="",
            font=("Segoe UI", 10, "bold"),
            text_color=PRIMARY,
            anchor="w",
        )
        self.label_versao.grid(row=2, column=0, sticky="w")

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.pack(fill="x", padx=24, pady=(0, 18))

        top_actions = ctk.CTkFrame(actions, fg_color="transparent")
        top_actions.pack(fill="x", pady=(0, 10))
        top_actions.grid_columnconfigure(0, weight=1)
        top_actions.grid_columnconfigure(1, weight=1)

        self.btn_selecionar = ctk.CTkButton(
            top_actions,
            text="Selecionar arquivo(s)",
            height=42,
            fg_color=PRIMARY,
            hover_color="#0955af",
            font=FONT_BUTTON,
            command=self.controller.selecionar_arquivos,
        )
        self.btn_selecionar.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.btn_limpar = ctk.CTkButton(
            top_actions,
            text="Limpar seleção",
            height=42,
            fg_color="#e9eef5",
            text_color=FG_TEXT,
            hover_color="#dde6f1",
            font=FONT_BUTTON,
            command=self.controller.limpar_selecao,
        )
        self.btn_limpar.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        self.btn_rhid = ctk.CTkButton(
            actions,
            text="Conectar ao RHiD",
            height=38,
            fg_color="transparent",
            border_width=1,
            border_color=PRIMARY,
            text_color=PRIMARY,
            hover_color="#e8f1fc",
            font=FONT_BUTTON,
            command=self._abrir_dialogo_rhid,
        )
        self.btn_rhid.pack(fill="x", pady=(0, 10))

        box_arquivo = ctk.CTkFrame(actions, fg_color=BG_BOX, corner_radius=14, border_width=1, border_color=BORDER)
        box_arquivo.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(
            box_arquivo,
            text="Arquivo selecionado",
            text_color=FG_MUTED,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=14, pady=(12, 2))
        self.label_arquivo = ctk.CTkLabel(
            box_arquivo,
            text="Nenhum arquivo selecionado",
            text_color=FG_TEXT,
            font=("Segoe UI", 12),
            wraplength=360,
            justify="left",
            anchor="w",
        )
        self.label_arquivo.pack(fill="x", padx=14, pady=(0, 12))
        self._vincular_quebra_texto(self.label_arquivo, box_arquivo, margem=28)

        filtro_box = ctk.CTkFrame(actions, fg_color=BG_BOX, corner_radius=12, border_width=1, border_color=BORDER)
        filtro_box.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(
            filtro_box,
            text="Filtro por nome do departamento",
            text_color=FG_MUTED,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=14, pady=(10, 4))

        self.combo_departamento = ctk.CTkComboBox(filtro_box, values=["Todos"], height=36)
        self.combo_departamento.pack(fill="x", padx=14, pady=(0, 12))
        self.combo_departamento.set("Todos")

        opcoes_box = ctk.CTkFrame(actions, fg_color=BG_BOX, corner_radius=12, border_width=1, border_color=BORDER)
        opcoes_box.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(
            opcoes_box,
            text="Abas adicionais",
            text_color=FG_MUTED,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=14, pady=(10, 4))

        checks = ctk.CTkFrame(opcoes_box, fg_color="transparent")
        checks.pack(fill="x", padx=12, pady=(0, 10))
        self.var_saldo = ctk.BooleanVar(value=True)
        self.var_resumo = ctk.BooleanVar(value=True)
        self.var_ranking = ctk.BooleanVar(value=True)
        self.check_saldo = ctk.CTkCheckBox(checks, text="Gerar aba SALDO", variable=self.var_saldo)
        self.check_resumo = ctk.CTkCheckBox(checks, text="Gerar aba RESUMO", variable=self.var_resumo)
        self.check_ranking = ctk.CTkCheckBox(checks, text="Gerar aba RANKING", variable=self.var_ranking)
        self.check_saldo.pack(anchor="w", pady=2)
        self.check_resumo.pack(anchor="w", pady=2)
        self.check_ranking.pack(anchor="w", pady=2)
        self.checkbox_saldo = self.check_saldo
        self.checkbox_resumo = self.check_resumo
        self.checkbox_ranking = self.check_ranking

        self.btn_processar = ctk.CTkButton(
            actions,
            text="Processar arquivo(s)",
            height=44,
            fg_color=SUCCESS,
            hover_color="#0d634d",
            font=FONT_BUTTON,
            command=self._processar_clicado,
        )
        self.btn_processar.pack(fill="x", pady=(0, 10))

        self.btn_cancelar = ctk.CTkButton(
            actions,
            text="Cancelar",
            height=36,
            fg_color="#f9e5e3",
            text_color=ERROR,
            hover_color="#f2d2cf",
            font=FONT_BUTTON,
            command=self.controller.cancelar_processamento,
            state="disabled",
        )
        self.btn_cancelar.pack(fill="x", pady=(0, 10))

        sec = ctk.CTkFrame(actions, fg_color="transparent")
        sec.pack(fill="x")

        self.btn_abrir = ctk.CTkButton(
            sec,
            text="Abrir arquivo",
            height=38,
            fg_color="#e9eef5",
            text_color=FG_TEXT,
            hover_color="#dde6f1",
            font=FONT_BUTTON,
            command=self.controller.abrir_arquivo_gerado,
            state="disabled",
        )
        self.btn_abrir.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_abrir_pasta = ctk.CTkButton(
            sec,
            text="Abrir pasta",
            height=38,
            fg_color="#e9eef5",
            text_color=FG_TEXT,
            hover_color="#dde6f1",
            font=FONT_BUTTON,
            command=self.controller.abrir_pasta_gerada,
            state="disabled",
        )
        self.btn_abrir_pasta.pack(side="left", fill="x", expand=True, padx=(5, 0))

        box_saida = ctk.CTkFrame(actions, fg_color=BG_BOX, corner_radius=14, border_width=1, border_color=BORDER)
        box_saida.pack(fill="x", pady=(10, 14))
        ctk.CTkLabel(box_saida, text="Pasta de saída", text_color=FG_MUTED, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        self.label_pasta_saida = ctk.CTkLabel(
            box_saida,
            text="Nenhuma pasta selecionada ainda.",
            text_color=FG_TEXT,
            font=("Segoe UI", 12),
            wraplength=360,
            justify="left",
            anchor="w",
        )
        self.label_pasta_saida.pack(fill="x", padx=14, pady=(0, 12))
        self._vincular_quebra_texto(self.label_pasta_saida, box_saida, margem=28)

        metricas = ctk.CTkFrame(actions, fg_color="transparent")
        metricas.pack(fill="x", pady=(0, 14))
        metricas.grid_columnconfigure(0, weight=1)
        metricas.grid_columnconfigure(1, weight=1)
        metricas.grid_columnconfigure(2, weight=1)

        self.metric_func = self._criar_box_metrica(metricas, "Funcionários", "0")
        self.metric_bt = self._criar_box_metrica(metricas, "Banco Total", "--:--")
        self.metric_bs = self._criar_box_metrica(metricas, "Banco Saldo", "--:--")

        self.metric_func.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.metric_bt.grid(row=0, column=1, sticky="nsew", padx=6)
        self.metric_bs.grid(row=0, column=2, sticky="nsew", padx=(6, 0))

        tempo_box = ctk.CTkFrame(actions, fg_color=BG_BOX, corner_radius=14, border_width=1, border_color=BORDER)
        tempo_box.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(tempo_box, text="Tempo de execução", text_color=FG_MUTED, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
        self.label_tempo = ctk.CTkLabel(tempo_box, text="--", text_color=FG_TITLE, font=("Segoe UI", 14, "bold"))
        self.label_tempo.pack(anchor="w", padx=14, pady=(0, 10))

        self.progress = ctk.CTkProgressBar(actions, height=12)
        self.progress.pack(fill="x", pady=(0, 10))
        self.progress.set(0)

        self.label_status = ctk.CTkLabel(
            actions,
            text="Informação: Aguardando arquivo.",
            text_color=FG_MUTED,
            font=FONT_STATUS,
            wraplength=360,
            justify="left",
            anchor="w",
        )
        self.label_status.pack(fill="x")
        self._vincular_quebra_texto(self.label_status, actions, margem=4)

        self.definir_estado(EstadoInterface.VAZIO, 0)

    def _configurar_atalhos(self):
        self.root.bind_all("<Control-o>", self._atalho_selecionar, add="+")
        self.root.bind_all("<Control-O>", self._atalho_selecionar, add="+")
        self.root.bind_all("<Control-Return>", self._atalho_processar, add="+")

    @staticmethod
    def _esta_habilitado(widget):
        return str(widget.cget("state")) != "disabled"

    @staticmethod
    def _vincular_quebra_texto(label, parent, margem=0, minimo=180):
        """Mantém textos longos dentro da largura realmente disponível."""

        def atualizar_quebra(evento):
            escala = ctk.ScalingTracker.get_widget_scaling(label)
            largura_logica = (evento.width / escala) - margem
            label.configure(wraplength=max(minimo, largura_logica))

        parent.bind("<Configure>", atualizar_quebra, add="+")

    def _atalho_selecionar(self, _evento=None):
        if self._esta_habilitado(self.btn_selecionar):
            self.controller.selecionar_arquivos()
        return "break"

    def _atalho_processar(self, _evento=None):
        if self._esta_habilitado(self.btn_processar):
            self._processar_clicado()
        return "break"

    def definir_estado(self, estado, total_arquivos=0):
        """Aplica aos widgets a configuração visual do estado informado."""
        if not isinstance(estado, EstadoInterface):
            estado = EstadoInterface(estado)

        configuracao = obter_configuracao_interface(estado, total_arquivos)
        estado_selecionar = "normal" if configuracao.selecionar_habilitado else "disabled"
        estado_limpar = "normal" if configuracao.limpar_habilitado else "disabled"
        estado_processar = "normal" if configuracao.processar_habilitado else "disabled"
        estado_cancelar = "normal" if configuracao.cancelar_habilitado else "disabled"
        estado_configuracao = "normal" if configuracao.configuracao_habilitada else "disabled"

        self.estado_interface = estado
        if estado in (EstadoInterface.PROCESSANDO, EstadoInterface.CANCELANDO):
            self.progress.configure(mode="indeterminate")
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(self._valor_progresso)

        self.btn_selecionar.configure(
            text=configuracao.texto_selecionar,
            state=estado_selecionar,
        )
        self.btn_limpar.configure(state=estado_limpar)
        self.btn_processar.configure(
            text=configuracao.texto_processar,
            state=estado_processar,
        )
        self.btn_cancelar.configure(
            text=configuracao.texto_cancelar,
            state=estado_cancelar,
        )
        self.combo_departamento.configure(
            state="readonly" if configuracao.configuracao_habilitada else "disabled"
        )
        for checkbox in (self.check_saldo, self.check_resumo, self.check_ranking):
            checkbox.configure(state=estado_configuracao)

    def _processar_clicado(self):
        self.controller.processar(
            self.combo_departamento.get(),
            gerar_saldo=self.var_saldo.get(),
            gerar_resumo=self.var_resumo.get(),
            gerar_ranking=self.var_ranking.get(),
        )

    def _ao_tentar_fechar(self):
        if self.controller.processamento_em_andamento:
            messagebox.showwarning(
                "Processamento em andamento",
                "Aguarde o processamento terminar antes de fechar a janela.",
                parent=self.root,
            )
            return

        self.root.destroy()

    def agendar_na_interface(self, atraso_ms, callback):
        """Agenda uma chamada no loop principal da interface."""
        return self.root.after(atraso_ms, callback)

    def _abrir_dialogo_rhid(self):
        if self._rhid_dialog is not None and self._rhid_dialog.winfo_exists():
            self._rhid_dialog.focus_force()
            return
        self._rhid_dialog = RhidConnectionDialog(
            self.root,
            self.controller.conectar_rhid,
            self.controller.gerar_relatorio_rhid,
        )

    def definir_conexao_rhid_ocupada(self, ocupado):
        if self._rhid_dialog is not None and self._rhid_dialog.winfo_exists():
            self._rhid_dialog.definir_ocupado(ocupado)

    def exibir_catalogo_rhid(self, empresas, departamentos):
        if self._rhid_dialog is not None and self._rhid_dialog.winfo_exists():
            self._rhid_dialog.exibir_catalogo(empresas, departamentos)

    def exibir_dominios_rhid(self, tenants):
        if self._rhid_dialog is not None and self._rhid_dialog.winfo_exists():
            self._rhid_dialog.exibir_dominios(tenants)

    def definir_geracao_rhid_ocupada(self, ocupado):
        if self._rhid_dialog is not None and self._rhid_dialog.winfo_exists():
            self._rhid_dialog.definir_geracao_ocupada(ocupado)

    def atualizar_progresso_rhid(self, valor, mensagem=""):
        if self._rhid_dialog is not None and self._rhid_dialog.winfo_exists():
            self._rhid_dialog.atualizar_progresso(valor, mensagem)

    def exibir_sucesso_rhid(self, mensagem):
        if self._rhid_dialog is not None and self._rhid_dialog.winfo_exists():
            self._rhid_dialog.exibir_sucesso_geracao(mensagem)

    def exibir_erro_rhid(self, mensagem):
        if self._rhid_dialog is not None and self._rhid_dialog.winfo_exists():
            self._rhid_dialog.exibir_erro(mensagem)
            messagebox.showerror(
                "Operação do RHiD não concluída",
                mensagem,
                parent=self._rhid_dialog,
            )

    def _criar_box_metrica(self, parent, titulo, valor):
        box = ctk.CTkFrame(parent, fg_color=BG_BOX, corner_radius=14, border_width=1, border_color=BORDER)
        ctk.CTkLabel(box, text=titulo, text_color=FG_MUTED, font=FONT_METRIC_TITLE).pack(pady=(12, 2))
        valor_label = ctk.CTkLabel(box, text=valor, text_color=FG_TITLE, font=FONT_METRIC_VALUE)
        valor_label.pack(pady=(0, 12))
        box.valor_label = valor_label
        return box

    def _montar_lateral(self, parent):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=20, border_width=1, border_color=BORDER)
        card.pack(fill="both", expand=True)
        ctk.CTkLabel(card, text="Últimos processamentos", text_color=FG_TITLE, font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=14, pady=(14, 10))
        self.historico_frame = ctk.CTkScrollableFrame(card, fg_color="transparent")
        self.historico_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def atualizar_departamentos(self, departamentos, selecionado="Todos"):
        self.combo_departamento.configure(values=departamentos)
        valor = selecionado if selecionado in departamentos else "Todos"
        self.combo_departamento.set(valor)

    def atualizar_arquivo(self, texto):
        self.label_arquivo.configure(text=texto)

    def atualizar_pasta_saida(self, texto):
        self.label_pasta_saida.configure(text=texto)

    def atualizar_metricas(self, funcionarios, banco_total, banco_saldo):
        self.metric_func.valor_label.configure(text=str(funcionarios))
        self.metric_bt.valor_label.configure(text=banco_total)
        self.metric_bs.valor_label.configure(text=banco_saldo)

    def atualizar_status(self, texto, tipo="info"):
        cor = PRIMARY
        if tipo == "success":
            cor = SUCCESS
        elif tipo == "warning":
            cor = WARNING
        elif tipo == "error":
            cor = ERROR

        prefixos = {
            "info": "Informação",
            "success": "Sucesso",
            "warning": "Atenção",
            "error": "Erro",
        }
        prefixo = prefixos.get(tipo, prefixos["info"])
        texto = str(texto)
        prefixos_existentes = tuple(f"{valor}:" for valor in prefixos.values())
        texto_acessivel = texto if texto.startswith(prefixos_existentes) else f"{prefixo}: {texto}"

        self.label_status.configure(text=texto_acessivel, text_color=cor)
        self.root.update_idletasks()

    def atualizar_progresso(self, valor):
        self._valor_progresso = valor
        self.progress.set(valor)
        self.root.update_idletasks()

    def atualizar_tempo_execucao(self, segundos):
        if segundos is None:
            self.label_tempo.configure(text="--")
        else:
            self.label_tempo.configure(text=f"{segundos:.1f}s")

    def atualizar_versao(self):
        self.label_versao.configure(text=f"Versão {APP_VERSION}")

    def habilitar_botao_abrir(self, habilitar):
        self.btn_abrir.configure(state="normal" if habilitar else "disabled")

    def habilitar_botao_abrir_pasta(self, habilitar):
        self.btn_abrir_pasta.configure(state="normal" if habilitar else "disabled")

    def renderizar_historico(self, itens):
        for widget in self.historico_frame.winfo_children():
            widget.destroy()

        if not itens:
            ctk.CTkLabel(self.historico_frame, text="Nenhum processamento registrado.", text_color=FG_MUTED, font=("Segoe UI", 11)).pack(anchor="w", padx=6, pady=6)
            return

        for item in itens:
            box = ctk.CTkFrame(self.historico_frame, fg_color=BG_BOX, corner_radius=12, border_width=1, border_color=BORDER)
            box.pack(fill="x", padx=4, pady=4)
            ctk.CTkLabel(box, text=item.get("data_execucao", ""), text_color=FG_MUTED, font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(10, 2))
            ctk.CTkLabel(box, text=f"Departamento: {item.get('departamento', 'Todos')}", text_color=FG_TEXT, font=("Segoe UI", 10)).pack(anchor="w", padx=10)
            ctk.CTkLabel(box, text=f"Funcionários: {item.get('quantidade_funcionarios', 0)}", text_color=FG_TEXT, font=("Segoe UI", 10)).pack(anchor="w", padx=10)
            ctk.CTkLabel(box, text=f"BT: {item.get('banco_total', '--:--')} | BS: {item.get('banco_saldo', '--:--')}", text_color=FG_TEXT, font=("Segoe UI", 10)).pack(anchor="w", padx=10)
            abas = []
            if item.get("gerou_saldo", True):
                abas.append("SALDO")
            if item.get("gerou_resumo", True):
                abas.append("RESUMO")
            if item.get("gerou_ranking", True):
                abas.append("RANKING")
            ctk.CTkLabel(box, text=f"Abas: {', '.join(abas) if abas else 'Somente principal'}", text_color=FG_MUTED, font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(0, 10))

    def run(self):
        self.root.mainloop()


def iniciar_app():
    app = MainWindow()
    app.run()
