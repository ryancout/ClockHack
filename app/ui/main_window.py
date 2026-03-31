import os
import sys
import ctypes

import tkinter as tk
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
        self.root.geometry(APP_GEOMETRY)
        self.root.minsize(MIN_WIDTH, MIN_HEIGHT)
        self.root.configure(fg_color=BG_APP)

        self.controller = MainController(self)
        self.logo_ref = None

        self._montar_layout()
        self.controller.iniciar()

    def _montar_layout(self):
        container = ctk.CTkFrame(self.root, fg_color=BG_APP, corner_radius=0)
        container.pack(fill="both", expand=True, padx=18, pady=18)

        esquerdo = ctk.CTkFrame(container, fg_color=BG_APP, corner_radius=0)
        esquerdo.pack(side="left", fill="both", expand=True, padx=(0, 10))

        direito = ctk.CTkFrame(container, fg_color=BG_APP, corner_radius=0, width=320)
        direito.pack(side="right", fill="y")
        direito.pack_propagate(False)

        self._montar_card_principal(esquerdo)
        self._montar_lateral(direito)

    def _montar_card_principal(self, parent):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=20, border_width=1, border_color=BORDER)
        card.pack(fill="both", expand=True)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(24, 8))

        logo_path = resource_path("app/assets/logo.png")
        if os.path.exists(logo_path):
            try:
                logo = ctk.CTkImage(light_image=Image.open(logo_path), size=(150, 76))
                ctk.CTkLabel(header, image=logo, text="").pack(anchor="center")
                self.logo_ref = logo
            except Exception:
                pass

        ctk.CTkLabel(card, text="Processador de Planilhas", font=FONT_TITLE, text_color=FG_TITLE).pack(pady=(4, 4))
        ctk.CTkLabel(
            card,
            text="Selecione arquivo(s), escolha o departamento, marque as abas desejadas, processe e salve a saída em Excel.",
            font=FONT_SUBTITLE,
            text_color=FG_MUTED,
            wraplength=760,
        ).pack(pady=(0, 10))

        self.label_versao = ctk.CTkLabel(card, text="", font=("Segoe UI", 10, "bold"), text_color=PRIMARY)
        self.label_versao.pack(pady=(0, 14))

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.pack(fill="x", padx=24, pady=(0, 12))

        top_actions = ctk.CTkFrame(actions, fg_color="transparent")
        top_actions.pack(fill="x", pady=(0, 10))

        ctk.CTkButton(
            top_actions,
            text="Selecionar arquivo(s)",
            height=42,
            fg_color=PRIMARY,
            hover_color="#0955af",
            font=FONT_BUTTON,
            command=self.controller.selecionar_arquivos,
        ).pack(side="left", fill="x", expand=True, padx=(0, 5))

        ctk.CTkButton(
            top_actions,
            text="Limpar seleção",
            height=42,
            fg_color="#e9eef5",
            text_color=FG_TEXT,
            hover_color="#dde6f1",
            font=FONT_BUTTON,
            command=self.controller.limpar_selecao,
        ).pack(side="left", fill="x", expand=True, padx=(5, 0))

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
        self.var_resumo = ctk.BooleanVar(value=True)
        self.var_ranking = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(checks, text="Gerar aba RESUMO", variable=self.var_resumo).pack(anchor="w", pady=2)
        ctk.CTkCheckBox(checks, text="Gerar aba RANKING", variable=self.var_ranking).pack(anchor="w", pady=2)

        ctk.CTkButton(
            actions,
            text="Processar arquivo(s)",
            height=44,
            fg_color=SUCCESS,
            hover_color="#0d634d",
            font=FONT_BUTTON,
            command=self._processar_clicado,
        ).pack(fill="x", pady=(0, 10))

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

        box_arquivo = ctk.CTkFrame(card, fg_color=BG_BOX, corner_radius=14, border_width=1, border_color=BORDER)
        box_arquivo.pack(fill="x", padx=24, pady=(0, 10))
        ctk.CTkLabel(box_arquivo, text="Arquivo selecionado", text_color=FG_MUTED, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        self.label_arquivo = ctk.CTkLabel(
            box_arquivo,
            text="Nenhum arquivo selecionado",
            text_color=FG_TEXT,
            font=("Segoe UI", 12),
            wraplength=720,
            justify="left",
        )
        self.label_arquivo.pack(anchor="w", padx=14, pady=(0, 12))

        box_saida = ctk.CTkFrame(card, fg_color=BG_BOX, corner_radius=14, border_width=1, border_color=BORDER)
        box_saida.pack(fill="x", padx=24, pady=(0, 14))
        ctk.CTkLabel(box_saida, text="Pasta de saída", text_color=FG_MUTED, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14, pady=(12, 2))
        self.label_pasta_saida = ctk.CTkLabel(
            box_saida,
            text="Nenhuma pasta selecionada ainda.",
            text_color=FG_TEXT,
            font=("Segoe UI", 12),
            wraplength=720,
            justify="left",
        )
        self.label_pasta_saida.pack(anchor="w", padx=14, pady=(0, 12))

        metricas = ctk.CTkFrame(card, fg_color="transparent")
        metricas.pack(fill="x", padx=24, pady=(0, 14))

        self.metric_func = self._criar_box_metrica(metricas, "Funcionários", "0")
        self.metric_bt = self._criar_box_metrica(metricas, "Banco Total", "--:--")
        self.metric_bs = self._criar_box_metrica(metricas, "Banco Saldo", "--:--")

        self.metric_func.pack(side="left", fill="both", expand=True, padx=(0, 6))
        self.metric_bt.pack(side="left", fill="both", expand=True, padx=6)
        self.metric_bs.pack(side="left", fill="both", expand=True, padx=(6, 0))

        tempo_box = ctk.CTkFrame(card, fg_color=BG_BOX, corner_radius=14, border_width=1, border_color=BORDER)
        tempo_box.pack(fill="x", padx=24, pady=(0, 10))
        ctk.CTkLabel(tempo_box, text="Tempo de execução", text_color=FG_MUTED, font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
        self.label_tempo = ctk.CTkLabel(tempo_box, text="--", text_color=FG_TITLE, font=("Segoe UI", 14, "bold"))
        self.label_tempo.pack(anchor="w", padx=14, pady=(0, 10))

        self.progress = ctk.CTkProgressBar(card, height=12)
        self.progress.pack(fill="x", padx=24, pady=(0, 10))
        self.progress.set(0)

        self.label_status = ctk.CTkLabel(card, text="Aguardando arquivo.", text_color=FG_MUTED, font=FONT_STATUS, wraplength=760)
        self.label_status.pack(fill="x", padx=24, pady=(0, 18))

    def _processar_clicado(self):
        self.controller.processar(
            self.combo_departamento.get(),
            gerar_resumo=self.var_resumo.get(),
            gerar_ranking=self.var_ranking.get(),
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
        self.label_status.configure(text=texto, text_color=cor)
        self.root.update_idletasks()

    def atualizar_progresso(self, valor):
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
