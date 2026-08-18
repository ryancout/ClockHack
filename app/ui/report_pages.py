"""Páginas reutilizáveis do fluxo principal do FAS Jornada.

O módulo contém apenas componentes de apresentação. A navegação, os
diálogos de arquivo e o processamento continuam sob responsabilidade da janela
principal e do controlador.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable

import customtkinter as ctk


FAS_BACKGROUND = "#2A495B"
CARD_BACKGROUND = "#FFFFFF"
BOX_BACKGROUND = "#F4F7F9"
TEXT_PRIMARY = "#173646"
TEXT_MUTED = "#617783"
BORDER_COLOR = "#D8E1E6"
PRIMARY = "#1769E0"
PRIMARY_HOVER = "#1257BC"
SUCCESS = "#16845F"
SUCCESS_HOVER = "#116C4D"
ERROR = "#B42318"
WARNING = "#8A4F00"


def _estado_widget(habilitado: bool, *, somente_leitura: bool = False) -> str:
    if not habilitado:
        return "disabled"
    return "readonly" if somente_leitura else "normal"


def _normalizar_progresso(valor: object) -> float:
    """Aceita fração ou percentual e limita o resultado entre zero e um."""
    try:
        progresso = float(valor)
    except (TypeError, ValueError):
        return 0.0
    if progresso > 1:
        progresso /= 100
    return max(0.0, min(1.0, progresso))


def _cor_status(tipo: str) -> str:
    return {
        "success": SUCCESS,
        "warning": WARNING,
        "error": ERROR,
    }.get(tipo, TEXT_MUTED)


class _ReportPage(ctk.CTkFrame):
    """Base visual comum, sem controlar posicionamento dentro da MainWindow."""

    def __init__(self, parent) -> None:
        super().__init__(parent, fg_color=FAS_BACKGROUND, corner_radius=0)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def _criar_card(
        self,
        *,
        parent=None,
        padx: int = 28,
        pady: int = 24,
    ) -> ctk.CTkFrame:
        parent = parent or self
        card = ctk.CTkFrame(
            parent,
            fg_color=CARD_BACKGROUND,
            corner_radius=22,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        sticky = "nsew" if parent is self else "ew"
        card.grid(row=0, column=0, sticky=sticky, padx=padx, pady=pady)
        card.grid_columnconfigure(0, weight=1)
        return card

    def _criar_area_rolavel(self) -> ctk.CTkScrollableFrame:
        area = ctk.CTkScrollableFrame(
            self,
            fg_color=FAS_BACKGROUND,
            corner_radius=0,
            scrollbar_button_color="#8195A0",
            scrollbar_button_hover_color="#A3B2BA",
        )
        area.grid(row=0, column=0, sticky="nsew")
        area.grid_columnconfigure(0, weight=1)
        return area

    @staticmethod
    def _titulo(parent, titulo: str, subtitulo: str, *, row: int) -> int:
        ctk.CTkLabel(
            parent,
            text=titulo,
            font=("Segoe UI", 25, "bold"),
            text_color=TEXT_PRIMARY,
            anchor="w",
        ).grid(row=row, column=0, sticky="ew", padx=28, pady=(14, 3))
        ctk.CTkLabel(
            parent,
            text=subtitulo,
            font=("Segoe UI", 11),
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=680,
        ).grid(row=row + 1, column=0, sticky="ew", padx=28, pady=(0, 18))
        return row + 2

    @staticmethod
    def _botao_voltar(parent, callback: Callable[[], None]) -> ctk.CTkButton:
        return ctk.CTkButton(
            parent,
            text="← Voltar",
            width=96,
            height=34,
            fg_color="transparent",
            text_color=TEXT_PRIMARY,
            hover_color=BOX_BACKGROUND,
            border_width=1,
            border_color=BORDER_COLOR,
            font=("Segoe UI", 10, "bold"),
            command=callback,
        )


class HomePage(_ReportPage):
    """Entrada simples do aplicativo com as duas origens em coluna."""

    def __init__(
        self,
        parent,
        ao_csv: Callable[[], None],
        ao_rhid: Callable[[], None],
    ) -> None:
        super().__init__(parent)
        card = self._criar_card(padx=32, pady=30)
        card.grid_rowconfigure(0, weight=1)

        apresentacao = ctk.CTkFrame(
            card,
            width=560,
            height=350,
            fg_color="transparent",
        )
        apresentacao.grid(row=0, column=0)
        apresentacao.grid_propagate(False)
        apresentacao.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            apresentacao,
            text="FAS Jornada",
            font=("Segoe UI", 34, "bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=0, column=0, sticky="ew", pady=(8, 4))
        ctk.CTkLabel(
            apresentacao,
            text="Relatórios e análises de jornada em um só lugar",
            font=("Segoe UI", 13),
            text_color=TEXT_MUTED,
        ).grid(row=1, column=0, sticky="ew")
        ctk.CTkLabel(
            apresentacao,
            text="Como deseja começar?",
            font=("Segoe UI", 17, "bold"),
            text_color=TEXT_PRIMARY,
        ).grid(row=2, column=0, sticky="ew", pady=(38, 20))

        acoes = ctk.CTkFrame(
            apresentacao,
            width=360,
            height=130,
            fg_color="transparent",
        )
        acoes.grid(row=3, column=0)
        acoes.grid_propagate(False)
        acoes.grid_columnconfigure(0, weight=1)
        acoes.grid_rowconfigure((0, 1), weight=1)

        self.btn_csv = ctk.CTkButton(
            acoes,
            text="Arquivo CSV\nImportar relatório já exportado",
            height=54,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            font=("Segoe UI", 14, "bold"),
            command=ao_csv,
        )
        self.btn_csv.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        self.btn_rhid = ctk.CTkButton(
            acoes,
            text="Integração RHiD\nGerar relatório diretamente",
            height=54,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            font=("Segoe UI", 14, "bold"),
            command=ao_rhid,
        )
        self.btn_rhid.grid(row=1, column=0, sticky="ew", pady=(6, 0))


class CsvPage(_ReportPage):
    """Seleção e configuração dos relatórios originados em CSV."""

    def __init__(
        self,
        parent,
        ao_voltar: Callable[[], None],
        ao_selecionar: Callable[[], None],
        ao_limpar: Callable[[], None],
        ao_processar: Callable[..., None],
    ) -> None:
        super().__init__(parent)
        self._ao_processar = ao_processar
        self._rolagem = self._criar_area_rolavel()
        card = self._criar_card(parent=self._rolagem, pady=4)

        self.btn_voltar = self._botao_voltar(card, ao_voltar)
        self.btn_voltar.grid(row=0, column=0, sticky="w", padx=28, pady=(12, 0))
        linha = self._titulo(
            card,
            "Processar arquivo CSV",
            "Selecione um ou mais arquivos. Cada CSV gerará seu próprio Excel.",
            row=1,
        )

        acoes = ctk.CTkFrame(card, fg_color="transparent")
        acoes.grid(row=linha, column=0, sticky="ew", padx=28)
        acoes.grid_columnconfigure((0, 1), weight=1, uniform="acoes_csv")
        self.btn_selecionar = ctk.CTkButton(
            acoes,
            text="Selecionar arquivo(s)",
            height=42,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            font=("Segoe UI", 11, "bold"),
            command=ao_selecionar,
        )
        self.btn_selecionar.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.btn_limpar = ctk.CTkButton(
            acoes,
            text="Limpar seleção",
            height=42,
            fg_color=BOX_BACKGROUND,
            hover_color="#E6EDF1",
            text_color=TEXT_PRIMARY,
            border_width=1,
            border_color=BORDER_COLOR,
            font=("Segoe UI", 11, "bold"),
            command=ao_limpar,
        )
        self.btn_limpar.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        arquivo_box = ctk.CTkFrame(
            card,
            fg_color=BOX_BACKGROUND,
            corner_radius=12,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        arquivo_box.grid(row=linha + 1, column=0, sticky="ew", padx=28, pady=(14, 10))
        arquivo_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            arquivo_box,
            text="Arquivo(s) selecionado(s)",
            font=("Segoe UI", 10, "bold"),
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=14, pady=(11, 2))
        self.label_arquivo = ctk.CTkLabel(
            arquivo_box,
            text="Nenhum arquivo selecionado",
            font=("Segoe UI", 11),
            text_color=TEXT_PRIMARY,
            anchor="w",
            justify="left",
            wraplength=650,
        )
        self.label_arquivo.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 11))

        configuracao = ctk.CTkFrame(card, fg_color="transparent")
        configuracao.grid(row=linha + 2, column=0, sticky="ew", padx=28)
        configuracao.grid_columnconfigure((0, 1), weight=1, uniform="config_csv")

        departamento_box = ctk.CTkFrame(
            configuracao,
            fg_color=BOX_BACKGROUND,
            corner_radius=12,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        departamento_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        departamento_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            departamento_box,
            text="Departamento",
            font=("Segoe UI", 10, "bold"),
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 5))
        self.combo_departamento = ctk.CTkComboBox(
            departamento_box,
            values=["Todos"],
            state="readonly",
            height=36,
        )
        self.combo_departamento.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 14))
        self.combo_departamento.set("Todos")

        opcoes_box = ctk.CTkFrame(
            configuracao,
            fg_color=BOX_BACKGROUND,
            corner_radius=12,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        opcoes_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        ctk.CTkLabel(
            opcoes_box,
            text="Abas do relatório",
            font=("Segoe UI", 10, "bold"),
            text_color=TEXT_MUTED,
            anchor="w",
        ).pack(fill="x", padx=14, pady=(12, 5))

        checks = ctk.CTkFrame(opcoes_box, fg_color="transparent")
        checks.pack(fill="x", padx=12, pady=(0, 12))
        checks.grid_columnconfigure((0, 1, 2), weight=1, uniform="abas_csv")
        self.var_saldo = ctk.BooleanVar(value=True)
        self.var_resumo = ctk.BooleanVar(value=True)
        self.var_ranking = ctk.BooleanVar(value=True)
        self.check_saldo = ctk.CTkCheckBox(
            checks, text="SALDO", variable=self.var_saldo
        )
        self.check_resumo = ctk.CTkCheckBox(
            checks, text="RESUMO", variable=self.var_resumo
        )
        self.check_ranking = ctk.CTkCheckBox(
            checks, text="RANKING", variable=self.var_ranking
        )
        self.check_saldo.grid(row=0, column=0, sticky="w", padx=2)
        self.check_resumo.grid(row=0, column=1, sticky="w", padx=2)
        self.check_ranking.grid(row=0, column=2, sticky="w", padx=2)
        # Compatibilidade com os nomes usados pela tela atual.
        self.checkbox_saldo = self.check_saldo
        self.checkbox_resumo = self.check_resumo
        self.checkbox_ranking = self.check_ranking

        self.label_status = ctk.CTkLabel(
            card,
            text="Selecione o(s) arquivo(s) para continuar.",
            font=("Segoe UI", 10),
            text_color=TEXT_MUTED,
            anchor="w",
            justify="left",
            wraplength=680,
        )
        self.label_status.grid(row=linha + 3, column=0, sticky="ew", padx=28, pady=(13, 8))

        self.btn_processar = ctk.CTkButton(
            card,
            text="Processar arquivo(s)",
            height=44,
            state="disabled",
            fg_color=SUCCESS,
            hover_color=SUCCESS_HOVER,
            font=("Segoe UI", 12, "bold"),
            command=self._processar,
        )
        self.btn_processar.grid(row=linha + 4, column=0, sticky="ew", padx=28, pady=(0, 24))

    def _processar(self) -> None:
        self._ao_processar(
            self.combo_departamento.get(),
            gerar_saldo=bool(self.var_saldo.get()),
            gerar_resumo=bool(self.var_resumo.get()),
            gerar_ranking=bool(self.var_ranking.get()),
        )

    def atualizar_arquivo(self, texto: object) -> None:
        self.label_arquivo.configure(text=str(texto))

    def atualizar_departamentos(
        self,
        departamentos: Iterable[object],
        selecionado: str = "Todos",
    ) -> None:
        valores = [str(item) for item in departamentos]
        if "Todos" not in valores:
            valores.insert(0, "Todos")
        valor = selecionado if selecionado in valores else "Todos"
        self.combo_departamento.configure(values=valores)
        self.combo_departamento.set(valor)

    def atualizar_status(self, texto: object, tipo: str = "info") -> None:
        self.label_status.configure(text=str(texto), text_color=_cor_status(tipo))

    def definir_acoes(
        self,
        *,
        selecionar_habilitado: bool,
        limpar_habilitado: bool,
        processar_habilitado: bool,
        configuracao_habilitada: bool = True,
        texto_selecionar: str | None = None,
        texto_processar: str | None = None,
    ) -> None:
        configuracoes_selecao = {
            "state": _estado_widget(selecionar_habilitado),
        }
        if texto_selecionar is not None:
            configuracoes_selecao["text"] = texto_selecionar
        self.btn_selecionar.configure(**configuracoes_selecao)
        self.btn_limpar.configure(state=_estado_widget(limpar_habilitado))

        configuracoes_processamento = {
            "state": _estado_widget(processar_habilitado),
        }
        if texto_processar is not None:
            configuracoes_processamento["text"] = texto_processar
        self.btn_processar.configure(**configuracoes_processamento)

        self.combo_departamento.configure(
            state=_estado_widget(configuracao_habilitada, somente_leitura=True)
        )
        estado_opcao = _estado_widget(configuracao_habilitada)
        for checkbox in (self.check_saldo, self.check_resumo, self.check_ranking):
            checkbox.configure(state=estado_opcao)

    def redefinir_opcoes(self) -> None:
        self.var_saldo.set(True)
        self.var_resumo.set(True)
        self.var_ranking.set(True)


class ProcessingPage(_ReportPage):
    """Acompanhamento compartilhado pelos fluxos CSV e RHiD."""

    def __init__(self, parent, ao_cancelar: Callable[[], None]) -> None:
        super().__init__(parent)
        card = self._criar_card(padx=56, pady=54)
        card.grid_rowconfigure(2, weight=1)

        self.label_origem = ctk.CTkLabel(
            card,
            text="Gerando relatório",
            font=("Segoe UI", 25, "bold"),
            text_color=TEXT_PRIMARY,
        )
        self.label_origem.grid(row=0, column=0, sticky="ew", padx=28, pady=(34, 8))
        self.label_status = ctk.CTkLabel(
            card,
            text="Preparando o processamento...",
            font=("Segoe UI", 11),
            text_color=TEXT_MUTED,
            justify="left",
            wraplength=650,
        )
        self.label_status.grid(row=1, column=0, sticky="ew", padx=28)

        conteudo = ctk.CTkFrame(card, fg_color="transparent")
        conteudo.grid(row=2, column=0, sticky="ew", padx=28)
        conteudo.grid_columnconfigure(0, weight=1)
        self.progress = ctk.CTkProgressBar(
            conteudo,
            height=12,
            progress_color=PRIMARY,
        )
        self.progress.grid(row=0, column=0, sticky="ew", pady=(24, 10))
        self.progress.set(0)
        self.label_percentual = ctk.CTkLabel(
            conteudo,
            text="0%",
            font=("Segoe UI", 11, "bold"),
            text_color=TEXT_PRIMARY,
        )
        self.label_percentual.grid(row=1, column=0, sticky="e")

        self.btn_cancelar = ctk.CTkButton(
            card,
            text="Cancelar processamento",
            height=40,
            fg_color="#FBE9E7",
            hover_color="#F5D5D1",
            text_color=ERROR,
            border_width=1,
            border_color="#F0C8C3",
            font=("Segoe UI", 11, "bold"),
            command=ao_cancelar,
        )
        self.btn_cancelar.grid(row=3, column=0, sticky="ew", padx=28, pady=(16, 32))

    def definir_origem(self, origem: object) -> None:
        texto = str(origem).strip()
        self.label_origem.configure(
            text=f"Gerando relatório — {texto}" if texto else "Gerando relatório"
        )

    def atualizar_progresso(self, valor: object, mensagem: str = "") -> None:
        progresso = _normalizar_progresso(valor)
        self.progress.set(progresso)
        self.label_percentual.configure(text=f"{round(progresso * 100)}%")
        if mensagem:
            self.atualizar_status(mensagem)

    def atualizar_status(self, texto: object, tipo: str = "info") -> None:
        self.label_status.configure(text=str(texto), text_color=_cor_status(tipo))

    def definir_cancelamento_habilitado(
        self,
        habilitado: bool,
        texto: str | None = None,
    ) -> None:
        configuracao = {"state": _estado_widget(habilitado)}
        if texto is not None:
            configuracao["text"] = texto
        self.btn_cancelar.configure(**configuracao)

    def iniciar_indeterminado(self) -> None:
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self.label_percentual.configure(text="")

    def parar_indeterminado(self) -> None:
        self.progress.stop()
        self.progress.configure(mode="determinate")


class SuccessPage(_ReportPage):
    """Resultado final com atalhos para os arquivos e um novo relatório."""

    def __init__(
        self,
        parent,
        ao_abrir: Callable[[], None],
        ao_abrir_pasta: Callable[[], None],
        ao_gerar_outro: Callable[[], None],
        ao_voltar: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(parent)
        self._rolagem = self._criar_area_rolavel()
        card = self._criar_card(
            parent=self._rolagem,
            padx=42,
            pady=16,
        )
        self.btn_voltar = self._botao_voltar(card, ao_voltar or (lambda: None))
        self.btn_voltar.grid(row=0, column=0, sticky="w", padx=28, pady=(22, 0))
        linha = self._titulo(
            card,
            "Relatório salvo",
            "O processamento foi concluído com sucesso.",
            row=1,
        )

        metricas = ctk.CTkFrame(card, fg_color="transparent")
        metricas.grid(row=linha, column=0, sticky="ew", padx=28)
        metricas.grid_columnconfigure((0, 1, 2), weight=1, uniform="metricas")
        self.metric_funcionarios = self._criar_metrica(metricas, "Funcionários", "0")
        self.metric_banco_total = self._criar_metrica(metricas, "Banco Total", "--:--")
        self.metric_banco_saldo = self._criar_metrica(metricas, "Banco Saldo", "--:--")
        self.metric_funcionarios.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        self.metric_banco_total.grid(row=0, column=1, sticky="nsew", padx=5)
        self.metric_banco_saldo.grid(row=0, column=2, sticky="nsew", padx=(5, 0))

        caminho_box = ctk.CTkFrame(
            card,
            fg_color=BOX_BACKGROUND,
            corner_radius=12,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        caminho_box.grid(row=linha + 1, column=0, sticky="ew", padx=28, pady=(14, 8))
        caminho_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            caminho_box,
            text="Local salvo",
            font=("Segoe UI", 10, "bold"),
            text_color=TEXT_MUTED,
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=14, pady=(11, 2))
        self.label_caminho = ctk.CTkLabel(
            caminho_box,
            text="Nenhum caminho informado.",
            font=("Segoe UI", 11),
            text_color=TEXT_PRIMARY,
            anchor="w",
            justify="left",
            wraplength=650,
        )
        self.label_caminho.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 11))

        self.label_status = ctk.CTkLabel(
            card,
            text="Relatório salvo com sucesso.",
            font=("Segoe UI", 11, "bold"),
            text_color=SUCCESS,
            anchor="w",
            justify="left",
            wraplength=680,
        )
        self.label_status.grid(row=linha + 2, column=0, sticky="ew", padx=28, pady=(6, 12))

        acoes = ctk.CTkFrame(card, fg_color="transparent")
        acoes.grid(row=linha + 3, column=0, sticky="ew", padx=28)
        acoes.grid_columnconfigure((0, 1), weight=1, uniform="abrir")
        self.btn_abrir = ctk.CTkButton(
            acoes,
            text="Abrir arquivo",
            height=40,
            fg_color=BOX_BACKGROUND,
            hover_color="#E6EDF1",
            text_color=TEXT_PRIMARY,
            border_width=1,
            border_color=BORDER_COLOR,
            font=("Segoe UI", 11, "bold"),
            command=ao_abrir,
        )
        self.btn_abrir.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.btn_abrir_pasta = ctk.CTkButton(
            acoes,
            text="Abrir pasta",
            height=40,
            fg_color=BOX_BACKGROUND,
            hover_color="#E6EDF1",
            text_color=TEXT_PRIMARY,
            border_width=1,
            border_color=BORDER_COLOR,
            font=("Segoe UI", 11, "bold"),
            command=ao_abrir_pasta,
        )
        self.btn_abrir_pasta.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        self.btn_gerar_outro = ctk.CTkButton(
            card,
            text="Gerar outro relatório",
            height=44,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            font=("Segoe UI", 12, "bold"),
            command=ao_gerar_outro,
        )
        self.btn_gerar_outro.grid(
            row=linha + 4,
            column=0,
            sticky="ew",
            padx=28,
            pady=(14, 26),
        )

    @staticmethod
    def _criar_metrica(parent, titulo: str, valor: str) -> ctk.CTkFrame:
        box = ctk.CTkFrame(
            parent,
            fg_color=BOX_BACKGROUND,
            corner_radius=12,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        ctk.CTkLabel(
            box,
            text=titulo,
            font=("Segoe UI", 10, "bold"),
            text_color=TEXT_MUTED,
        ).pack(pady=(12, 2))
        label_valor = ctk.CTkLabel(
            box,
            text=valor,
            font=("Segoe UI", 20, "bold"),
            text_color=TEXT_PRIMARY,
        )
        label_valor.pack(pady=(0, 12))
        box.valor_label = label_valor
        return box

    def atualizar_metricas(
        self,
        funcionarios: object,
        banco_total: object,
        banco_saldo: object,
    ) -> None:
        self.metric_funcionarios.valor_label.configure(text=str(funcionarios))
        self.metric_banco_total.valor_label.configure(text=str(banco_total))
        self.metric_banco_saldo.valor_label.configure(text=str(banco_saldo))

    def atualizar_caminho(self, caminho: object) -> None:
        self.label_caminho.configure(text=str(caminho))

    def atualizar_status(self, texto: object, tipo: str = "success") -> None:
        self.label_status.configure(text=str(texto), text_color=_cor_status(tipo))

    def atualizar_resultado(
        self,
        funcionarios: object,
        banco_total: object,
        banco_saldo: object,
        caminho: object,
        status: str = "Relatório salvo com sucesso.",
    ) -> None:
        self.atualizar_metricas(funcionarios, banco_total, banco_saldo)
        self.atualizar_caminho(caminho)
        self.atualizar_status(status, "success")

    def definir_abertura_habilitada(
        self,
        arquivo: bool,
        pasta: bool,
    ) -> None:
        self.btn_abrir.configure(state=_estado_widget(arquivo))
        self.btn_abrir_pasta.configure(state=_estado_widget(pasta))


__all__ = [
    "CsvPage",
    "FAS_BACKGROUND",
    "HomePage",
    "ProcessingPage",
    "SuccessPage",
]
