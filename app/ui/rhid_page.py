"""Página integrada do RHiD para o fluxo principal do FAS Jornada.

O widget não abre janelas secundárias. Login, escolha de cliente e definição
do escopo são etapas internas da mesma página.
"""

from __future__ import annotations

import calendar
from collections.abc import Iterable
from datetime import date, datetime
from typing import Callable

import customtkinter as ctk

from app.core.config import (
    BG_APP,
    BG_BOX,
    BG_CARD,
    BORDER,
    ERROR,
    FG_MUTED,
    FG_TEXT,
    FG_TITLE,
    FONT_BUTTON,
    PRIMARY,
    SUCCESS,
)
from app.ui.responsive import LayoutDensity, LayoutProfile


TODOS_OS_SETORES = "Todos os setores"
RHID_FORGOT_PASSWORD_URL = "https://www.rhid.com.br/v2/#/forgot_password"
ETAPA_LOGIN = "login"
ETAPA_DOMINIO = "dominio"
ETAPA_ESCOPO = "escopo"


def periodo_padrao(hoje: date | None = None) -> tuple[str, str]:
    """Retorna o primeiro dia do mês e hoje em ``DD/MM/AAAA``."""

    hoje = hoje or date.today()
    return hoje.replace(day=1).strftime("%d/%m/%Y"), hoje.strftime("%d/%m/%Y")


def validar_periodo(data_inicial: str, data_final: str) -> tuple[str, str]:
    """Valida datas brasileiras e as converte para o contrato ISO da API."""

    try:
        inicio = datetime.strptime(str(data_inicial).strip(), "%d/%m/%Y").date()
        fim = datetime.strptime(str(data_final).strip(), "%d/%m/%Y").date()
    except (TypeError, ValueError) as exc:
        raise ValueError("Use as datas no formato DD/MM/AAAA.") from exc
    if fim < inicio:
        raise ValueError("A data final não pode ser anterior à data inicial.")
    return inicio.isoformat(), fim.isoformat()


def _chave_id(valor: object) -> str:
    return str(valor).strip()


def _nome_empresa(empresa: object) -> str:
    nome = getattr(empresa, "label", "")
    if callable(nome):
        nome = nome()
    return str(nome or getattr(empresa, "name", "") or "Empresa").strip()


def _nome_departamento(departamento: object) -> str:
    return str(getattr(departamento, "name", "") or "Departamento").strip()


def _rotulo_com_id(nome: str, identificador: object) -> str:
    """Inclui o ID visível para que homônimos nunca sejam confundidos."""

    return f"{nome} — ID {identificador}"


def catalogar_por_id(
    itens: Iterable[object],
    obter_nome: Callable[[object], str],
) -> tuple[dict[str, object], dict[str, object]]:
    """Cria mapas estáveis de rótulo/ID, removendo só IDs repetidos."""

    id_por_rotulo: dict[str, object] = {}
    item_por_chave: dict[str, object] = {}
    for item in itens:
        identificador = getattr(item, "id", None)
        chave = _chave_id(identificador)
        if identificador is None or not chave or chave in item_por_chave:
            continue
        item_por_chave[chave] = item
        id_por_rotulo[_rotulo_com_id(obter_nome(item), identificador)] = identificador
    return id_por_rotulo, item_por_chave


def departamentos_da_empresa(
    departamentos: Iterable[object], empresa_id: object
) -> tuple[object, ...]:
    """Filtra o catálogo ativo, incluindo cadastros globais do RHiD."""

    chave_empresa = _chave_id(empresa_id)
    vistos: set[str] = set()
    resultado: list[object] = []
    for departamento in departamentos:
        identificador = getattr(departamento, "id", None)
        chave = _chave_id(identificador)
        empresa_departamento = getattr(departamento, "company_id", None)
        chave_empresa_departamento = _chave_id(empresa_departamento)
        departamento_global = empresa_departamento is None or chave_empresa_departamento in {
            "",
            "0",
        }
        if (
            identificador is None
            or not chave
            or chave in vistos
            or (
                not departamento_global
                and chave_empresa_departamento != chave_empresa
            )
        ):
            continue
        vistos.add(chave)
        resultado.append(departamento)
    return tuple(resultado)


class _CalendarioInline(ctk.CTkFrame):
    """Calendário mensal pequeno, sem dependências externas."""

    _DIAS_SEMANA = ("S", "T", "Q", "Q", "S", "S", "D")

    def __init__(self, parent, ao_selecionar: Callable[[date], None]):
        super().__init__(
            parent,
            fg_color=BG_CARD,
            corner_radius=12,
            border_width=1,
            border_color=BORDER,
        )
        self._ao_selecionar = ao_selecionar
        self._ano = date.today().year
        self._mes = date.today().month
        self.btn_anterior = ctk.CTkButton(
            self,
            text="‹",
            width=28,
            height=26,
            fg_color="transparent",
            text_color=FG_TEXT,
            hover_color=BG_BOX,
            command=lambda: self._mudar_mes(-1),
        )
        self.btn_anterior.grid(row=0, column=0, padx=3, pady=3)
        self.label_mes = ctk.CTkLabel(
            self, text="", text_color=FG_TITLE, font=("Segoe UI", 11, "bold")
        )
        self.label_mes.grid(row=0, column=1, columnspan=5, sticky="ew")
        self.btn_proximo = ctk.CTkButton(
            self,
            text="›",
            width=28,
            height=26,
            fg_color="transparent",
            text_color=FG_TEXT,
            hover_color=BG_BOX,
            command=lambda: self._mudar_mes(1),
        )
        self.btn_proximo.grid(row=0, column=6, padx=3, pady=3)

        for coluna, texto in enumerate(self._DIAS_SEMANA):
            ctk.CTkLabel(
                self,
                text=texto,
                text_color=FG_MUTED,
                font=("Segoe UI", 9, "bold"),
            ).grid(row=1, column=coluna, sticky="ew")

        self._botoes_dia: list[ctk.CTkButton] = []
        for indice in range(42):
            botao = ctk.CTkButton(
                self,
                text="",
                width=25,
                height=22,
                font=("Segoe UI", 9),
                fg_color="transparent",
                text_color=FG_TEXT,
                hover_color="#e8f1fc",
            )
            botao.grid(
                row=2 + indice // 7,
                column=indice % 7,
                padx=1,
                pady=1,
            )
            self._botoes_dia.append(botao)
        self._renderizar()

    def abrir_em(self, data_referencia: date) -> None:
        self._ano = data_referencia.year
        self._mes = data_referencia.month
        self._renderizar()

    def definir_habilitado(self, habilitado: bool) -> None:
        estado = "normal" if habilitado else "disabled"
        self.btn_anterior.configure(state=estado)
        self.btn_proximo.configure(state=estado)
        for botao in self._botoes_dia:
            botao.configure(state=estado if botao.cget("text") else "disabled")

    def _mudar_mes(self, deslocamento: int) -> None:
        indice = self._ano * 12 + self._mes - 1 + deslocamento
        self._ano, mes_zero = divmod(indice, 12)
        self._mes = mes_zero + 1
        self._renderizar()

    def _renderizar(self) -> None:
        nomes = (
            "janeiro",
            "fevereiro",
            "março",
            "abril",
            "maio",
            "junho",
            "julho",
            "agosto",
            "setembro",
            "outubro",
            "novembro",
            "dezembro",
        )
        self.label_mes.configure(text=f"{nomes[self._mes - 1].title()} {self._ano}")
        semanas = calendar.Calendar(firstweekday=0).monthdayscalendar(
            self._ano, self._mes
        )
        dias = [dia for semana in semanas for dia in semana]
        dias.extend([0] * (42 - len(dias)))
        hoje = date.today()
        for botao, dia in zip(self._botoes_dia, dias):
            if not dia:
                botao.configure(text="", state="disabled", command=lambda: None)
                continue
            selecionada = date(self._ano, self._mes, dia)
            eh_hoje = selecionada == hoje
            botao.configure(
                text=str(dia),
                state="normal",
                fg_color=PRIMARY if eh_hoje else "transparent",
                text_color="#ffffff" if eh_hoje else FG_TEXT,
                command=lambda valor=selecionada: self._ao_selecionar(valor),
            )


class RhidPage(ctk.CTkFrame):
    """Fluxo visual de conexão e geração direta pelo RHiD."""

    def __init__(
        self,
        parent,
        ao_conectar,
        ao_gerar,
        ao_voltar,
        ao_esqueci_senha=None,
    ):
        super().__init__(parent, fg_color=BG_APP, corner_radius=0)
        self._ao_conectar = ao_conectar
        self._ao_gerar = ao_gerar
        self._ao_voltar = ao_voltar
        self._ao_esqueci_senha = ao_esqueci_senha
        self._etapa = ETAPA_LOGIN
        self._conexao_em_andamento = False
        self._geracao_em_andamento = False
        self._catalogo_carregado = False
        self._empresas: tuple[object, ...] = ()
        self._departamentos: tuple[object, ...] = ()
        self._empresa_id_por_rotulo: dict[str, object] = {}
        self._empresa_por_chave: dict[str, object] = {}
        self._departamentos_visiveis: tuple[object, ...] = ()
        self._departamento_variaveis: dict[str, ctk.BooleanVar] = {}
        self._departamento_id_por_chave: dict[str, object] = {}
        self._departamento_checks: list[ctk.CTkCheckBox] = []
        self._dominio_por_rotulo: dict[str, str] = {}
        self._dominio_preenchido = ""
        self._entrada_calendario = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._montar_cabecalho()
        self._montar_conteudo()
        self._mostrar_etapa(ETAPA_LOGIN)
        self.after(100, self.entry_email.focus_set)

    @property
    def etapa_atual(self) -> str:
        return self._etapa

    def _montar_cabecalho(self) -> None:
        cabecalho = ctk.CTkFrame(self, fg_color="transparent")
        self._cabecalho = cabecalho
        cabecalho.grid(row=0, column=0, sticky="ew", padx=24, pady=(8, 5))
        cabecalho.grid_columnconfigure((0, 2), weight=1, uniform="laterais")
        self.btn_voltar = ctk.CTkButton(
            cabecalho,
            text="← Voltar",
            width=92,
            fg_color="transparent",
            text_color="#ffffff",
            hover_color="#203c4d",
            font=FONT_BUTTON,
            command=self._voltar,
        )
        self.btn_voltar.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(
            cabecalho,
            text="Relatório pelo RHiD",
            text_color="#ffffff",
            font=("Segoe UI", 21, "bold"),
        ).grid(row=0, column=1)
        # Mantém o título no centro exato, compensando o botão à esquerda.
        ctk.CTkFrame(cabecalho, width=92, height=1, fg_color="transparent").grid(
            row=0, column=2, sticky="e"
        )

    def _montar_conteudo(self) -> None:
        self._rolagem = ctk.CTkFrame(
            self,
            fg_color=BG_APP,
            corner_radius=0,
        )
        self._rolagem.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 8))
        self._rolagem.grid_columnconfigure(0, weight=1)
        self._rolagem.grid_rowconfigure(0, weight=1)

        self._etapas = ctk.CTkFrame(self._rolagem, fg_color="transparent")
        self._etapas.grid(row=0, column=0, sticky="new")
        self._etapas.grid_columnconfigure(0, weight=1)
        self._montar_login()
        self._montar_dominio()
        self._montar_escopo()

        self.feedback = ctk.CTkFrame(self._rolagem, fg_color="transparent")
        self.feedback.grid(row=1, column=0, sticky="ew", padx=2, pady=(4, 0))
        self.feedback.grid_columnconfigure(0, weight=1)
        self.progress_geracao = ctk.CTkProgressBar(self.feedback, height=9)
        self.progress_geracao.grid(row=0, column=0, sticky="ew")
        self.progress_geracao.set(0)
        self.progress_geracao.grid_remove()
        self.label_status = ctk.CTkLabel(
            self.feedback,
            text="",
            text_color=FG_MUTED,
            justify="left",
            anchor="w",
            wraplength=760,
        )
        self.label_status.grid(row=1, column=0, sticky="ew", pady=(6, 0))

    def aplicar_densidade(self, profile: LayoutProfile) -> None:
        """Compacta margens do fluxo RHiD sem recorrer a rolagem."""

        if profile.density is LayoutDensity.DENSE:
            header_pad, content_pad, bottom_pad = 8, 6, 3
        elif profile.density is LayoutDensity.COMPACT:
            header_pad, content_pad, bottom_pad = 14, 10, 5
        else:
            header_pad, content_pad, bottom_pad = 24, 20, 8
        self._cabecalho.grid_configure(
            padx=header_pad,
            pady=(max(2, bottom_pad), max(2, bottom_pad - 1)),
        )
        self._rolagem.grid_configure(
            padx=content_pad,
            pady=(0, bottom_pad),
        )

    def _novo_card(self) -> ctk.CTkFrame:
        card = ctk.CTkFrame(
            self._etapas,
            fg_color=BG_CARD,
            corner_radius=18,
            border_width=1,
            border_color=BORDER,
        )
        card.grid_columnconfigure(0, weight=1)
        return card

    @staticmethod
    def _titulo(card, titulo: str, subtitulo: str) -> None:
        ctk.CTkLabel(
            card,
            text=titulo,
            text_color=FG_TITLE,
            font=("Segoe UI", 19, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=24, pady=(14, 2))
        ctk.CTkLabel(
            card,
            text=subtitulo,
            text_color=FG_MUTED,
            justify="left",
            anchor="w",
            wraplength=720,
        ).grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 9))

    @staticmethod
    def _campo(parent, titulo: str, linha: int, **kwargs):
        ctk.CTkLabel(parent, text=titulo, text_color=FG_MUTED, anchor="w").grid(
            row=linha, column=0, sticky="ew", padx=24, pady=(8, 4)
        )
        campo = ctk.CTkEntry(parent, height=38, **kwargs)
        campo.grid(row=linha + 1, column=0, sticky="ew", padx=24)
        return campo

    def _montar_login(self) -> None:
        self.frame_login = self._novo_card()
        self._titulo(
            self.frame_login,
            "Entrar no RHiD",
            "Use as mesmas credenciais do portal. Seus dados são enviados diretamente ao RHiD.",
        )
        self.entry_email = self._campo(
            self.frame_login, "E-mail", 2, placeholder_text="nome@empresa.com"
        )
        self.entry_senha = self._campo(
            self.frame_login, "Senha", 4, show="●", placeholder_text="Sua senha"
        )
        self.var_lembrar = ctk.BooleanVar(value=False)
        self.check_lembrar = ctk.CTkCheckBox(
            self.frame_login,
            text="Lembrar neste computador",
            variable=self.var_lembrar,
            text_color=FG_TEXT,
            fg_color=PRIMARY,
            hover_color="#0955af",
        )
        self.check_lembrar.grid(row=6, column=0, sticky="w", padx=24, pady=(12, 8))
        self.btn_esqueci_senha = ctk.CTkButton(
            self.frame_login,
            text="Esqueci minha senha",
            width=145,
            height=30,
            fg_color="transparent",
            text_color=PRIMARY,
            hover_color=BG_BOX,
            font=("Segoe UI", 10, "underline"),
            command=self._abrir_recuperacao_senha,
        )
        self.btn_esqueci_senha.grid(
            row=7, column=0, sticky="w", padx=18, pady=(0, 6)
        )
        self.btn_conectar = ctk.CTkButton(
            self.frame_login,
            text="Conectar",
            height=42,
            fg_color=PRIMARY,
            hover_color="#0955af",
            font=FONT_BUTTON,
            command=self._conectar_login,
        )
        self.btn_conectar.grid(row=8, column=0, sticky="ew", padx=24, pady=(4, 24))
        self.entry_senha.bind("<Return>", lambda _evento: self._conectar_login())

    def _abrir_recuperacao_senha(self) -> None:
        if self._conexao_em_andamento or self._geracao_em_andamento:
            return
        if self._ao_esqueci_senha is not None:
            self._ao_esqueci_senha()

    def _montar_dominio(self) -> None:
        self.frame_dominio = self._novo_card()
        self._titulo(
            self.frame_dominio,
            "Selecione o cliente",
            "Sua conta possui acesso a mais de um ambiente do RHiD.",
        )
        ctk.CTkLabel(
            self.frame_dominio,
            text="Cliente / domínio",
            text_color=FG_MUTED,
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", padx=24, pady=(8, 4))
        self.combo_dominio = ctk.CTkComboBox(
            self.frame_dominio, values=["Selecione"], state="readonly", height=38
        )
        self.combo_dominio.grid(row=3, column=0, sticky="ew", padx=24)
        self.combo_dominio.set("Selecione")
        self.btn_confirmar_dominio = ctk.CTkButton(
            self.frame_dominio,
            text="Continuar",
            height=42,
            fg_color=PRIMARY,
            hover_color="#0955af",
            font=FONT_BUTTON,
            command=self._conectar_dominio,
        )
        self.btn_confirmar_dominio.grid(
            row=4, column=0, sticky="ew", padx=24, pady=(16, 24)
        )

    def _montar_escopo(self) -> None:
        self.frame_escopo = self._novo_card()
        self._titulo(
            self.frame_escopo,
            "Configurar relatório",
            "Escolha a empresa, os setores, o período e as abas do Excel.",
        )

        conteudo = ctk.CTkFrame(self.frame_escopo, fg_color="transparent")
        conteudo.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 14))
        conteudo.grid_columnconfigure(0, weight=3, uniform="escopo")
        conteudo.grid_columnconfigure(1, weight=2, uniform="escopo")

        selecao = ctk.CTkFrame(conteudo, fg_color="transparent")
        selecao.grid(row=0, column=0, sticky="nsew", padx=(0, 9))
        selecao.grid_columnconfigure(0, weight=1)

        configuracao = ctk.CTkFrame(conteudo, fg_color="transparent")
        configuracao.grid(row=0, column=1, sticky="nsew", padx=(9, 0))
        configuracao.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            selecao, text="Empresa", text_color=FG_MUTED, anchor="w"
        ).grid(row=0, column=0, sticky="ew", pady=(2, 4))
        self.combo_empresa = ctk.CTkComboBox(
            selecao,
            values=["Conecte-se primeiro"],
            state="disabled",
            height=38,
            command=self._empresa_alterada,
        )
        self.combo_empresa.grid(row=1, column=0, sticky="ew")

        ctk.CTkLabel(
            selecao,
            text="Setores / departamentos",
            text_color=FG_MUTED,
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", pady=(10, 4))
        self.departamentos_box = ctk.CTkFrame(
            selecao,
            fg_color=BG_BOX,
            corner_radius=12,
            border_width=1,
            border_color=BORDER,
        )
        self.departamentos_box.grid(row=3, column=0, sticky="ew")
        self.departamentos_box.grid_columnconfigure(0, weight=1)
        self.var_todos_setores = ctk.BooleanVar(value=True)
        self.check_todos_setores = ctk.CTkCheckBox(
            self.departamentos_box,
            text=TODOS_OS_SETORES,
            variable=self.var_todos_setores,
            command=self._alternar_todos_setores,
            text_color=FG_TEXT,
            fg_color=PRIMARY,
            hover_color="#0955af",
        )
        self.check_todos_setores.grid(
            row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(7, 4)
        )

        self._montar_periodo(configuracao)
        self._montar_opcoes_abas(configuracao)
        self.btn_gerar = ctk.CTkButton(
            configuracao,
            text="Gerar relatório",
            height=42,
            fg_color=SUCCESS,
            hover_color="#0b654f",
            font=FONT_BUTTON,
            state="disabled",
            command=self._gerar_relatorio,
        )
        self.btn_gerar.grid(row=2, column=0, sticky="ew", pady=(8, 0))

    def _montar_periodo(self, parent) -> None:
        periodo = ctk.CTkFrame(parent, fg_color="transparent")
        periodo.grid(row=0, column=0, sticky="ew", pady=(2, 1))
        periodo.grid_columnconfigure((0, 1), weight=1)
        inicio_padrao, fim_padrao = periodo_padrao()

        for coluna, (titulo, valor) in enumerate(
            (("Data inicial", inicio_padrao), ("Data final", fim_padrao))
        ):
            bloco = ctk.CTkFrame(periodo, fg_color="transparent")
            bloco.grid(
                row=0,
                column=coluna,
                sticky="ew",
                padx=(0, 6) if coluna == 0 else (6, 0),
            )
            bloco.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(
                bloco,
                text=titulo,
                text_color=FG_MUTED,
                anchor="w",
            ).grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
            entrada = ctk.CTkEntry(bloco, height=38)
            entrada.grid(row=1, column=0, sticky="ew", padx=(0, 5))
            entrada.insert(0, valor)
            botao = ctk.CTkButton(
                bloco,
                text="Abrir",
                width=52,
                height=38,
                fg_color=BG_BOX,
                text_color=PRIMARY,
                hover_color="#e8f1fc",
                border_width=1,
                border_color=BORDER,
                command=lambda campo=entrada: self._alternar_calendario(campo),
            )
            botao.grid(row=1, column=1)
            if coluna == 0:
                self.entry_data_inicial = entrada
                self.btn_calendario_inicial = botao
            else:
                self.entry_data_final = entrada
                self.btn_calendario_final = botao

        self.calendario = _CalendarioInline(self.frame_escopo, self._data_escolhida)
        self.calendario.place_forget()

    def _montar_opcoes_abas(self, parent) -> None:
        opcoes = ctk.CTkFrame(parent, fg_color=BG_BOX, corner_radius=12)
        opcoes.grid(row=1, column=0, sticky="ew", pady=(9, 4))
        opcoes.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(
            opcoes,
            text="Abas adicionais",
            text_color=FG_MUTED,
            anchor="w",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(7, 4))
        self.var_gerar_saldo = ctk.BooleanVar(value=True)
        self.var_gerar_resumo = ctk.BooleanVar(value=True)
        self.var_gerar_ranking = ctk.BooleanVar(value=True)
        self.check_abas: list[ctk.CTkCheckBox] = []
        for indice, (texto, variavel) in enumerate(
            (
                ("SALDO", self.var_gerar_saldo),
                ("RESUMO", self.var_gerar_resumo),
                ("RANKING", self.var_gerar_ranking),
            )
        ):
            check = ctk.CTkCheckBox(
                opcoes,
                text=texto,
                variable=variavel,
                text_color=FG_TEXT,
                fg_color=PRIMARY,
                hover_color="#0955af",
            )
            check.grid(
                row=1 + indice // 2,
                column=indice % 2,
                sticky="w",
                padx=12,
                pady=(0, 8),
            )
            self.check_abas.append(check)

    def _mostrar_etapa(self, etapa: str) -> None:
        if etapa not in {ETAPA_LOGIN, ETAPA_DOMINIO, ETAPA_ESCOPO}:
            raise ValueError(f"Etapa desconhecida: {etapa}")
        self._etapa = etapa
        for frame in (self.frame_login, self.frame_dominio, self.frame_escopo):
            frame.grid_remove()
        {
            ETAPA_LOGIN: self.frame_login,
            ETAPA_DOMINIO: self.frame_dominio,
            ETAPA_ESCOPO: self.frame_escopo,
        }[etapa].grid(row=0, column=0, sticky="ew")
        self.label_status.configure(text="")
        if etapa != ETAPA_ESCOPO:
            self.calendario.place_forget()
            self._entrada_calendario = None
            self.progress_geracao.grid_remove()
        self._atualizar_estados_controles()

    def _voltar(self) -> None:
        if self._conexao_em_andamento or self._geracao_em_andamento:
            return
        if self._etapa in {ETAPA_DOMINIO, ETAPA_ESCOPO}:
            self._mostrar_etapa(ETAPA_LOGIN)
            return
        self._ao_voltar()

    def _conectar_login(self) -> None:
        if self._conexao_em_andamento:
            return
        email, senha, dominio, _lembrar = self.obter_credenciais_digitadas()
        if not email or not senha:
            self.exibir_erro("Informe o e-mail e a senha do RHiD.")
            return
        self._ao_conectar(email, senha, dominio)

    def _conectar_dominio(self) -> None:
        if self._conexao_em_andamento:
            return
        rotulo = self.combo_dominio.get()
        dominio = self._dominio_por_rotulo.get(rotulo)
        email = self.entry_email.get().strip()
        senha = self.entry_senha.get()
        if not dominio:
            self.exibir_erro("Selecione um cliente válido.")
            return
        self._dominio_preenchido = dominio
        self._ao_conectar(email, senha, dominio)

    def obter_credenciais_digitadas(self) -> tuple[str, str, str, bool]:
        """Obtém credenciais; a persistência segura fica a cargo do controller."""

        dominio = self._dominio_preenchido
        if self._dominio_por_rotulo:
            dominio = self._dominio_por_rotulo.get(self.combo_dominio.get(), dominio)
        return (
            self.entry_email.get().strip(),
            self.entry_senha.get(),
            dominio,
            bool(self.var_lembrar.get()),
        )

    @staticmethod
    def _preencher_entry(entry, valor: str) -> None:
        estado = entry.cget("state")
        if estado == "disabled":
            entry.configure(state="normal")
        entry.delete(0, "end")
        entry.insert(0, valor or "")
        if estado == "disabled":
            entry.configure(state="disabled")

    def preencher_credenciais(self, email: str, senha: str, dominio: str = "") -> None:
        self._preencher_entry(self.entry_email, str(email or ""))
        self._preencher_entry(self.entry_senha, str(senha or ""))
        self._dominio_preenchido = str(dominio or "").strip()
        self.var_lembrar.set(bool(email or senha))

    def definir_ocupado(self, ocupado: bool) -> None:
        self._conexao_em_andamento = bool(ocupado)
        if ocupado:
            self.label_status.configure(
                text="Conectando ao RHiD...", text_color=FG_MUTED
            )
        self._atualizar_estados_controles()

    def exibir_dominios(self, tenants) -> None:
        """Mostra a etapa somente quando há mais de um cliente disponível."""

        self.definir_ocupado(False)
        self._dominio_por_rotulo = {}
        unicos: dict[str, object] = {}
        for tenant in tenants or ():
            dominio = str(getattr(tenant, "domain", "") or "").strip()
            if dominio and dominio.casefold() not in unicos:
                unicos[dominio.casefold()] = tenant

        if not unicos:
            self._mostrar_etapa(ETAPA_LOGIN)
            self.exibir_erro("Nenhum cliente foi retornado pelo RHiD.")
            return

        if len(unicos) == 1:
            tenant = next(iter(unicos.values()))
            self._dominio_preenchido = str(
                getattr(tenant, "domain", "") or ""
            ).strip()
            email, senha, dominio, _lembrar = self.obter_credenciais_digitadas()
            self._ao_conectar(email, senha, dominio)
            return

        for tenant in unicos.values():
            dominio = str(getattr(tenant, "domain", "") or "").strip()
            nome = str(getattr(tenant, "name", "") or dominio).strip()
            rotulo = f"{nome} — {dominio}" if nome != dominio else dominio
            if rotulo in self._dominio_por_rotulo:
                tenant_id = getattr(tenant, "tenant_id", None)
                rotulo = _rotulo_com_id(rotulo, tenant_id or dominio)
            self._dominio_por_rotulo[rotulo] = dominio

        rotulos = list(self._dominio_por_rotulo)
        self.combo_dominio.configure(values=rotulos)
        selecionado = next(
            (
                rotulo
                for rotulo, dominio in self._dominio_por_rotulo.items()
                if dominio.casefold() == self._dominio_preenchido.casefold()
            ),
            rotulos[0],
        )
        self.combo_dominio.set(selecionado)
        self._mostrar_etapa(ETAPA_DOMINIO)
        self.label_status.configure(
            text=f"{len(rotulos)} clientes disponíveis.", text_color=FG_MUTED
        )

    def exibir_catalogo(self, empresas, departamentos) -> None:
        """Exibe apenas o catálogo ativo entregue pela integração."""

        self.definir_ocupado(False)
        self._empresas = tuple(empresas or ())
        self._departamentos = tuple(departamentos or ())
        (
            self._empresa_id_por_rotulo,
            self._empresa_por_chave,
        ) = catalogar_por_id(self._empresas, _nome_empresa)
        rotulos = list(self._empresa_id_por_rotulo)
        self._catalogo_carregado = bool(rotulos and self._departamentos)
        self.combo_empresa.configure(
            values=rotulos or ["Nenhuma empresa disponível"],
            state="readonly" if rotulos else "disabled",
        )
        self._mostrar_etapa(ETAPA_ESCOPO)
        if rotulos:
            self.combo_empresa.set(rotulos[0])
            self._empresa_alterada(rotulos[0])
        else:
            self.combo_empresa.set("Nenhuma empresa disponível")
            self._renderizar_departamentos(())

        if not rotulos:
            self.label_status.configure(
                text="O RHiD não retornou empresas disponíveis.", text_color=ERROR
            )
        elif not self._departamentos:
            self.label_status.configure(
                text="Nenhum setor possui funcionário ativo.", text_color=ERROR
            )
        else:
            self._atualizar_status_setores()
        self._atualizar_estados_controles()

    def _empresa_alterada(self, rotulo: str) -> None:
        empresa_id = self._empresa_id_por_rotulo.get(rotulo)
        if empresa_id is None:
            self._renderizar_departamentos(())
            self._atualizar_estados_controles()
            return
        departamentos = departamentos_da_empresa(self._departamentos, empresa_id)
        self._renderizar_departamentos(departamentos)
        self._atualizar_status_setores()
        self._atualizar_estados_controles()

    def _renderizar_departamentos(self, departamentos: Iterable[object]) -> None:
        for check in self._departamento_checks:
            check.destroy()
        self._departamento_checks.clear()
        self._departamento_variaveis.clear()
        self._departamento_id_por_chave.clear()
        self._departamentos_visiveis = tuple(departamentos)
        self.var_todos_setores.set(True)

        quantidade = len(self._departamentos_visiveis)
        colunas = 1 if quantidade <= 5 else 2 if quantidade <= 10 else 3
        self.departamentos_box.grid_columnconfigure(
            tuple(range(colunas)), weight=1, uniform="departamentos"
        )

        for indice, departamento in enumerate(self._departamentos_visiveis):
            identificador = getattr(departamento, "id")
            chave = _chave_id(identificador)
            variavel = ctk.BooleanVar(value=False)
            self._departamento_variaveis[chave] = variavel
            self._departamento_id_por_chave[chave] = identificador
            check = ctk.CTkCheckBox(
                self.departamentos_box,
                text=_nome_departamento(departamento),
                variable=variavel,
                state="disabled",
                text_color=FG_TEXT,
                fg_color=PRIMARY,
                hover_color="#0955af",
                checkbox_width=18,
                checkbox_height=18,
                font=("Segoe UI", 11),
            )
            check.grid(
                row=1 + indice // colunas,
                column=indice % colunas,
                sticky="w",
                padx=10,
                pady=3,
            )
            self._departamento_checks.append(check)

    def _alternar_todos_setores(self) -> None:
        todos = bool(self.var_todos_setores.get())
        for variavel in self._departamento_variaveis.values():
            if todos:
                variavel.set(False)
        self._atualizar_estados_controles()

    def _atualizar_status_setores(self) -> None:
        quantidade = len(self._departamentos_visiveis)
        if quantidade:
            self.label_status.configure(
                text=f"Conectado. {quantidade} setor(es) ativo(s) nesta empresa.",
                text_color=SUCCESS,
            )
        else:
            self.label_status.configure(
                text="Esta empresa não possui setores com funcionários ativos.",
                text_color=ERROR,
            )

    def _alternar_calendario(self, entrada) -> None:
        if self._geracao_em_andamento:
            return
        if self._entrada_calendario is entrada and self.calendario.winfo_ismapped():
            self.calendario.place_forget()
            self._entrada_calendario = None
            return
        try:
            referencia = datetime.strptime(entrada.get().strip(), "%d/%m/%Y").date()
        except ValueError:
            referencia = date.today()
        self._entrada_calendario = entrada
        self.calendario.abrir_em(referencia)
        self.calendario.place(relx=0.78, rely=0.27, anchor="n")
        self.calendario.lift()

    def _data_escolhida(self, valor: date) -> None:
        if self._entrada_calendario is None:
            return
        self._preencher_entry(self._entrada_calendario, valor.strftime("%d/%m/%Y"))
        self.calendario.place_forget()
        self._entrada_calendario = None

    def obter_parametros_geracao(
        self,
    ) -> tuple[object, tuple[object, ...] | None, str, str, bool, bool, bool]:
        rotulo_empresa = self.combo_empresa.get()
        if rotulo_empresa not in self._empresa_id_por_rotulo:
            raise ValueError("Selecione uma empresa válida.")
        empresa_id = self._empresa_id_por_rotulo[rotulo_empresa]
        if not self._departamentos_visiveis:
            raise ValueError("A empresa selecionada não possui setores ativos.")

        if self.var_todos_setores.get():
            departamentos_ids = None
        else:
            departamentos_ids = tuple(
                self._departamento_id_por_chave[chave]
                for chave, variavel in self._departamento_variaveis.items()
                if variavel.get()
            )
            if not departamentos_ids:
                raise ValueError("Selecione ao menos um setor ou marque Todos os setores.")

        inicio_iso, fim_iso = validar_periodo(
            self.entry_data_inicial.get(), self.entry_data_final.get()
        )
        return (
            empresa_id,
            departamentos_ids,
            inicio_iso,
            fim_iso,
            bool(self.var_gerar_saldo.get()),
            bool(self.var_gerar_resumo.get()),
            bool(self.var_gerar_ranking.get()),
        )

    def _gerar_relatorio(self) -> None:
        if self._geracao_em_andamento or self._ao_gerar is None:
            return
        try:
            parametros = self.obter_parametros_geracao()
        except ValueError as exc:
            self.exibir_erro(str(exc))
            return
        try:
            self._ao_gerar(*parametros)
        except Exception as exc:  # mantém falhas síncronas dentro do fluxo visual
            self.exibir_erro(str(exc) or "Não foi possível gerar o relatório.")

    def definir_geracao_ocupada(self, ocupado: bool) -> None:
        self._geracao_em_andamento = bool(ocupado)
        if ocupado:
            self.progress_geracao.set(0)
            self.progress_geracao.grid()
            self.label_status.configure(
                text="Gerando relatório...", text_color=FG_MUTED
            )
        self._atualizar_estados_controles()

    def atualizar_progresso(self, valor, mensagem: str = "") -> None:
        try:
            progresso = float(valor)
        except (TypeError, ValueError):
            progresso = 0.0
        if progresso > 1:
            progresso /= 100
        progresso = min(1.0, max(0.0, progresso))
        self.progress_geracao.grid()
        self.progress_geracao.set(progresso)
        if mensagem:
            self.label_status.configure(text=str(mensagem), text_color=FG_MUTED)

    def exibir_erro(self, mensagem: str) -> None:
        self._conexao_em_andamento = False
        self._geracao_em_andamento = False
        self.label_status.configure(text=str(mensagem), text_color=ERROR)
        self._atualizar_estados_controles()

    def exibir_sucesso_geracao(
        self, mensagem: str = "Relatório salvo com sucesso."
    ) -> None:
        self._geracao_em_andamento = False
        self.progress_geracao.grid()
        self.progress_geracao.set(1)
        self.label_status.configure(text=str(mensagem), text_color=SUCCESS)
        self._atualizar_estados_controles()

    def preparar_novo_relatorio(self) -> None:
        """Volta ao escopo mantendo a empresa e restaurando escolhas padrão."""

        self._geracao_em_andamento = False
        self.var_todos_setores.set(True)
        for variavel in self._departamento_variaveis.values():
            variavel.set(False)
        self.var_gerar_saldo.set(True)
        self.var_gerar_resumo.set(True)
        self.var_gerar_ranking.set(True)
        inicio, fim = periodo_padrao()
        self._preencher_entry(self.entry_data_inicial, inicio)
        self._preencher_entry(self.entry_data_final, fim)
        self.progress_geracao.set(0)
        self.progress_geracao.grid_remove()
        self._mostrar_etapa(ETAPA_ESCOPO)
        self._atualizar_status_setores()
        self._atualizar_estados_controles()

    def _atualizar_estados_controles(self) -> None:
        bloqueado = self._conexao_em_andamento or self._geracao_em_andamento
        estado_login = "disabled" if bloqueado else "normal"
        self.entry_email.configure(state=estado_login)
        self.entry_senha.configure(state=estado_login)
        self.check_lembrar.configure(state=estado_login)
        self.btn_esqueci_senha.configure(state=estado_login)
        self.btn_conectar.configure(
            state=estado_login,
            text="Conectando..." if self._conexao_em_andamento else "Conectar",
        )
        self.combo_dominio.configure(state="disabled" if bloqueado else "readonly")
        self.btn_confirmar_dominio.configure(
            state="disabled" if bloqueado else "normal"
        )

        escopo_ativo = (
            self._catalogo_carregado
            and self._etapa == ETAPA_ESCOPO
            and not bloqueado
        )
        self.combo_empresa.configure(state="readonly" if escopo_ativo else "disabled")
        self.check_todos_setores.configure(
            state="normal" if escopo_ativo and self._departamentos_visiveis else "disabled"
        )
        individuais_ativos = escopo_ativo and not self.var_todos_setores.get()
        for check in self._departamento_checks:
            check.configure(state="normal" if individuais_ativos else "disabled")
        estado_escopo = "normal" if escopo_ativo else "disabled"
        self.entry_data_inicial.configure(state=estado_escopo)
        self.entry_data_final.configure(state=estado_escopo)
        self.btn_calendario_inicial.configure(state=estado_escopo)
        self.btn_calendario_final.configure(state=estado_escopo)
        self.calendario.definir_habilitado(escopo_ativo)
        for check in self.check_abas:
            check.configure(state=estado_escopo)
        pode_gerar = escopo_ativo and bool(self._departamentos_visiveis)
        self.btn_gerar.configure(
            state="normal" if pode_gerar else "disabled",
            text="Gerando..." if self._geracao_em_andamento else "Gerar relatório",
        )
        self.btn_voltar.configure(state="disabled" if bloqueado else "normal")


__all__ = [
    "ETAPA_DOMINIO",
    "ETAPA_ESCOPO",
    "ETAPA_LOGIN",
    "RHID_FORGOT_PASSWORD_URL",
    "RhidPage",
    "TODOS_OS_SETORES",
    "catalogar_por_id",
    "departamentos_da_empresa",
    "periodo_padrao",
    "validar_periodo",
]
