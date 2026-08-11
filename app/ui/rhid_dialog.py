"""Janela de conexão e seleção do escopo disponível no RHiD."""

from __future__ import annotations

from datetime import date
from typing import Callable

import customtkinter as ctk

from app.core.config import BG_BOX, BORDER, ERROR, FG_MUTED, FG_TEXT, PRIMARY, SUCCESS


TODOS_OS_SETORES = "Todos os setores"
TODAS_AS_EMPRESAS = "Todas as empresas"


def validar_periodo(data_inicial: str, data_final: str) -> tuple[str, str]:
    """Valida o período e devolve datas ISO prontas para a integração."""
    try:
        inicio = date.fromisoformat(str(data_inicial).strip())
        fim = date.fromisoformat(str(data_final).strip())
    except (TypeError, ValueError) as exc:
        raise ValueError("Use as datas no formato AAAA-MM-DD.") from exc

    if fim < inicio:
        raise ValueError("A data final não pode ser anterior à data inicial.")
    if (fim - inicio).days > 31:
        raise ValueError("O período do relatório não pode ultrapassar 31 dias.")
    return inicio.isoformat(), fim.isoformat()


def periodo_padrao(hoje: date | None = None) -> tuple[str, str]:
    """Usa o mês corrente, do primeiro dia até hoje, como período inicial."""
    hoje = hoje or date.today()
    return hoje.replace(day=1).isoformat(), hoje.isoformat()


def _chave_id(valor) -> str:
    return str(valor).strip()


def _pertence_a_empresa(departamento, empresa_id) -> bool:
    return _chave_id(getattr(departamento, "company_id", "")) == _chave_id(empresa_id)


class RhidConnectionDialog(ctk.CTkToplevel):
    def __init__(self, parent, ao_conectar, ao_gerar: Callable | None = None):
        super().__init__(parent)
        self._ao_conectar = ao_conectar
        self._ao_gerar = ao_gerar
        self._empresas = ()
        self._departamentos = ()
        self._empresa_por_id = {}
        self._empresa_id_por_rotulo = {TODAS_AS_EMPRESAS: None}
        self._departamento_por_id = {}
        self._departamento_id_por_rotulo = {TODOS_OS_SETORES: None}
        self._catalogo_carregado = False
        self._conexao_em_andamento = False
        self._geracao_em_andamento = False
        self._dominio_somente_leitura = False

        self.title("Conectar ao RHiD")
        self.geometry("540x620")
        self.minsize(500, 560)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self,
            text="Integração RHiD",
            font=("Segoe UI", 22, "bold"),
            text_color=FG_TEXT,
        ).grid(row=0, column=0, sticky="w", padx=24, pady=(22, 4))
        ctk.CTkLabel(
            self,
            text="A senha é usada somente para criar o token e não é salva.",
            text_color=FG_MUTED,
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 14))

        credenciais = ctk.CTkFrame(
            self, fg_color=BG_BOX, corner_radius=12, border_width=1, border_color=BORDER
        )
        self.credenciais = credenciais
        credenciais.grid(row=2, column=0, sticky="ew", padx=24)
        credenciais.grid_columnconfigure(0, weight=1)

        self.entry_email = self._campo(credenciais, "E-mail", 0)
        self.entry_senha = self._campo(credenciais, "Senha", 2, show="●")
        ctk.CTkLabel(
            credenciais,
            text="Cliente / domínio",
            text_color=FG_MUTED,
        ).grid(row=4, column=0, sticky="w", padx=14, pady=(10, 4))
        self.combo_dominio = ctk.CTkComboBox(
            credenciais,
            values=["Automático"],
            state="normal",
        )
        self.combo_dominio.grid(row=5, column=0, sticky="ew", padx=14)
        self.combo_dominio.set("Automático")
        self.btn_conectar = ctk.CTkButton(
            credenciais,
            text="Conectar",
            fg_color=PRIMARY,
            command=self._conectar,
        )
        self.btn_conectar.grid(row=6, column=0, sticky="ew", padx=14, pady=(8, 14))

        self.label_status = ctk.CTkLabel(
            self,
            text="Aguardando conexão.",
            text_color=FG_MUTED,
            justify="left",
            anchor="w",
            wraplength=470,
        )
        self.label_status.grid(row=3, column=0, sticky="ew", padx=24, pady=(12, 4))

        self.escopo = ctk.CTkFrame(
            self, fg_color=BG_BOX, corner_radius=12, border_width=1, border_color=BORDER
        )
        self.escopo.grid(row=4, column=0, sticky="ew", padx=24, pady=(6, 0))
        self.escopo.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.escopo, text="Empresa / unidade (opcional)", text_color=FG_MUTED).grid(
            row=0, column=0, sticky="w", padx=14, pady=(12, 4)
        )
        self.combo_empresa = ctk.CTkComboBox(
            self.escopo,
            values=["Conecte-se primeiro"],
            state="disabled",
            command=self._empresa_alterada,
        )
        self.combo_empresa.grid(row=1, column=0, sticky="ew", padx=14)
        ctk.CTkLabel(self.escopo, text="Setor / departamento", text_color=FG_MUTED).grid(
            row=2, column=0, sticky="w", padx=14, pady=(10, 4)
        )
        self.combo_departamento = ctk.CTkComboBox(
            self.escopo,
            values=[TODOS_OS_SETORES],
            state="disabled",
        )
        self.combo_departamento.grid(row=3, column=0, sticky="ew", padx=14)

        periodo = ctk.CTkFrame(self.escopo, fg_color="transparent")
        periodo.grid(row=4, column=0, sticky="ew", padx=14, pady=(10, 10))
        periodo.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkLabel(periodo, text="Data inicial (AAAA-MM-DD)", text_color=FG_MUTED).grid(
            row=0, column=0, sticky="w", padx=(0, 5), pady=(0, 4)
        )
        ctk.CTkLabel(periodo, text="Data final (AAAA-MM-DD)", text_color=FG_MUTED).grid(
            row=0, column=1, sticky="w", padx=(5, 0), pady=(0, 4)
        )
        self.entry_data_inicial = ctk.CTkEntry(periodo)
        self.entry_data_inicial.grid(row=1, column=0, sticky="ew", padx=(0, 5))
        self.entry_data_final = ctk.CTkEntry(periodo)
        self.entry_data_final.grid(row=1, column=1, sticky="ew", padx=(5, 0))
        data_inicial, data_final = periodo_padrao()
        self.entry_data_inicial.insert(0, data_inicial)
        self.entry_data_final.insert(0, data_final)

        self.progress_geracao = ctk.CTkProgressBar(self.escopo, height=10)
        self.progress_geracao.grid(row=5, column=0, sticky="ew", padx=14, pady=(0, 10))
        self.progress_geracao.set(0)

        self.btn_gerar = ctk.CTkButton(
            self.escopo,
            text="Gerar Excel",
            fg_color=SUCCESS,
            command=self._gerar_excel,
            state="disabled",
        )
        self.btn_gerar.grid(row=6, column=0, sticky="ew", padx=14, pady=(0, 14))

        self.btn_fechar = ctk.CTkButton(
            self,
            text="Fechar",
            fg_color="transparent",
            border_width=1,
            border_color=BORDER,
            text_color=FG_TEXT,
            command=self._fechar,
        )
        self.btn_fechar.grid(row=5, column=0, sticky="ew", padx=24, pady=(10, 20))

        self.escopo.grid_remove()
        self._atualizar_estados_controles()
        self.after(100, self.entry_email.focus_set)

    @staticmethod
    def _campo(parent, titulo, row, **kwargs):
        ctk.CTkLabel(parent, text=titulo, text_color=FG_MUTED).grid(
            row=row, column=0, sticky="w", padx=14, pady=(10, 4)
        )
        campo = ctk.CTkEntry(parent, **kwargs)
        campo.grid(row=row + 1, column=0, sticky="ew", padx=14)
        return campo

    def _conectar(self):
        senha = self.entry_senha.get()
        self.entry_senha.delete(0, "end")
        self._ao_conectar(
            self.entry_email.get(),
            senha,
            "" if self.combo_dominio.get() == "Automático" else self.combo_dominio.get(),
        )

    def configurar_ao_gerar(self, callback: Callable | None):
        """Configura o callback que receberá empresa, setor e período selecionados."""
        self._ao_gerar = callback
        self._atualizar_estados_controles()

    def definir_ocupado(self, ocupado: bool):
        self._conexao_em_andamento = ocupado
        self._atualizar_estados_controles()
        if ocupado:
            self.label_status.configure(text="Conectando ao RHiD...", text_color=FG_MUTED)

    def definir_geracao_ocupada(self, ocupado: bool):
        self._geracao_em_andamento = ocupado
        if ocupado:
            self.progress_geracao.set(0)
        self._atualizar_estados_controles()
        if ocupado:
            self.label_status.configure(text="Gerando arquivo Excel...", text_color=FG_MUTED)

    def _atualizar_estados_controles(self):
        bloqueado = self._conexao_em_andamento or self._geracao_em_andamento
        estado_credencial = "disabled" if bloqueado else "normal"
        self.entry_email.configure(state=estado_credencial)
        self.entry_senha.configure(state=estado_credencial)
        self.combo_dominio.configure(
            state=(
                "disabled"
                if bloqueado
                else "readonly" if self._dominio_somente_leitura else "normal"
            )
        )
        self.btn_conectar.configure(
            state=estado_credencial,
            text="Conectando..." if self._conexao_em_andamento else "Conectar",
        )

        estado_escopo = (
            "readonly" if self._catalogo_carregado and not bloqueado else "disabled"
        )
        self.combo_empresa.configure(state=estado_escopo)
        self.combo_departamento.configure(state=estado_escopo)
        estado_data = "normal" if self._catalogo_carregado and not bloqueado else "disabled"
        self.entry_data_inicial.configure(state=estado_data)
        self.entry_data_final.configure(state=estado_data)

        pode_gerar = self._catalogo_carregado and self._ao_gerar is not None and not bloqueado
        self.btn_gerar.configure(
            state="normal" if pode_gerar else "disabled",
            text="Gerando..." if self._geracao_em_andamento else "Gerar Excel",
        )
        self.btn_fechar.configure(state="disabled" if bloqueado else "normal")

    def _fechar(self):
        if self._conexao_em_andamento or self._geracao_em_andamento:
            return
        self.destroy()

    def exibir_catalogo(self, empresas, departamentos):
        self.definir_ocupado(False)
        self.credenciais.grid_remove()
        self.escopo.grid()
        self._empresas = tuple(empresas)
        self._departamentos = tuple(departamentos)
        self._empresa_por_id = {}
        self._empresa_id_por_rotulo = {TODAS_AS_EMPRESAS: None}

        for empresa in self._empresas:
            empresa_id = getattr(empresa, "id", None)
            if empresa_id is None:
                continue
            chave = _chave_id(empresa_id)
            if not chave or chave in self._empresa_por_id:
                continue
            self._empresa_por_id[chave] = empresa
            nome = str(getattr(empresa, "label", "") or f"Empresa {empresa_id}").strip()
            self._empresa_id_por_rotulo[f"{nome} — ID {empresa_id}"] = empresa_id

        self._catalogo_carregado = bool(self._departamentos)
        rotulos = list(self._empresa_id_por_rotulo)
        self.combo_empresa.configure(values=rotulos)
        self.combo_empresa.set(TODAS_AS_EMPRESAS)
        self._empresa_alterada(TODAS_AS_EMPRESAS)
        self._atualizar_estados_controles()

        if self._catalogo_carregado:
            self.label_status.configure(
                text=(
                    f"Conectado. {len(self._departamentos)} setor(es) com funcionários ativos."
                ),
                text_color=SUCCESS,
            )
        else:
            self.label_status.configure(
                text="Conectado, mas nenhum setor possui funcionário ativo.",
                text_color=ERROR,
            )

    def exibir_erro(self, mensagem: str):
        self.definir_ocupado(False)
        if self._geracao_em_andamento:
            self.definir_geracao_ocupada(False)
        self.label_status.configure(text=mensagem, text_color=ERROR)

    def exibir_erro_geracao(self, mensagem: str):
        self.definir_geracao_ocupada(False)
        self.label_status.configure(text=mensagem, text_color=ERROR)

    def exibir_sucesso_geracao(self, mensagem: str = "Arquivo Excel gerado com sucesso."):
        self.definir_geracao_ocupada(False)
        self.progress_geracao.set(1)
        self.label_status.configure(text=mensagem, text_color=SUCCESS)

    def atualizar_progresso(self, valor, mensagem: str = ""):
        try:
            progresso = float(valor)
        except (TypeError, ValueError):
            progresso = 0.0
        if progresso > 1:
            progresso /= 100
        progresso = max(0.0, min(1.0, progresso))
        self.progress_geracao.set(progresso)
        if mensagem:
            self.label_status.configure(text=str(mensagem), text_color=FG_MUTED)

    def exibir_dominios(self, tenants):
        dominios = [tenant.domain for tenant in tenants]
        self._dominio_somente_leitura = bool(dominios)
        self.combo_dominio.configure(values=dominios or ["Automático"])
        self.combo_dominio.set(dominios[0] if dominios else "Automático")
        self._atualizar_estados_controles()

    def _empresa_alterada(self, rotulo):
        empresa_id = self._empresa_id_por_rotulo.get(rotulo)
        self._departamento_por_id = {}
        self._departamento_id_por_rotulo = {TODOS_OS_SETORES: None}

        if rotulo in self._empresa_id_por_rotulo:
            for departamento in self._departamentos:
                departamento_id = getattr(departamento, "id", None)
                if departamento_id is None:
                    continue
                if empresa_id is not None and not _pertence_a_empresa(departamento, empresa_id):
                    continue
                chave = _chave_id(departamento_id)
                if not chave or chave in self._departamento_por_id:
                    continue
                self._departamento_por_id[chave] = departamento
                nome = str(
                    getattr(departamento, "name", "")
                    or f"Departamento {departamento_id}"
                ).strip()
                self._departamento_id_por_rotulo[
                    f"{nome} — ID {departamento_id}"
                ] = departamento_id

        rotulos = list(self._departamento_id_por_rotulo)
        self.combo_departamento.configure(values=rotulos)
        self.combo_departamento.set(TODOS_OS_SETORES)

    def obter_parametros_geracao(self) -> tuple[object, object | None, str, str]:
        rotulo_empresa = self.combo_empresa.get()
        if rotulo_empresa not in self._empresa_id_por_rotulo:
            raise ValueError("Selecione uma empresa válida.")
        empresa_id = self._empresa_id_por_rotulo[rotulo_empresa]

        rotulo_departamento = self.combo_departamento.get()
        if rotulo_departamento not in self._departamento_id_por_rotulo:
            raise ValueError("Selecione um setor válido.")
        departamento_id = self._departamento_id_por_rotulo[rotulo_departamento]
        data_inicial, data_final = validar_periodo(
            self.entry_data_inicial.get(), self.entry_data_final.get()
        )
        return empresa_id, departamento_id, data_inicial, data_final

    def _gerar_excel(self):
        if self._ao_gerar is None or self._geracao_em_andamento:
            return
        try:
            parametros = self.obter_parametros_geracao()
        except ValueError as exc:
            self.exibir_erro_geracao(str(exc))
            return

        try:
            self._ao_gerar(*parametros)
        except Exception as exc:  # callback de UI: apresenta falha síncrona sem fechar a janela
            self.exibir_erro_geracao(str(exc) or "Não foi possível gerar o arquivo Excel.")
