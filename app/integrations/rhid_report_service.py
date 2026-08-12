"""Orquestra o CSV gerado pelo RHiD e o pipeline validado do FAS Jornada."""

from __future__ import annotations

import csv
import io
import os
import re
import tempfile
import unicodedata
from dataclasses import dataclass
from datetime import date
from typing import Callable

from app.core.logger import logger
from app.integrations.rhid_client import RhidApiError, RhidClient
from app.services.workbook_pipeline_service import processar_arquivo


@dataclass(frozen=True, slots=True)
class RhidReportPlan:
    company_id: int | None
    department_id: int | tuple[int, ...] | None
    company_label: str
    department_label: str
    data_inicial: date
    data_final: date
    caminho_saida: str
    gerar_saldo: bool = True
    gerar_resumo: bool = True
    gerar_ranking: bool = True


_CABECALHOS_RHID = (
    "Nome do funcionário",
    "Número de matrícula",
    "Nome do departamento",
    "CPF do funcionário",
    "Total Normais",
    "Total Trabalhado",
    "Horas Previstas",
    "Dia Falta",
    "Falta e Atraso",
    "Abono",
    "Extras Total",
    "Banco Total",
    "Banco Saldo",
)


def _normalizar(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    return " ".join(
        "".join(letra for letra in texto if not unicodedata.combining(letra))
        .casefold()
        .split()
    )


def _identificador_matricula(valor) -> str:
    return re.sub(r"[^0-9a-z]+", "", _normalizar(valor))


def _validar_csv_consolidado(conteudo_csv: bytes) -> None:
    """Impede que um relatório diário seja tratado como uma linha por pessoa."""

    try:
        texto = conteudo_csv.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise RhidApiError("O CSV do RHiD não está em UTF-8.") from exc

    leitor = csv.reader(io.StringIO(texto, newline=""), delimiter=";")
    try:
        cabecalhos = next(leitor)
    except StopIteration as exc:
        raise RhidApiError("O RHiD gerou um CSV vazio.") from exc

    indices = {_normalizar(valor): indice for indice, valor in enumerate(cabecalhos)}
    ausentes = [nome for nome in _CABECALHOS_RHID if _normalizar(nome) not in indices]
    if ausentes:
        raise RhidApiError(
            "O RHiD não retornou o relatório consolidado. Colunas ausentes: "
            + ", ".join(ausentes)
            + "."
        )

    indice_matricula = indices[_normalizar("Número de matrícula")]
    indice_cpf = indices[_normalizar("CPF do funcionário")]
    identidades = set()
    matricula_por_cpf = {}
    cpf_por_matricula = {}
    quantidade = 0

    for numero_linha, linha in enumerate(leitor, start=2):
        if not linha or not any(str(valor).strip() for valor in linha):
            continue
        if len(linha) < len(cabecalhos):
            raise RhidApiError(f"O CSV do RHiD está incompleto na linha {numero_linha}.")

        matricula = _identificador_matricula(linha[indice_matricula])
        cpf = re.sub(r"\D", "", str(linha[indice_cpf] or ""))
        if not matricula and not cpf:
            raise RhidApiError(
                f"O CSV do RHiD possui funcionário sem matrícula e CPF na linha {numero_linha}."
            )

        if matricula and cpf:
            cpf_anterior = cpf_por_matricula.setdefault(matricula, cpf)
            matricula_anterior = matricula_por_cpf.setdefault(cpf, matricula)
            if cpf_anterior != cpf or matricula_anterior != matricula:
                raise RhidApiError("O RHiD retornou identificações conflitantes de funcionários.")

        identidade = ("matricula", matricula) if matricula else ("cpf", cpf)
        if identidade in identidades:
            raise RhidApiError(
                "O RHiD retornou mais de uma linha para a mesma matrícula. "
                "O Excel não foi gerado para evitar somar o saldo duas vezes."
            )
        identidades.add(identidade)
        quantidade += 1

    if quantidade == 0:
        raise RhidApiError("O RHiD não retornou funcionários para o período selecionado.")


def processar_relatorio_rhid(
    cliente: RhidClient,
    plano: RhidReportPlan,
    ao_progresso: Callable[[int], None] | None = None,
) -> dict:
    """Baixa o CSV internamente e o entrega, sem adaptação, ao pipeline atual."""

    reportar = ao_progresso or (lambda _valor: None)
    conteudo_csv = cliente.gerar_relatorio_csv(
        plano.company_id,
        plano.department_id,
        plano.data_inicial,
        plano.data_final,
        ao_progresso=lambda valor: reportar(min(80, round(valor * 0.8))),
    )
    _validar_csv_consolidado(conteudo_csv)

    pasta_saida = os.path.dirname(os.path.abspath(plano.caminho_saida))
    caminho_temporario = ""
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            prefix=".clockhack_rhid_",
            suffix=".csv",
            dir=pasta_saida,
            delete=False,
        ) as arquivo_temporario:
            arquivo_temporario.write(conteudo_csv)
            caminho_temporario = arquivo_temporario.name

        reportar(85)
        resultado = processar_arquivo(
            caminho_temporario,
            plano.caminho_saida,
            "Todos",
            gerar_saldo=plano.gerar_saldo,
            gerar_ranking=plano.gerar_ranking,
            gerar_resumo=plano.gerar_resumo,
        )
        resultado["tipo_entrada"] = "RHID"
        resultado["departamento"] = plano.department_label
        resultado["empresa_rhid"] = plano.company_label
        resultado["periodo_rhid"] = (
            f"{plano.data_inicial.isoformat()} a {plano.data_final.isoformat()}"
        )
        reportar(100)
        return resultado
    finally:
        if caminho_temporario:
            try:
                os.remove(caminho_temporario)
            except FileNotFoundError:
                pass
            except OSError:
                logger.warning("Não foi possível remover o CSV temporário do RHiD.")
