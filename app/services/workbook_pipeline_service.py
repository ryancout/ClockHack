"""Orquestra a transformação do CSV no relatório Excel final."""

import os

from app.domain import (
    RegistroFuncionario,
    formatar_horas,
    obter_final_matricula,
    para_minutos,
)
from app.reports import criar_aba_ranking, criar_aba_resumo, criar_aba_saldo
from app.services.calculator_service import calcular_totais
from app.services.filter_service import aplicar_filtro_departamento, listar_departamentos
from app.services.reader_service import carregar_workbook
from app.services.validator_service import (
    mapear_colunas,
    validar_arquivo_entrada,
    validar_colunas,
    validar_resultado,
)
from app.services.worksheet_formatting_service import formatar_aba_principal
from app.services.writer_service import escrever_resultado


def obter_departamentos(caminho_arquivo):
    """Lista os departamentos disponíveis em um arquivo de entrada válido."""
    validar_arquivo_entrada(caminho_arquivo)
    wb = carregar_workbook(caminho_arquivo)
    ws = wb.active
    colunas = mapear_colunas(ws)
    validar_colunas(colunas)
    return listar_departamentos(ws, colunas["nome do departamento"])


def _obter_coluna_faltas(colunas):
    """Localiza a primeira coluna de faltas aceita pelo relatório."""
    for nome_coluna, indice in colunas.items():
        if nome_coluna == "faltas" or "falta" in nome_coluna:
            return indice
    return None


def _extrair_registros(ws, colunas):
    """Converte as linhas da planilha em registros de domínio."""
    col_nome = colunas["nome do funcionário"]
    col_matricula = colunas["número de matrícula"]
    col_depart = colunas["nome do departamento"]
    col_bt = colunas["banco total"]
    col_bs = colunas["banco saldo"]
    col_faltas = _obter_coluna_faltas(colunas)

    registros = []
    for row in range(2, ws.max_row + 1):
        nome = ws.cell(row=row, column=col_nome).value
        matricula = ws.cell(row=row, column=col_matricula).value
        final_matricula = obter_final_matricula(matricula)
        departamento = ws.cell(row=row, column=col_depart).value
        banco_total_valor = ws.cell(row=row, column=col_bt).value
        banco_saldo_valor = ws.cell(row=row, column=col_bs).value
        faltas = ws.cell(row=row, column=col_faltas).value if col_faltas else ""

        if nome in (None, ""):
            continue

        registros.append(
            RegistroFuncionario(
                nome=nome,
                final_matricula=final_matricula,
                departamento=departamento,
                banco_total_minutos=para_minutos(banco_total_valor),
                banco_saldo_minutos=para_minutos(banco_saldo_valor),
                faltas=faltas,
            )
        )

    return registros


def _mascarar_matriculas_na_planilha(ws, col_matricula):
    """Mantém somente os três dígitos finais da matrícula na saída."""
    ws.cell(row=1, column=col_matricula, value="Final da matrícula")
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=col_matricula)
        cell.value = obter_final_matricula(cell.value)
        cell.number_format = "@"


def _atualizar_abas_auxiliares(
    wb,
    registros,
    *,
    gerar_saldo,
    gerar_ranking,
    gerar_resumo,
):
    """Cria ou remove as abas opcionais conforme a seleção do usuário."""
    if gerar_saldo:
        criar_aba_saldo(wb, registros)
    elif "SALDO" in wb.sheetnames:
        del wb["SALDO"]

    if gerar_ranking:
        criar_aba_ranking(wb, registros)
    elif "RANKING" in wb.sheetnames:
        del wb["RANKING"]

    if gerar_resumo:
        criar_aba_resumo(wb, registros)
    elif "RESUMO" in wb.sheetnames:
        del wb["RESUMO"]


def processar_arquivo(
    caminho_arquivo,
    caminho_saida,
    departamento="Todos",
    gerar_saldo=True,
    gerar_ranking=True,
    gerar_resumo=True,
):
    """Processa um CSV e grava o relatório Excel no caminho informado."""
    validar_arquivo_entrada(caminho_arquivo)
    wb = carregar_workbook(caminho_arquivo)
    ws = wb.active

    colunas = mapear_colunas(ws)
    validar_colunas(colunas)

    col_nome = colunas["nome do funcionário"]
    col_matricula = colunas["número de matrícula"]
    col_depart = colunas["nome do departamento"]
    col_bt = colunas["banco total"]
    col_bs = colunas["banco saldo"]

    aplicar_filtro_departamento(ws, col_depart, departamento)
    registros = _extrair_registros(ws, colunas)
    _mascarar_matriculas_na_planilha(ws, col_matricula)

    resultado_calc = calcular_totais(ws, col_nome, col_bt, col_bs)
    validar_resultado(resultado_calc["quantidade_funcionarios"])

    escrever_resultado(
        ws,
        col_nome,
        col_bt,
        col_bs,
        resultado_calc["soma_bt"],
        resultado_calc["soma_bs"],
    )

    _atualizar_abas_auxiliares(
        wb,
        registros,
        gerar_saldo=gerar_saldo,
        gerar_ranking=gerar_ranking,
        gerar_resumo=gerar_resumo,
    )
    formatar_aba_principal(ws)
    wb.save(caminho_saida)

    return {
        "caminho_saida": caminho_saida,
        "banco_total": formatar_horas(resultado_calc["soma_bt"]),
        "banco_saldo": formatar_horas(resultado_calc["soma_bs"]),
        "quantidade_funcionarios": resultado_calc["quantidade_funcionarios"],
        "tipo_entrada": os.path.splitext(caminho_arquivo)[1].lower().replace(".", "").upper(),
        "departamento": departamento,
        "gerou_ranking": gerar_ranking,
        "gerou_saldo": gerar_saldo,
        "gerou_resumo": gerar_resumo,
    }
