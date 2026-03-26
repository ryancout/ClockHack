import os
from openpyxl.chart import BarChart, Reference

from app.services.reader_service import carregar_workbook
from app.services.validator_service import mapear_colunas, validar_colunas, validar_resultado
from app.services.filter_service import aplicar_filtro_departamento, listar_departamentos
from app.services.calculator_service import calcular_totais
from app.services.writer_service import escrever_resultado
from app.services.time_service import formatar_horas, para_minutos


def obter_departamentos(caminho_arquivo):
    wb = carregar_workbook(caminho_arquivo)
    ws = wb.active

    colunas = mapear_colunas(ws)
    validar_colunas(colunas)

    return listar_departamentos(ws, colunas["Nome do departamento"])


def criar_aba_ranking(wb, dados):
    """
    Cria uma aba com:
    - TOP devedores (mais saldo negativo)
    - TOP hora extra (mais saldo positivo)
    """
    # se já existir, remove e recria
    if "RANKING" in wb.sheetnames:
        del wb["RANKING"]

    ws = wb.create_sheet("RANKING")

    # separar dados
    negativos = sorted(
        [d for d in dados if d["saldo"] < 0],
        key=lambda x: x["saldo"]
    )

    positivos = sorted(
        [d for d in dados if d["saldo"] > 0],
        key=lambda x: x["saldo"],
        reverse=True
    )

    linha = 1

    # título 1
    ws.cell(row=linha, column=1, value="TOP DEVEDORES")
    linha += 1

    ws.append(["Funcionário", "Departamento", "Banco Saldo"])
    linha += 1

    for d in negativos[:10]:
        ws.append([d["nome"], d["departamento"], d["saldo_fmt"]])
        linha += 1

    linha += 2

    # título 2
    ws.cell(row=linha, column=1, value="TOP HORAS EXTRAS")
    linha += 1

    ws.append(["Funcionário", "Departamento", "Banco Saldo"])
    linha += 1

    for d in positivos[:10]:
        ws.append([d["nome"], d["departamento"], d["saldo_fmt"]])

    # ajustar largura básica
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 15


def criar_aba_resumo(wb, dados):
    """
    Cria uma aba RESUMO com:
    - total de horas por departamento
    - gráfico automático
    """
    # se já existir, remove e recria
    if "RESUMO" in wb.sheetnames:
        del wb["RESUMO"]

    ws = wb.create_sheet("RESUMO")

    resumo = {}

    for d in dados:
        departamento = d["departamento"] if d["departamento"] not in (None, "") else "SEM DEPARTAMENTO"
        resumo.setdefault(departamento, 0)
        resumo[departamento] += d["saldo"]

    ws.append(["Departamento", "Horas"])

    for departamento, total_min in resumo.items():
        total_horas = total_min / 60
        ws.append([departamento, total_horas])

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 15

    # gráfico
    if len(resumo) > 0:
        chart = BarChart()
        chart.title = "Horas por Departamento"
        chart.y_axis.title = "Horas"
        chart.x_axis.title = "Departamento"
        chart.height = 8
        chart.width = 16

        data = Reference(ws, min_col=2, min_row=1, max_row=len(resumo) + 1)
        cats = Reference(ws, min_col=1, min_row=2, max_row=len(resumo) + 1)

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        ws.add_chart(chart, "D2")


def processar_arquivo(caminho_arquivo, caminho_saida, departamento="Todos"):
    wb = carregar_workbook(caminho_arquivo)
    ws = wb.active

    colunas = mapear_colunas(ws)
    validar_colunas(colunas)

    col_nome = colunas["Nome do funcionário"]
    col_depart = colunas["Nome do departamento"]
    col_bt = colunas["Banco Total"]
    col_bs = colunas["Banco Saldo"]

    # aplica filtro por departamento antes de processar
    aplicar_filtro_departamento(ws, col_depart, departamento)

    # coleta dados para ranking e resumo
    dados = []
    for row in range(2, ws.max_row + 1):
        nome = ws.cell(row=row, column=col_nome).value
        dep = ws.cell(row=row, column=col_depart).value
        saldo_valor = ws.cell(row=row, column=col_bs).value

        if nome in (None, ""):
            continue

        saldo_min = para_minutos(saldo_valor)

        dados.append({
            "nome": nome,
            "departamento": dep,
            "saldo": saldo_min,
            "saldo_fmt": formatar_horas(saldo_min)
        })

    # cálculos
    resultado_calc = calcular_totais(ws, col_nome, col_bt, col_bs)

    validar_resultado(resultado_calc["quantidade_funcionarios"])

    # escreve linha TOTAL na aba principal
    escrever_resultado(
        ws,
        col_nome,
        col_bt,
        col_bs,
        resultado_calc["soma_bt"],
        resultado_calc["soma_bs"]
    )

    # PASSO 4 entra aqui:
    # criar abas extras antes de salvar
    criar_aba_ranking(wb, dados)
    criar_aba_resumo(wb, dados)

    # salva o arquivo
    wb.save(caminho_saida)

    return {
        "caminho_saida": caminho_saida,
        "banco_total": formatar_horas(resultado_calc["soma_bt"]),
        "banco_saldo": formatar_horas(resultado_calc["soma_bs"]),
        "quantidade_funcionarios": resultado_calc["quantidade_funcionarios"],
        "tipo_entrada": os.path.splitext(caminho_arquivo)[1].lower().replace(".", "").upper(),
        "departamento": departamento
    }