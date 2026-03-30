import csv
import os
from openpyxl import Workbook, load_workbook
from app.core.exceptions import ArquivoInvalidoError

def carregar_workbook(caminho_arquivo):
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    if extensao == ".xlsx":
        return load_workbook(caminho_arquivo)
    if extensao == ".csv":
        wb = Workbook()
        ws = wb.active
        with open(caminho_arquivo, "r", encoding="utf-8-sig", newline="") as arquivo_csv:
            leitor = csv.reader(arquivo_csv, delimiter=";")
            for linha in leitor:
                ws.append(linha)
        return wb
    raise ArquivoInvalidoError("Formato de arquivo não suportado. Envie .xlsx ou .csv.")
