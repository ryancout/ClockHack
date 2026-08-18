import csv
import os

from openpyxl import Workbook

from app.core.exceptions import ArquivoInvalidoError


def carregar_workbook(caminho_arquivo):
    extensao = os.path.splitext(caminho_arquivo)[1].lower()
    if extensao == ".csv":
        wb = Workbook()
        ws = wb.active
        if ws is None:
            raise ArquivoInvalidoError("Não foi possível criar a planilha de trabalho.")
        with open(caminho_arquivo, "r", encoding="utf-8-sig", newline="") as arquivo_csv:
            leitor = csv.reader(arquivo_csv, delimiter=";")
            for linha in leitor:
                ws.append(linha)
        return wb
    raise ArquivoInvalidoError("Formato de arquivo não suportado. Selecione um arquivo CSV (.csv).")
