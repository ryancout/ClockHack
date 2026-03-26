from app.core.config import COLUNAS_OBRIGATORIAS, MIN_FUNCIONARIOS_ALERTA
from app.core.exceptions import ColunaObrigatoriaError, ValidacaoNegocioError

def mapear_colunas(ws):
    colunas = {}

    for row in range(1, 10):
        temp = {}
        for col in range(1, ws.max_column + 1):
            valor = ws.cell(row=row, column=col).value
            if valor:
                chave_original = str(valor).strip()
                chave_normalizada = chave_original.lower()
                temp[chave_normalizada] = col

        if "nome do departamento" in temp:
            return temp

    raise Exception("Cabeçalho não encontrado. Verifique a planilha.")

def validar_colunas(colunas):
    faltando = [c for c in COLUNAS_OBRIGATORIAS if c.lower() not in colunas]
    if faltando:
        raise ColunaObrigatoriaError(f"Colunas obrigatórias não encontradas: {', '.join(faltando)}")

def validar_resultado(quantidade_funcionarios):
    if quantidade_funcionarios < MIN_FUNCIONARIOS_ALERTA:
        raise ValidacaoNegocioError(f"Arquivo suspeito: apenas {quantidade_funcionarios} funcionários encontrados.")
