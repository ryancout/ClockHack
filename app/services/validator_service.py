from app.core.config import COLUNAS_OBRIGATORIAS, MIN_FUNCIONARIOS_ALERTA
from app.core.exceptions import ColunaObrigatoriaError, ValidacaoNegocioError

def mapear_colunas(ws):
    colunas = {}
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(row=1, column=col).value
        if valor:
            colunas[str(valor).strip()] = col
    return colunas

def validar_colunas(colunas):
    faltando = [c for c in COLUNAS_OBRIGATORIAS if c not in colunas]
    if faltando:
        raise ColunaObrigatoriaError(f"Colunas obrigatórias não encontradas: {', '.join(faltando)}")

def validar_resultado(quantidade_funcionarios):
    if quantidade_funcionarios < MIN_FUNCIONARIOS_ALERTA:
        raise ValidacaoNegocioError(f"Arquivo suspeito: apenas {quantidade_funcionarios} funcionários encontrados.")
