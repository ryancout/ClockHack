import os
from app.core.config import (
    COLUNAS_OBRIGATORIAS,
    EXTENSOES_ENTRADA_SUPORTADAS,
    MAX_FILE_SIZE_MB,
    MIN_FUNCIONARIOS_ALERTA,
)
from app.core.exceptions import ArquivoInvalidoError, ColunaObrigatoriaError, ValidacaoNegocioError
from app.services.file_service import nome_curto


def validar_arquivo_entrada(caminho):
    if not caminho:
        raise ArquivoInvalidoError("Nenhum arquivo foi informado.")
    if not os.path.exists(caminho):
        raise ArquivoInvalidoError(f"Arquivo não encontrado: {nome_curto(caminho)}")
    if os.path.isdir(caminho):
        raise ArquivoInvalidoError(f"Selecione um arquivo válido: {nome_curto(caminho)}")

    extensao = os.path.splitext(caminho)[1].lower()
    if extensao not in EXTENSOES_ENTRADA_SUPORTADAS:
        raise ArquivoInvalidoError(
            f"Formato não suportado para {nome_curto(caminho)}. Selecione um arquivo CSV (.csv)."
        )

    tamanho_bytes = os.path.getsize(caminho)
    if tamanho_bytes <= 0:
        raise ArquivoInvalidoError(f"O arquivo {nome_curto(caminho)} está vazio.")

    tamanho_mb = tamanho_bytes / (1024 * 1024)
    if tamanho_mb > MAX_FILE_SIZE_MB:
        raise ArquivoInvalidoError(
            f"O arquivo {nome_curto(caminho)} excede o limite de {MAX_FILE_SIZE_MB} MB."
        )


def mapear_colunas(ws):
    colunas = {}
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(row=1, column=col).value
        if valor:
            chave_original = str(valor).strip()
            chave_normalizada = chave_original.lower()
            colunas[chave_normalizada] = col

    if "nome do departamento" in colunas:
        return colunas

    raise ArquivoInvalidoError(
        "Cabeçalho não encontrado na primeira linha. Verifique o formato do relatório CSV."
    )


def validar_colunas(colunas):
    faltando = [c for c in COLUNAS_OBRIGATORIAS if c.lower() not in colunas]
    if faltando:
        raise ColunaObrigatoriaError(
            "Colunas obrigatórias não encontradas: " + ", ".join(faltando)
        )


def validar_resultado(quantidade_funcionarios):
    if quantidade_funcionarios < MIN_FUNCIONARIOS_ALERTA:
        raise ValidacaoNegocioError(
            f"Arquivo suspeito: apenas {quantidade_funcionarios} funcionário(s) encontrado(s)."
        )
