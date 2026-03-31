from pathlib import Path
import re


def nome_curto(caminho_arquivo: str) -> str:
    return Path(caminho_arquivo).name


def _slug_texto(texto: str) -> str:
    texto = str(texto or "").strip()
    texto = re.sub(r"[^\w\-. ]+", "", texto, flags=re.UNICODE)
    return texto.replace(" ", "_") or "arquivo"


def sugerir_nome_saida(caminho_arquivo: str, departamento: str) -> str:
    base = Path(caminho_arquivo).stem
    depto = "todos" if not departamento or departamento == "Todos" else _slug_texto(departamento)
    return f"{_slug_texto(base)}_{depto}_tratado.xlsx"


def tipo_arquivo(caminho_arquivo: str) -> str:
    return Path(caminho_arquivo).suffix.lower().replace(".", "").upper()


def garantir_extensao_xlsx(caminho_arquivo: str) -> str:
    caminho = str(caminho_arquivo or "").strip()
    if not caminho.lower().endswith(".xlsx"):
        caminho += ".xlsx"
    return caminho
