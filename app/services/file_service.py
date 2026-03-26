from pathlib import Path
from datetime import datetime

def nome_curto(caminho_arquivo: str) -> str:
    return Path(caminho_arquivo).name

def sugerir_nome_saida(caminho_arquivo: str, departamento: str) -> str:
    base = Path(caminho_arquivo).stem
    depto = "todos" if not departamento or departamento == "Todos" else departamento.strip().replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base}_{depto}_tratado_{timestamp}.xlsx"

def tipo_arquivo(caminho_arquivo: str) -> str:
    return Path(caminho_arquivo).suffix.lower().replace(".", "").upper()
