from pathlib import Path

def nome_curto(caminho_arquivo: str) -> str:
    return Path(caminho_arquivo).name

def sugerir_nome_saida(caminho_arquivo: str, departamento: str) -> str:
    base = Path(caminho_arquivo).stem
    depto = "todos" if not departamento or departamento == "Todos" else departamento.strip().replace(" ", "_")
    return f"{base}_{depto}_tratado.xlsx"

def tipo_arquivo(caminho_arquivo: str) -> str:
    return Path(caminho_arquivo).suffix.lower().replace(".", "").upper()
