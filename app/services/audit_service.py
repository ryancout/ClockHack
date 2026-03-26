import json
from datetime import datetime
from app.core.config import DATA_DIR, AUDIT_FILE

DATA_DIR.mkdir(parents=True, exist_ok=True)

def _load():
    if not AUDIT_FILE.exists():
        return []
    try:
        with open(AUDIT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def registrar_evento(acao: str, detalhes: dict):
    data = _load()
    data.insert(0, {
        "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "acao": acao,
        "detalhes": detalhes
    })
    data = data[:500]
    with open(AUDIT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
