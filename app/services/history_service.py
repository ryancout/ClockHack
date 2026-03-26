import json
from datetime import datetime
from app.core.config import DATA_DIR, HISTORY_FILE

DATA_DIR.mkdir(parents=True, exist_ok=True)

def _load():
    if not HISTORY_FILE.exists():
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def _save(data):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def registrar_historico(item: dict):
    data = _load()
    item["data_execucao"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    data.insert(0, item)
    data = data[:100]
    _save(data)

def ultimos_processamentos(limit=8):
    return _load()[:limit]
