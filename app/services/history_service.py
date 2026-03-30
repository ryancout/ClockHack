import json
import os
import tempfile
from datetime import datetime
from pathlib import Path
from app.core.config import DATA_DIR, HISTORY_FILE, MAX_HISTORICO

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
    fd, temp_path = tempfile.mkstemp(prefix=HISTORY_FILE.stem + "_", suffix=".tmp", dir=str(HISTORY_FILE.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, HISTORY_FILE)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def _sanitize_item(item: dict) -> dict:
    safe = dict(item)
    if "arquivo_origem" in safe:
        safe["arquivo_origem"] = Path(str(safe["arquivo_origem"])).name
    if "arquivo_saida" in safe:
        safe["arquivo_saida"] = Path(str(safe["arquivo_saida"])).name
    return safe


def registrar_historico(item: dict):
    data = _load()
    safe_item = _sanitize_item(item)
    safe_item["data_execucao"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    data.insert(0, safe_item)
    data = data[:MAX_HISTORICO]
    _save(data)


def ultimos_processamentos(limit=8):
    return _load()[:limit]
