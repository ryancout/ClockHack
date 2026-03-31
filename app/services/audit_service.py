import json
import os
import tempfile
from datetime import datetime
from pathlib import Path
from app.core.config import DATA_DIR, AUDIT_FILE, MAX_AUDIT

DATA_DIR.mkdir(parents=True, exist_ok=True)


def _load():
    if not AUDIT_FILE.exists():
        return []
    try:
        with open(AUDIT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save(data):
    fd, temp_path = tempfile.mkstemp(prefix=AUDIT_FILE.stem + "_", suffix=".tmp", dir=str(AUDIT_FILE.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, AUDIT_FILE)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def _sanitize_value(value):
    if isinstance(value, dict):
        return {k: _sanitize_value(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_sanitize_value(v) for v in value]
    if isinstance(value, str):
        normalized = value.replace("\\", "/")
        if "/" in normalized or ":" in normalized:
            return Path(normalized).name
    return value


def registrar_evento(acao: str, detalhes: dict):
    data = _load()
    data.insert(0, {
        "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "acao": acao,
        "detalhes": _sanitize_value(detalhes),
    })
    data = data[:MAX_AUDIT]
    _save(data)
