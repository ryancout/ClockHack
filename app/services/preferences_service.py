import json
import os
import tempfile
from app.core.config import DEFAULT_PREFERENCES, DATA_DIR, PREFERENCES_FILE

DATA_DIR.mkdir(parents=True, exist_ok=True)


def _atomic_write_json(path, data):
    fd, temp_path = tempfile.mkstemp(prefix=path.stem + "_", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def carregar_preferencias():
    if not PREFERENCES_FILE.exists():
        salvar_preferencias(DEFAULT_PREFERENCES.copy())
        return DEFAULT_PREFERENCES.copy()
    try:
        with open(PREFERENCES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = DEFAULT_PREFERENCES.copy()
    final = DEFAULT_PREFERENCES.copy()
    final.update(data)
    return final


def salvar_preferencias(preferencias: dict):
    _atomic_write_json(PREFERENCES_FILE, preferencias)
