import json
from app.core.config import DEFAULT_PREFERENCES, DATA_DIR, PREFERENCES_FILE

DATA_DIR.mkdir(parents=True, exist_ok=True)

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
    with open(PREFERENCES_FILE, "w", encoding="utf-8") as f:
        json.dump(preferencias, f, ensure_ascii=False, indent=2)
