import os
from pathlib import Path
from app.core.version import APP_NAME, APP_VERSION

BASE_DIR = Path(__file__).resolve().parents[2]
APP_DIR = BASE_DIR / "app"
ASSETS_DIR = APP_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "logo.png"


def _user_data_base_dir() -> Path:
    local_appdata = os.environ.get("LOCALAPPDATA")
    appdata = os.environ.get("APPDATA")
    home = Path.home()

    if local_appdata:
        return Path(local_appdata) / "ProcessadorPlanilhasFAS"
    if appdata:
        return Path(appdata) / "ProcessadorPlanilhasFAS"
    return home / ".processador_planilhas_fas"


RUNTIME_BASE_DIR = _user_data_base_dir()
DATA_DIR = RUNTIME_BASE_DIR / "data"
LOG_DIR = RUNTIME_BASE_DIR / "logs"
UPLOAD_DIR = RUNTIME_BASE_DIR / "uploads"

HISTORY_FILE = DATA_DIR / "history.json"
PREFERENCES_FILE = DATA_DIR / "preferences.json"
AUDIT_FILE = DATA_DIR / "audit.json"

APP_TITLE = f"{APP_NAME} - V{APP_VERSION}"
APP_GEOMETRY = "1220x760"
MIN_WIDTH = 1220
MIN_HEIGHT = 760

EXTENSOES_ACEITAS = [
    ("Arquivos Excel e CSV", "*.xlsx *.csv"),
    ("Arquivos Excel", "*.xlsx"),
    ("Arquivos CSV", "*.csv"),
]

COLUNAS_OBRIGATORIAS = [
    "Nome do funcionário",
    "Nome do departamento",
    "Banco Total",
    "Banco Saldo",
]

MIN_FUNCIONARIOS_ALERTA = 1
MAX_HISTORICO = 100
MAX_AUDIT = 500
MAX_FILE_SIZE_MB = 25

BG_APP = "#eef3f8"
BG_CARD = "#ffffff"
BG_BOX = "#f7fafc"
FG_TITLE = "#17324d"
FG_TEXT = "#31475e"
FG_MUTED = "#6b7f92"
BORDER = "#d5dfe8"
PRIMARY = "#0b63ce"
SUCCESS = "#0f7a5f"
WARNING = "#b76e00"
ERROR = "#b42318"

FONT_TITLE = ("Segoe UI", 24, "bold")
FONT_SUBTITLE = ("Segoe UI", 11)
FONT_BUTTON = ("Segoe UI", 11, "bold")
FONT_STATUS = ("Segoe UI", 10)
FONT_METRIC_TITLE = ("Segoe UI", 10, "bold")
FONT_METRIC_VALUE = ("Segoe UI", 20, "bold")

DEFAULT_PREFERENCES = {
    "last_open_dir": "",
    "last_save_dir": "",
    "last_department": "Todos",
}
