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
POWERBI_SENDS_FILE = DATA_DIR / "powerbi_sends.json"

APP_TITLE = f"{APP_NAME} - V{APP_VERSION}"
APP_GEOMETRY = "1180x760"
MIN_WIDTH = 820
MIN_HEIGHT = 680

EXTENSOES_ENTRADA_SUPORTADAS = frozenset({".csv"})
TIPOS_ARQUIVO_ENTRADA = [("Arquivos CSV", "*.csv")]

COLUNAS_OBRIGATORIAS = [
    "Nome do funcionário",
    "Número de matrícula",
    "Nome do departamento",
    "Banco Total",
    "Banco Saldo",
]

MIN_FUNCIONARIOS_ALERTA = 1
MAX_HISTORICO = 100
MAX_AUDIT = 500
MAX_FILE_SIZE_MB = 25

# Integração Power BI (aplicativo público/desktop, sem segredo embutido).
POWER_BI_CLIENT_ID = "1ca7e64b-41ea-4735-9c79-44370c865324"
POWER_BI_TENANT_ID = "66d8306c-231a-4ac8-ada8-6b3b1b198852"
POWER_BI_WORKSPACE_ID = "9cc540aa-9bed-42eb-a0b9-2ab650317e12"
POWER_BI_WORKSPACE_NAME = "FAS Jornada Analytics"
# O sufixo versiona o contrato de colunas sem apagar modelos/históricos antigos.
POWER_BI_DATASET_NAME = "FAS Jornada Analytics v2"
POWER_BI_TABLE_NAME = "Jornada"

BG_APP = "#2a495b"
BG_CARD = "#ffffff"
BG_BOX = "#f3f6f8"
FG_TITLE = "#17324d"
FG_TEXT = "#31475e"
FG_MUTED = "#52677a"
BORDER = "#d5dfe8"
PRIMARY = "#256fe8"
SUCCESS = "#16805f"
WARNING = "#8a4f00"
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
