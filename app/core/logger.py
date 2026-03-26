import logging
from datetime import datetime
from app.core.config import LOG_DIR

LOG_DIR.mkdir(parents=True, exist_ok=True)
log_file = LOG_DIR / f"processador_{datetime.now().strftime('%Y%m%d')}.log"

logger = logging.getLogger("processador_planilhas_fas")
logger.setLevel(logging.INFO)

if not logger.handlers:
    handler = logging.FileHandler(log_file, encoding="utf-8")
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
