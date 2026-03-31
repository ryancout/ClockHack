import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from app.core.config import LOG_DIR

LOG_DIR.mkdir(parents=True, exist_ok=True)
log_file = LOG_DIR / f"processador_{datetime.now().strftime('%Y%m%d')}.log"

logger = logging.getLogger("processador_planilhas_fas")
logger.setLevel(logging.INFO)
logger.propagate = False

if not logger.handlers:
    handler = RotatingFileHandler(log_file, encoding="utf-8", maxBytes=1_500_000, backupCount=5)
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
