"""Configuração explícita do logger da aplicação."""

import logging
from datetime import datetime
from logging.handlers import RotatingFileHandler

from app.core.config import LOG_DIR


logger = logging.getLogger("processador_planilhas_fas")
logger.setLevel(logging.INFO)
logger.propagate = False


def configurar_logger():
    """Configura o arquivo rotativo uma única vez e devolve o logger."""
    if logger.handlers:
        return logger

    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_file = LOG_DIR / f"processador_{datetime.now().strftime('%Y%m%d')}.log"
        handler = RotatingFileHandler(
            log_file,
            encoding="utf-8",
            maxBytes=1_500_000,
            backupCount=5,
        )
        formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    except OSError:
        logger.addHandler(logging.NullHandler())

    return logger
