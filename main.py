"""Ponto de entrada do aplicativo desktop."""

from app.core.logger import configurar_logger
from app.ui.main_window import iniciar_app


if __name__ == "__main__":
    configurar_logger()
    iniciar_app()
