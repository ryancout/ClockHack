import os
import sys

def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)
from app.ui.main_window import iniciar_app

if __name__ == "__main__":
    iniciar_app()
