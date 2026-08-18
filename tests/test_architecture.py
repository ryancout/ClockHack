import importlib
from pathlib import Path


def test_ponto_de_entrada_pode_ser_importado_sem_iniciar_interface():
    modulo = importlib.import_module("main")

    assert callable(modulo.iniciar_app)
    assert callable(modulo.configurar_logger)


def test_recursos_e_funcoes_de_tempo_tem_uma_unica_fonte():
    raiz = Path(__file__).resolve().parents[1]

    assert (raiz / "app" / "assets" / "icon.ico").is_file()
    assert not (raiz / "icone.ico").exists()
    assert not (raiz / "app" / "services" / "time_service.py").exists()
