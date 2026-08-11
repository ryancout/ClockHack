import importlib


def test_ponto_de_entrada_pode_ser_importado_sem_iniciar_interface():
    modulo = importlib.import_module("main")

    assert callable(modulo.iniciar_app)
    assert callable(modulo.configurar_logger)
