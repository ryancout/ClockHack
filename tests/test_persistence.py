import json

from app.services import audit_service, history_service, preferences_service


def test_preferencias_criam_diretorio_somente_ao_salvar(tmp_path, monkeypatch):
    caminho = tmp_path / "dados" / "preferences.json"
    monkeypatch.setattr(preferences_service, "PREFERENCES_FILE", caminho)

    preferences_service.salvar_preferencias({"last_department": "Teste"})

    assert json.loads(caminho.read_text(encoding="utf-8")) == {
        "last_department": "Teste"
    }


def test_historico_cria_diretorio_e_remove_caminhos(tmp_path, monkeypatch):
    caminho = tmp_path / "dados" / "history.json"
    monkeypatch.setattr(history_service, "HISTORY_FILE", caminho)

    history_service.registrar_historico(
        {
            "arquivo_origem": "C:/privado/entrada.csv",
            "arquivo_saida": "C:/privado/saida.xlsx",
        }
    )

    item = json.loads(caminho.read_text(encoding="utf-8"))[0]
    assert item["arquivo_origem"] == "entrada.csv"
    assert item["arquivo_saida"] == "saida.xlsx"


def test_auditoria_cria_diretorio_e_sanitiza_caminho(tmp_path, monkeypatch):
    caminho = tmp_path / "dados" / "audit.json"
    monkeypatch.setattr(audit_service, "AUDIT_FILE", caminho)

    audit_service.registrar_evento(
        "teste",
        {"arquivo": "C:/privado/entrada.csv"},
    )

    item = json.loads(caminho.read_text(encoding="utf-8"))[0]
    assert item["acao"] == "teste"
    assert item["detalhes"]["arquivo"] == "entrada.csv"
