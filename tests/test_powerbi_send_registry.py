from datetime import datetime, timezone

from app.services.analytics_service import PowerBiSnapshot
from app.services.powerbi_send_registry import (
    PowerBiSendRegistry,
    calcular_fingerprint_snapshot,
)


def _snapshot(report_id="r1", generated_at="2026-08-18T10:00:00"):
    return PowerBiSnapshot(
        report_id,
        (
            {
                "IDRelatorio": report_id,
                "GeradoEm": generated_at,
                "Arquivo": "relatorio.xlsx",
                "Matricula": "001",
                "BancoSaldoMinutos": -60,
            },
        ),
        "relatorio.xlsx",
    )


def test_fingerprint_ignora_metadados_variaveis_do_mesmo_conteudo():
    primeiro = calcular_fingerprint_snapshot(_snapshot())
    segundo = calcular_fingerprint_snapshot(
        _snapshot(report_id="r2", generated_at="2026-08-18T11:00:00")
    )

    assert primeiro == segundo


def test_fingerprint_muda_quando_dado_analitico_muda():
    original = _snapshot()
    alterado = PowerBiSnapshot(
        "r2",
        ({**original.rows[0], "BancoSaldoMinutos": -120},),
        "outro.xlsx",
    )

    assert calcular_fingerprint_snapshot(original) != calcular_fingerprint_snapshot(
        alterado
    )


def test_registro_persiste_e_localiza_sem_guardar_caminho_completo(tmp_path):
    arquivo = tmp_path / "powerbi_sends.json"
    registry = PowerBiSendRegistry(arquivo)
    momento = datetime(2026, 8, 18, 12, 0, tzinfo=timezone.utc)

    salvo = registry.register(
        fingerprint="abc",
        report_id="relatorio-1",
        dataset_id="dataset-1",
        source_file="C:/Usuarios/Ryan/relatorio.xlsx",
        row_count=47,
        sent_at=momento,
    )

    assert salvo.source_file == "relatorio.xlsx"
    assert PowerBiSendRegistry(arquivo).find("abc") == salvo
    assert "C:/Usuarios/Ryan" not in arquivo.read_text(encoding="utf-8")
