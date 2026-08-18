from datetime import datetime, timezone

from openpyxl import Workbook

from app.services.analytics_service import preparar_snapshot_powerbi


def test_snapshot_powerbi_trata_indicadores_e_nunca_exporta_cpf(tmp_path):
    caminho = tmp_path / "jornada.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet"
    sheet.append(
        [
            "Nome do funcionário",
            "Número de matrícula",
            "Nome do departamento",
            "CPF do funcionário",
            "Total Trabalhado",
            "Horas Previstas",
            "Dia Falta",
            "Extras Total",
            "Banco Total",
            "Banco Saldo",
        ]
    )
    sheet.append(
        ["Ana", "000123", "Operações", "123.456.789-00", "08:30", "08:00", 1, "01:15", "02:00", "-09:00"]
    )
    sheet.append([None, None, None, None, None, None, None, None, "02:00", "-09:00"])
    workbook.save(caminho)

    snapshot = preparar_snapshot_powerbi(
        caminho,
        {
            "tipo_entrada": "RHID",
            "empresa": "FAS",
            "setores": "Operações",
            "periodo_inicial": "2026-08-01",
            "periodo_final": "2026-08-31",
        },
        report_id="relatorio-1",
        generated_at=datetime(2026, 8, 18, 10, 0, tzinfo=timezone.utc),
    )

    assert snapshot.row_count == 1
    linha = snapshot.rows[0]
    assert linha["IDRelatorio"] == "relatorio-1"
    assert linha["Matricula"] == "000123"
    assert linha["TotalTrabalhadoMinutos"] == 510
    assert linha["HorasPrevistasMinutos"] == 480
    assert linha["BancoSaldoMinutos"] == -540
    assert linha["ClassificacaoSaldo"] == "Saldo crítico"
    assert linha["PercentualTrabalhado"] == 106.25
    assert linha["PeriodoInicial"] == "2026-08-01"
    assert all("cpf" not in chave.casefold() for chave in linha)
    assert "123.456.789-00" not in str(linha)


def test_snapshot_manual_aceita_colunas_minimas_e_usa_zero_nas_opcionais(tmp_path):
    caminho = tmp_path / "manual.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.append(
        [
            "Nome do funcionário",
            "Número de matrícula",
            "Nome do departamento",
            "Banco Total",
            "Banco Saldo",
        ]
    )
    sheet.append(["Bia", "77", "RH", "25:00", "03:30"])
    workbook.save(caminho)

    linha = preparar_snapshot_powerbi(caminho, {"tipo_entrada": "CSV"}).rows[0]

    assert linha["TotalTrabalhadoMinutos"] == 0
    assert linha["BancoTotalMinutos"] == 1500
    assert linha["Origem"] == "CSV"
