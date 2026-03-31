from app.services.file_service import garantir_extensao_xlsx, sugerir_nome_saida


def test_sugerir_nome_saida():
    assert sugerir_nome_saida("c:/tmp/arquivo.csv", "DP") == "arquivo_DP_tratado.xlsx"


def test_garantir_extensao_xlsx():
    assert garantir_extensao_xlsx("relatorio_final") == "relatorio_final.xlsx"
    assert garantir_extensao_xlsx("relatorio_final.xlsx") == "relatorio_final.xlsx"
