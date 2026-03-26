from app.services.file_service import sugerir_nome_saida

def test_sugerir_nome_saida():
    assert sugerir_nome_saida("c:/tmp/arquivo.csv", "DP") == "arquivo_DP_tratado.xlsx"
