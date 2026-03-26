from app.services.time_service import para_minutos, formatar_horas

def test_para_minutos():
    assert para_minutos("01:30") == 90

def test_formatar_horas():
    assert formatar_horas(1783) == "29:43"
