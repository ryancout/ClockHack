def calcular_totais(ws, col_nome, col_bt, col_bs):
    from app.services.time_service import para_minutos
    soma_bt = 0
    soma_bs = 0
    quantidade_funcionarios = 0
    for row in range(2, ws.max_row + 1):
        nome = ws.cell(row=row, column=col_nome).value
        if nome not in (None, ""):
            quantidade_funcionarios += 1
        soma_bt += para_minutos(ws.cell(row=row, column=col_bt).value)
        soma_bs += para_minutos(ws.cell(row=row, column=col_bs).value)
    return {"soma_bt": soma_bt, "soma_bs": soma_bs, "quantidade_funcionarios": quantidade_funcionarios}
