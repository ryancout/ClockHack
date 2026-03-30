def listar_departamentos(ws, col_departamento):
    departamentos = set()
    for row in range(2, ws.max_row + 1):
        valor = ws.cell(row=row, column=col_departamento).value
        if valor not in (None, ""):
            departamentos.add(str(valor).strip())
    return ["Todos"] + sorted(departamentos)

def aplicar_filtro_departamento(ws, col_departamento, departamento):
    if not departamento or departamento == "Todos":
        return
    linhas_manter = [1]
    for row in range(2, ws.max_row + 1):
        valor = ws.cell(row=row, column=col_departamento).value
        if str(valor).strip() == departamento:
            linhas_manter.append(row)
    for row in range(ws.max_row, 1, -1):
        if row not in linhas_manter:
            ws.delete_rows(row, 1)
