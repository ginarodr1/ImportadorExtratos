wb = openpyxl.load_workbook(arquivo, data_only=True)
sheet = wb.active

dados_importados = []

for row in range(5, sheet.max_row + 1):
    coluna_h = sheet.cell(row=row, column=1).value
    if coluna_h and isinstance(coluna_h, str) and coluna_h.strip().lower() == "S A L D O":
        break

    coluna_a = sheet.cell(row=row, column=1).value  #? data
    coluna_h = sheet.cell(row=row, column=2).value  #? descrição
    coluna_f = sheet.cell(row=row, column=3).value  #? n° Doc
    coluna_i = sheet.cell(row=row, column=4).value  #? débito
    coluna_e = sheet.cell(row=row, column=5).value  #? crédito
    coluna_f = sheet.cell(row=row, column=6).value  #? saldo