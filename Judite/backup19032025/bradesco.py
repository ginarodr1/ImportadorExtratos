def importar_dados_arquivo(self, arquivo):
        wb = openpyxl.load_workbook(arquivo, data_only=True)
        sheet = wb.active

        dados_importados = []

        for row in range(11, sheet.max_row + 1):
            coluna_a = sheet.cell(row=row, column=1).value
            if coluna_a and isinstance(coluna_a, str) and coluna_a.strip().lower() == "total":
                break

            coluna_a = sheet.cell(row=row, column=1).value  #? data
            coluna_b = sheet.cell(row=row, column=2).value  #? descrição
            coluna_c = sheet.cell(row=row, column=3).value  #? n° Doc
            coluna_d = sheet.cell(row=row, column=4).value  #? débito
            coluna_e = sheet.cell(row=row, column=5).value  #? crédito
            coluna_f = sheet.cell(row=row, column=6).value  #? saldo
            

            def tratar_valor(valor):
                if valor is None or valor == "":
                    return 0
                valor = str(valor).replace(".", "").replace(",", ".")
                try:
                    return float(valor)
                except ValueError:
                    return 0

            coluna_d = tratar_valor(coluna_d)
            coluna_e = tratar_valor(coluna_e)

            valor_total = coluna_d + coluna_e
            dados_importados.append([
                coluna_a, coluna_b, coluna_c, valor_total, coluna_f,
                "", "", "", "", "", "", "", "", ""
            ])

        df = pd.DataFrame(dados_importados, columns=[
            "DataLEB", "DescriçãoLEB", "NDocLEB", "ValorLEB", "SaldoLEB",
            "LançamentoLC", "DataLC", "DébitoLC", "D-C/CLC", "CréditoLC",
            "C-C/CLC", "CNPJLC", "HistóricoLC", "ValorLC"
        ])

        for i in self.tree.get_children(): #? limpar a treeview e adicionar os dados
            self.tree.delete(i)

        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))