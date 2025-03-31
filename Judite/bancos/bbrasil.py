def acao_brasil(arquivo):
    print("\n=== INÍCIO DO PROCESSAMENTO ===")
    print(f"Arquivo recebido: {arquivo}")
    try:
                extensao = arquivo.lower().split('.')[-1]
                print(f"Extensão detectada: {extensao}")
                dados_importados = []
                saldo_final_calculado = 0  # Inicializa a variável aqui
                
                if extensao == 'xls':
                    print("\n=== PROCESSANDO ARQUIVO XLS ===")
                    import xlrd
                    
                    print("Abrindo workbook...")
                    wb = xlrd.open_workbook(arquivo)
                    sheet = wb.sheet_by_index(0)
                    print(f"Planilha aberta: {sheet.name}")
                    print(f"Dimensões: {sheet.nrows} linhas x {sheet.ncols} colunas")
                    
                    # Lê o saldo inicial (F10)
                    print("\nBuscando saldo inicial...")
                    saldo_inicial = sheet.cell_value(9, 5)
                    print(f"Valor bruto encontrado em F10: {saldo_inicial}")
                    print(f"Tipo do valor: {type(saldo_inicial)}")
                    
                    # Formata e confirma saldo inicial
                    if isinstance(saldo_inicial, str):
                        print("Convertendo saldo inicial de string para float...")
                        saldo_inicial = float(saldo_inicial.replace(".", "").replace(",", "."))
                    saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                    print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")
                    
                    resposta = messagebox.askyesno("Confirmação de saldo", 
                                                 f"O saldo inicial é de R${saldo_inicial_frmt}?")
                    if not resposta:
                        print("Usuário não confirmou o saldo inicial. Abortando...")
                        return
                        
                    print("Atualizando campo de saldo inicial na interface...")
                    self.saldo_inicial_entry.delete(0, tk.END)
                    self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)
                    
                    # Processa as linhas
                    saldo_final_calculado = saldo_inicial
                    
                    print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                    print(f"Total de linhas na planilha: {sheet.nrows}")
                    
                    for row in range(10, sheet.nrows):
                        try:
                            print(f"\nProcessando linha {row+1}:")
                            data = sheet.cell_value(row, 0)
                            print(f"Data encontrada: {data} (tipo: {type(data)})")
                            
                            if not data:
                                print("Linha vazia, pulando...")
                                continue
                            if isinstance(data, str) and "total" in data.lower():
                                print("Encontrada linha de total, parando processamento...")
                                break
                                
                            historico = sheet.cell_value(row, 1)
                            num_doc = sheet.cell_value(row, 2)
                            credito = sheet.cell_value(row, 3)
                            debito = sheet.cell_value(row, 4)
                            saldo = sheet.cell_value(row, 5)
                            
                            print(f"Valores lidos:")
                            print(f"  Histórico: {historico}")
                            print(f"  Nº Doc: {num_doc}")
                            print(f"  Crédito: {credito}")
                            print(f"  Débito: {debito}")
                            print(f"  Saldo: {saldo}")
                            
                            def tratar_valor(valor):
                                print(f"Tratando valor: {valor} (tipo: {type(valor)})")
                                if valor is None or valor == "":
                                    print("Valor vazio, retornando 0.0")
                                    return 0.0
                                if isinstance(valor, str):
                                    print("Convertendo string para float...")
                                    valor = valor.replace(".", "").replace(",", ".")
                                try:
                                    resultado = float(valor)
                                    print(f"Valor convertido: {resultado}")
                                    return resultado
                                except ValueError as e:
                                    print(f"Erro ao converter valor: {e}")
                                    return 0.0
                            
                            valor_credito = tratar_valor(credito)
                            valor_debito = tratar_valor(debito)
                            valor_total = valor_credito + valor_debito
                            print(f"Valor total calculado: {valor_total}")
                            
                            # Formata a data se for um número
                            if isinstance(data, float):
                                print("Convertendo data de float para string...")
                                data = xlrd.xldate_as_datetime(data, wb.datemode).strftime('%d/%m/%Y')
                                print(f"Data convertida: {data}")
                            
                            print("Adicionando linha aos dados importados...")
                            dados_importados.append([
                                data, historico, num_doc, valor_total, saldo,
                                "", "", "", "", "", "", "", "", ""
                            ])
                            
                            saldo_final_calculado += valor_total
                            print(f"Novo saldo calculado: {saldo_final_calculado}")
                            
                        except Exception as e:
                            print(f"ERRO ao processar linha {row+1}:")
                            print(f"Detalhes do erro: {str(e)}")
                            traceback.print_exc()
                            continue
                    
                else:  # xlsx
                    print("\n=== PROCESSANDO ARQUIVO XLSX ===")
                    wb = openpyxl.load_workbook(arquivo, data_only=True)
                    sheet = wb.active
                    print(f"Planilha ativa: {sheet.title}")
                    
                    # Lê o saldo inicial (F10)
                    print("\nBuscando saldo inicial...")
                    saldo_inicial_celula = sheet['F10'].value
                    print(f"Valor bruto encontrado em F10: {saldo_inicial_celula}")
                    
                    if isinstance(saldo_inicial_celula, str):
                        saldo_inicial = float(saldo_inicial_celula.replace(".", "").replace(",", "."))
                    else:
                        saldo_inicial = float(saldo_inicial_celula)
                        
                    saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                    print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")
                    
                    resposta = messagebox.askyesno("Confirmação de saldo", 
                                                 f"O saldo inicial é de R${saldo_inicial_frmt}?")
                    if not resposta:
                        print("Usuário não confirmou o saldo inicial. Abortando...")
                        return
                        
                    self.saldo_inicial_entry.delete(0, tk.END)
                    self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)
                    
                    # Inicializa o saldo final calculado com o saldo inicial
                    saldo_final_calculado = saldo_inicial
                    
                    print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                    for row in range(11, sheet.max_row + 1):
                        try:
                            print(f"\nProcessando linha {row}:")
                            data = sheet.cell(row=row, column=1).value
                            print(f"Data encontrada: {data}")
                            
                            if not data:
                                print("Linha vazia, pulando...")
                                continue
                            if isinstance(data, str) and "total" in data.lower():
                                print("Encontrada linha de total, parando processamento...")
                                break
                                
                            historico = sheet.cell(row=row, column=2).value
                            num_doc = sheet.cell(row=row, column=3).value
                            credito = sheet.cell(row=row, column=4).value
                            debito = sheet.cell(row=row, column=5).value
                            saldo = sheet.cell(row=row, column=6).value
                            
                            print(f"Valores lidos:")
                            print(f"  Histórico: {historico}")
                            print(f"  Nº Doc: {num_doc}")
                            print(f"  Crédito: {credito}")
                            print(f"  Débito: {debito}")
                            print(f"  Saldo: {saldo}")
                            
                            def tratar_valor(valor):
                                if valor is None or valor == "":
                                    return 0.0
                                if isinstance(valor, str):
                                    valor = valor.replace(".", "").replace(",", ".")
                                try:
                                    return float(valor)
                                except ValueError:
                                    return 0.0
                            
                            valor_credito = tratar_valor(credito)
                            valor_debito = tratar_valor(debito)
                            valor_total = valor_credito + valor_debito
                            print(f"Valor total calculado: {valor_total}")
                            
                            dados_importados.append([
                                data, historico, num_doc, valor_total, saldo,
                                "", "", "", "", "", "", "", "", ""
                            ])
                            
                            saldo_final_calculado += valor_total
                            print(f"Novo saldo calculado: {saldo_final_calculado}")
                            
                        except Exception as e:
                            print(f"ERRO ao processar linha {row}:")
                            print(f"Detalhes do erro: {str(e)}")
                            continue
                
                print("\n=== ATUALIZANDO INTERFACE ===")
                print("Formatando saldo final...")
                saldo_final_calculado_frmt = locale.format_string("%.2f", saldo_final_calculado, grouping=True)
                print(f"Saldo final formatado: R${saldo_final_calculado_frmt}")
                
                print("Atualizando campo de saldo final...")
                self.saldo_final_calculado_entry.delete(0, tk.END)
                self.saldo_final_calculado_entry.insert(0, saldo_final_calculado_frmt)
                
                print("\nLimpando Treeview...")
                for i in self.tree.get_children():
                    self.tree.delete(i)
                    
                print("Inserindo dados na Treeview...")
                print(f"Total de registros a inserir: {len(dados_importados)}")
                for dados in dados_importados:
                    self.tree.insert("", "end", values=dados)
                    
                print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
                print(f"Total de linhas processadas: {len(dados_importados)}")
                
            except Exception as e:
                print("\n=== ERRO FATAL ===")
                print(f"Erro: {str(e)}")
                print("Stack trace:")
                traceback.print_exc()
                messagebox.showerror("Erro", 
                    "Erro ao processar o arquivo. Verifique se:\n\n" +
                    "1. O arquivo está no formato correto\n" +
                    "2. O arquivo não está em modo de exibição protegida\n" +
                    "3. O arquivo está fechado no Excel\n\n" +
                    f"Erro: {str(e)}")
                return
    