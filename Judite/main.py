import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import openpyxl
from openpyxl import Workbook
import pandas as pd
import csv
import os
import re
import locale
from datetime import datetime, date
import traceback
import xlrd
from unidecode import unidecode

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def open_main_window():
    root = tk.Tk()
    app = ImportadorExtratos(root)
    root.mainloop()

class ImportadorExtratos:
    def __init__(self, root):
        self.root = root
        self.root.title("Importador de Extratos Bancários")
        self.root.geometry("1000x600")
        self.root.config(bg='#D9D9D9')
        self.root.state("zoomed")
        self.root.iconbitmap(r"C:\Users\regina.santos\Desktop\Automacao\Judite\icon.ico")
        self.root.protocol("WM_DELETE_WINDOW", self.fechar_janela)
        csv_file_path = r"C:\Users\regina.santos\Desktop\Automacao\Judite\lancamentoscontas1.csv"
        self.df_banco_dados = pd.read_csv(csv_file_path, delimiter=';')

        for i in range(6):
            root.grid_columnconfigure(i, weight=0)
        for i in range(6):
            root.grid_rowconfigure(i, weight=0)

        self.janelas_filhas = [] #? armazenar todas as janelas filhas

        #! -------------------- TÍTULO PRINCIPAL -------------------- #

        self.title_label = tk.Label(root, text="Importador de Extratos", font=("Roboto", 17, "bold"), bg='#D9D9D9')
        self.title_label.grid(row=3, column=0, padx=10, sticky="w")

        #! -------------------- BOTÕES PRINCIPAIS -------------------- #

        self.btn_importar = tk.Button(root, text="Importar", command=self.mostrar_selecao_conta, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_importar.grid(row=4, column=0, padx=10, sticky="w")

        self.btn_limpar = tk.Button(root, text="Limpar Tudo", command=self.confirmar_limpar_dados, font=("Roboto", 10), bg="#FF5733", fg="white", relief="groove", padx=5, pady=3)
        self.btn_limpar.grid(row=4, column=0, padx=80, sticky="w")

        self.btn_classificar1p = tk.Button(root, text="Classificar 1P", command=self.classificar_dados, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_classificar1p.grid(row=2, column=3, padx=90, sticky="e")

        self.btn_limpar1p = tk.Button(root, text="Limpar 1P", command=self.limpar_1p, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_limpar1p.grid(row=2, column=3, padx=5, sticky="e")

        self.btn_exportar = tk.Button(root, text="Exportar", command=self.exportar_dados, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_exportar.grid(row=2, column=4, sticky="w")

        #! -------------------- CAMPOS DE SALDO -------------------- #
        #* -------------- LANÇAMENTO EXTRATO BANCÁRIO -------------- #

        self.saldo_inicial_label = tk.Label(root, text="Saldo Inicial Importado>", font=("Roboto", 11), bg='#D9D9D9', anchor='e')
        self.saldo_inicial_label.grid(row=2, column=1, padx=1, sticky="w")
        self.saldo_inicial_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_inicial_entry.grid(row=2, column=1, columnspan=2, padx=160, sticky="w")

        self.saldo_final_label = tk.Label(root, text="Saldo Final Importado>", font=("Roboto", 11), bg='#D9D9D9', anchor='e')
        self.saldo_final_label.grid(row=3, column=1, padx=3, sticky="w")
        self.saldo_final_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_final_entry.grid(row=3, column=1, columnspan=2, padx=160, sticky="w")

        self.saldo_final_calculado_label = tk.Label(root, text="Saldo Final Calculado>", font=("Roboto", 11), bg='#D9D9D9', anchor='e')
        self.saldo_final_calculado_label.grid(row=4, column=1, padx=3, sticky="w")
        self.saldo_final_calculado_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_final_calculado_entry.grid(row=4, column=1, columnspan=2, padx=160, sticky="w")

        self.diferenca_label = tk.Label(root, text="Diferença>", font=("Roboto", 11), bg='#D9D9D9', anchor='e')
        self.diferenca_label.grid(row=5, column=1, padx=82, sticky="w")
        self.diferenca_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.diferenca_entry.grid(row=5, column=1, columnspan=2, padx=160, sticky="w")

        self.empresa_label = tk.Label(root, text="Empresa>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.empresa_label.grid(row=3, column=2, padx=51, sticky="w")
        self.empresa_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.empresa_entry.grid(row=3, column=2, padx=125, sticky="w")

        self.conta_label = tk.Label(root, text="Conta>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.conta_label.grid(row=4, column=2, padx=72, sticky="w")
        self.conta_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.conta_entry.grid(row=4, column=2, padx=125, sticky="w")

        self.centro_custo_label = tk.Label(root, text="C/Custo>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.centro_custo_label.grid(row=5, column=2, padx=54, sticky="w")
        self.centro_custo_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.centro_custo_entry.grid(row=5, column=2, padx=125, sticky="w")

        self.banco_label = tk.Label(root, text="Banco>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.banco_label.grid(row=4, column=2, padx=1, sticky="e")
        self.banco_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=20)
        self.banco_entry.grid(row=4, column=3, padx=1, sticky="w")
        
        self.agencia_conta_label = tk.Label(root, text="Agência/Conta>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.agencia_conta_label.grid(row=5, column=2, padx=1, sticky="e")
        self.agencia_conta_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=20)
        self.agencia_conta_entry.grid(row=5, column=3, padx=1, sticky="w")

        self.saldo_final_contabil_label = tk.Label(root, text="Saldo Final Contábil>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.saldo_final_contabil_label.grid(row=4, column=3, columnspan=4, padx=95, sticky="e")
        self.saldo_final_contabil_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.saldo_final_contabil_entry.grid(row=4, column=3, columnspan=4, padx=10, sticky="e")

        self.diferenca_extrato_bancario_label = tk.Label(root, text="Diferença c/Extrato Bancário>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.diferenca_extrato_bancario_label.grid(row=5, column=3, columnspan=4, padx=95, sticky="e")
        self.diferenca_extrato_bancario_entry = tk.Entry(root, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.diferenca_extrato_bancario_entry.grid(row=5, column=3, columnspan=4, padx=10, sticky="e")

        self.linhadeajuda_label = tk.Label(root, text="Coluna 0", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        self.linhadeajuda_label.grid(row=1, column=0)
        self.linhadeajuda_label = tk.Label(root, text="Coluna 1", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        self.linhadeajuda_label.grid(row=1, column=1)
        self.linhadeajuda_label = tk.Label(root, text="Coluna 2", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        self.linhadeajuda_label.grid(row=1, column=2)
        self.linhadeajuda_label = tk.Label(root, text="Coluna 3", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        self.linhadeajuda_label.grid(row=1, column=3)
        self.linhadeajuda_label = tk.Label(root, text="Coluna 4", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        self.linhadeajuda_label.grid(row=1, column=4)
        self.linhadeajuda_label = tk.Label(root, text="Coluna 5", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        self.linhadeajuda_label.grid(row=1, column=5)

        #* -------------------- TREEVIEW PARA EXIBIR OS DADOS IMPORTADOS -------------------- #
        self.frame_tree = tk.Frame(root)
        self.frame_tree.grid(row=7, column=0, columnspan=6, pady=5, padx=10, sticky="nsew")

        self.tree = ttk.Treeview(self.frame_tree, columns=( #? cria a Treeview dentro do frame
            "DataLEB", "DescricaoLEB", "NDocLEB", "ValorLEB", "SaldoLEB",
            "LancamentoLC", "DataLC", "DebitoLC", "D-C/CLC", "CreditoLC",
            "C-C/CLC", "CNPJLC", "HistoricoLC", "ValorLC"
        ), show="headings")

        self.scrollbar_y = ttk.Scrollbar(self.frame_tree, orient="vertical", command=self.tree.yview) #? cria a scrollbar vertical
        self.scrollbar_y.pack(side="right", fill="y")

        self.tree.configure(yscrollcommand=self.scrollbar_y.set) #? configura a treeview para usar a scrollbar

        self.tree.pack(side="left", fill="both", expand=True) #? empacota a treeview
        
        self.tree.tag_configure('negativo', foreground='red') #? torna os valores negativos em vermelho

        self.tree.tag_configure('linha_par', background='#F0F0F0') #? cinza
        self.tree.tag_configure('linha_impar', background='#ffffff') #? branco

        #* -------------------- DEFINIR CABEÇALHOS DAS COLUNAS ------------------- #
        self.tree.heading("DataLEB", text="Data")
        self.tree.heading("DescricaoLEB", text="Descrição")
        self.tree.heading("NDocLEB", text="N° Doc")
        self.tree.heading("ValorLEB", text="Valor")
        self.tree.heading("SaldoLEB", text="Saldo")
        self.tree.heading("LancamentoLC", text="Lançamento")
        self.tree.heading("DataLC", text="Data")
        self.tree.heading("DebitoLC", text="Débito")
        self.tree.heading("D-C/CLC", text="D-C/C")
        self.tree.heading("CreditoLC", text="Crédito")
        self.tree.heading("C-C/CLC", text="C-C/C")
        self.tree.heading("CNPJLC", text="CNPJ")
        self.tree.heading("HistoricoLC", text="Histórico")
        self.tree.heading("ValorLC", text="Valor")

        #* -------------------- AJUSTAR O TAMANHO DAS COLUNAS -------------------- #
        self.tree.column("DataLEB", width=60)
        self.tree.column("DescricaoLEB", width=150)
        self.tree.column("NDocLEB", width=60)
        self.tree.column("ValorLEB", width=60)
        self.tree.column("SaldoLEB", width=60)
        self.tree.column("LancamentoLC", width=150)
        self.tree.column("DataLC", width=60)
        self.tree.column("DebitoLC", width=60)
        self.tree.column("D-C/CLC", width=50)
        self.tree.column("CreditoLC", width=60)
        self.tree.column("C-C/CLC", width=50)
        self.tree.column("CNPJLC", width=80)
        self.tree.column("HistoricoLC", width=100)
        self.tree.column("ValorLC", width=80)

        #* -------------------- TORNAR A TREEVIEW EXPANSÍVEL -------------------- #
        root.grid_rowconfigure(7, weight=1)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=1)
        root.grid_columnconfigure(2, weight=1)
        root.grid_columnconfigure(3, weight=1)
        root.grid_columnconfigure(4, weight=1)
        root.grid_columnconfigure(5, weight=1)

    def mostrar_selecao_conta(self): #? cria janela toplevel
        janela_selecao_conta = tk.Toplevel(self.root)
        app = TelaSelecaoConta(janela_selecao_conta, self.abrir_explorador_arquivos)
        self.janelas_filhas.append(janela_selecao_conta) #? adiciona janela filha à lista

    def abrir_tela_nova_empresa(self):
        if not hasattr(self, 'tela_nova_empresa') or not self.tela_nova_empresa.winfo_exists():
            self.tela_nova_empresa = tk.Toplevel(self.root)
            app = TelaNovaEmpresa(self.tela_nova_empresa, self.salvar_nova_empresa)
            self.janelas_filhas.append(self.tela_nova_empresa) #? adiciona janela filha à lista

    def abrir_tela_nova_conta(self):
        if not hasattr(self, 'tela_nova_conta') or not self.tela_nova_conta.winfo_exists():
            self.tela_nova_conta = tk.Toplevel(self.root)
            app = TelaNovaConta(self.tela_nova_conta, self.salvar_nova_conta)
            self.janelas_filhas.append(self.tela_nova_conta) #? adiciona janela filha à lista

    def fechar_janela(self):
        for janela in self.janelas_filhas: #? fecha todas as janelas filhas
            if janela.winfo_exists():
                janela.destroy()
        self.root.destroy() #? fecha a janela principal

    def abrir_explorador_arquivos(self, empresa, conta, arquivo, respostas_safra=None):
        self.empresa_entry.delete(0, tk.END)
        self.empresa_entry.insert(0, empresa.split(" - ")[0])
        self.conta_entry.delete(0, tk.END)
        self.conta_entry.insert(0, conta.split(" - ")[0])
        self.banco_entry.delete(0, tk.END)
        self.banco_entry.insert(0, conta.split(" - ")[1])

        agencia = conta.split(" - ")[3]
        conta_bancaria = conta.split(" - ")[4]
        self.agencia_conta_entry.delete(0, tk.END)
        self.agencia_conta_entry.insert(0, f"{agencia}/{conta_bancaria}")

        self.detectar_banco(arquivo, respostas_safra)

    def detectar_banco(self, nome_arquivo, respostas_safra=None):
        bancos = [
            "Bco.Brasil", "Banco do Brasil", "BANCO DO BRASIL", "BB", "BRASIL", "Brasil",
            "Bco.Inter", "Banco Inter", "BANCO INTER", "Inter", "INTER",
            "Bco.Caixa", "Banco Caixa Eletrônica", "BANCO CAIXA ELETRÔNICA", "Caixa Eletrônica", "Caixa", 
            "Bco.Bradesco", "Banco Bradesco", "BANCO BRADESCO", "Bradesco", "EXTRATO BRADESCO",
            "Bco.Grafeno", "Banco Grafeno", "BANCO GRAFENO", "GRAFENO", "Grafeno",
            "Bco.Pagseguro", "Banco Pagseguro", "BANCO PAGSEGURO", "PAGSEGURO", "Pagseguro",
            "Bco.C6Bank", "Banco C6 Bank", "C6 Bank", "BANCO C6 BANK", "C6 BANK", "C6BANK", "C6",
            "Bco.Itaú", "Banco Itaú", "BANCO ITAÚ", "ITAÚ", "Itaú",
            "Bco.Santander", "Banco Santander", "BANCO SANTANDER", "SANTANDER", "Santander",
            "Bco.HSBC", "Banco HSBC", "BANCO HSBC", "HSBC",
            "Bco.Safra", "Banco Safra", "BANCO SAFRA", "SAFRA", "Safra",
            "Bco.Suisse", "Banco Suisse", "Banco Credit Suisse", "Credit Suisse", "CREDIT SUISSE", "SUISSE", "BANCO SUISSE",
            "Bco.Daycoval", "Banco Daycoval", "BANCO DAYCOVAL", "DAYCOVAL", "Daycoval",
            "Bco.Itaú", "Banco Itaú", "BANCO ITAÚ", "ITAÚ", "Itaú", "Itau", "ITAU",
        ]
    
        for banco in bancos:
            if banco in nome_arquivo:
                print(f"Banco detectado: {banco}")
                self.executar_acao_para_banco(banco, nome_arquivo, respostas_safra)
                break
        else:
            print("Nenhum banco detectado no nome do arquivo.")

    def copiar_coluna(self, coluna_origem, coluna_destino):
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')

            valor_origem = values[self.tree["columns"].index(coluna_origem)]
            self.tree.set(item, coluna_destino, valor_origem)

        #* -------------------- ETAPA DA CLISSIFICAÇÃO -------------------- #
    def classificar_dados(self):
        conta_bancaria = self.conta_entry.get()
        incluir_banco = messagebox.askyesno("Classificar Dados", "Deseja incluir a identificação do banco no histórico contábil?")
        
        conta_ativo = None #? carregar contas.csv para obter a conta ativo
        try:
            if os.path.exists('contas.csv'):
                df_contas = pd.read_csv('contas.csv')
                filtro = (df_contas['Empresa'] == self.empresa_entry.get()) & (df_contas['Banco'] == self.banco_entry.get())
                conta_encontrada = df_contas[filtro]
                if not conta_encontrada.empty:
                    conta_ativo = str(conta_encontrada.iloc[0]['Conta Ativo']).strip()
                    print(f"Conta ativo encontrada: {conta_ativo}")
        except Exception as e:
            print(f"Erro ao carregar contas.csv: {e}")

        padroes_lancamentos = { #? dicionário de padrões de lançamentos
            #! BRADESCO
            "TRANSF CC PARA CC PJ": "TRANSF CC PARA CC PJ", 
            "TED-TRANSF ELET DISPON REMET.": "TED-TRANSF ELET DISPON REMET.",
            "TRANSFERENCIA PIX REM:": "TRANSFERENCIA PIX REM:",
            "TRANSFERENCIA PIX DES:": "TRANSFERENCIA PIX DES:",
            "TRANSFERENCIA PIX REM: SHOCK METAIS NAO FERRAGENS": "TRANSFERENCIA PIX REM: SHOCK METAIS NAO FERRAGENS",
            "ENCARGOS C GARANTIA ENCARGO": "ENCARGOS C GARANTIA ENCARGO",
            "ENCARGOS C GARANTIDA ENCARGO CONTR": "ENCARGOS C GARANTIDA ENCARGO", 
            "ENCARGOS C GARANTIDA IOF     CONTR": "ENCARGOS C GARANTIDA IOF",
            "TARIFA REGISTRO COBRANCA": "TARIFA REGISTRO COBRANCA",
            "PAGTO ELETRON  COBRANCA": "PAGTO ELETRON  COBRANCA",
            "VISA CREDITO": "VISA CREDITO",
            "RECEBIMENTO FORNECEDOR": "RECEBIMENTO FORNECEDOR",
            "PAGTO ELETRONICO TRIBUTO NET EMPR LIC ELET": "PAGTO ELETRONICO TRIBUTO NET EMPR LIC ELET",
            "OPERACAO CAPITAL GIRO": "OPERACAO CAPITAL GIRO",
            "PARCELA OPER CREDITO CONTR": "PARCELA OPER CREDITO CONTR",
            "PAGTO ELETRONICO TRIBUTO INTERNET": "PAGTO ELETRONICO TRIBUTO INTERNET",
            "TARIFA AUTORIZ COBRANCA TR TIT PAGO CARTORIO": "TAR",
            "TED D CC HBANK* DEST. SHOCK METAIS NAO FER": "TED D CC HBANK* DEST. SHOCK METAIS NAO FER",
            "TAR PACOTE MENSAL BASICO": "Tarifa Bancária",
            "RECEBIMENTO TED D REMET.SHOCK METAIS N FERRO": "RECEBIMENTO TED D REMET.SHOCK METAIS N FERRO",
            "TRANSF.EXCEDENTEGARANTIA": "TRANSF.EXCEDENTEGARANTIA",
            "PARCELA OPER CREDITO CONTR": "PARCELA OPER CREDITO CONTR",
            "PIX QR CODE DINAMICO DES:": "PIX QR CODE DINAMICO",
            "TARIFA BANCARIA TRANSF PGTO PIX": "TARIFA BANCARIA",
            "PAGTO ELETRON COBRANCA CONTR": "PAGTO ELETRON COBRANCA",
            "PAGTO ELETRON COBRANCA": "PAGTO ELETRON COBRANCA",
            "TAR COMANDADA COBRANCA": "Tarifa Bancária",
        }

        
        for item in self.tree.get_children(): #? processar cada item na treeview
            values = list(self.tree.item(item, 'values'))
            descricao = values[self.tree["columns"].index("DescricaoLEB")] #? copiar e simplificar descrições
            
            lancamento_simplificado = descricao #? procurar por padrões conhecidos

            if descricao.startswith("TARIFA AUTORIZ COBRANCA TR TIT PAGO CARTORIO"):
                lancamento_simplificado = "TAR"
            else:
                for padrao in padroes_lancamentos:
                    if descricao.startswith(padrao):
                        lancamento_simplificado = padroes_lancamentos[padrao]
                        break
            
            lancamento_idx = self.tree["columns"].index("LancamentoLC") #? atualizar o lançamento
            values[lancamento_idx] = lancamento_simplificado
            
            if incluir_banco: #? adicionar identificação do banco se solicitado
                banco_nome = self.banco_entry.get()
                agencia_conta = self.agencia_conta_entry.get()
                agencia_nome = agencia_conta.split('/')[0]
                conta_nome = agencia_conta.split('/')[1]
                
                if values[lancamento_idx]:  #? verifica se o lançamento não está vazio
                    values[lancamento_idx] = f"{values[lancamento_idx]} - Bco.{banco_nome} Ag.{agencia_nome} CC.{conta_nome}"
            
            self.tree.item(item, values=values)

        df_descricoes_normalizadas = self.df_banco_dados['Descricao'].apply(
            lambda x: unidecode(str(x).strip().upper())
        )

        #? verifica os itens na treeview para classificar débito e crédito
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            descricao_idx = self.tree["columns"].index("LancamentoLC")
            descricao = values[descricao_idx]  #? a descricao está na segunda coluna

            descricao_normalizada = unidecode(str(descricao).strip().upper())

            correspondencia = self.df_banco_dados[
                df_descricoes_normalizadas == descricao_normalizada
            ]

            if not correspondencia.empty: #? obter os valores correspondentes usando índices
                descricao_banco = correspondencia.iloc[0, 5]  #? coluna F (descricao) no banco
                debito = correspondencia.iloc[0, 6]  #? coluna G (debito) no banco
                credito = correspondencia.iloc[0, 9]  #? coluna J (credito) no banco

                if "#BCO" in str(debito):
                    debito = conta_bancaria
                if "#BCO" in str(credito):
                    credito = conta_bancaria

                #? atualizar apenas as colunas relevantes
                novos_valores = list(values)  #? copiar os valores existentes
                novos_valores[7] = debito  #? atualizar a coluna de debito
                novos_valores[9] = credito  #? atualizar a coluna de credito
                self.tree.item(item, values=novos_valores)

        self.copiar_coluna("DataLEB", "DataLC")
        self.copiar_coluna("ValorLEB", "ValorLC")
        self.copiar_coluna("LancamentoLC", "HistoricoLC")

        messagebox.showinfo("Classificar Dados", "Dados classificados com sucesso!")

    def executar_acao_para_banco(self, nome_banco, arquivo, respostas_safra=None):
        print(f"Banco detectado: {nome_banco}")
        acoes = {
            "Bco.Bradesco": self.acao_bradesco, "Banco Bradesco": self.acao_bradesco, "BANCO BRADESCO": self.acao_bradesco, "Bradesco": self.acao_bradesco, "EXTRATO BRADESCO": self.acao_bradesco,
            "Bco.Safra": self.acao_safra, "Banco Safra": self.acao_safra, "BANCO SAFRA": self.acao_safra, "SAFRA": self.acao_safra, "Safra": self.acao_safra, "EXTRATO SAFRA": self.acao_safra,
            "Bco.Daycoval": self.acao_daycoval, "Banco Daycoval": self.acao_daycoval, "BANCO DAYCOVAL": self.acao_daycoval, "DAYCOVAL": self.acao_daycoval, "Daycoval": self.acao_daycoval, "EXTRATO DAYCOVAL": self.acao_daycoval,
            "Bco.Itaú": self.acao_itau, "Banco Itaú": self.acao_itau, "BANCO ITAÚ": self.acao_itau, "ITAÚ": self.acao_itau, "Itaú": self.acao_itau, "EXTRATO ITAU": self.acao_itau, "EXTRATO ITAÚ": self.acao_itau,
            "Bco.Brasil": self.acao_brasil, "Banco do Brasil": self.acao_brasil, "BANCO DO BRASIL": self.acao_brasil, "BB": self.acao_brasil, "BRASIL": self.acao_brasil, "Brasil": self.acao_brasil, "EXTRATO BB": self.acao_brasil,
            "Bco.Santander": self.acao_santander, "Banco Santander": self.acao_santander, "BANCO SANTANDER": self.acao_santander, "SANTANDER": self.acao_santander, "Santander": self.acao_santander, "EXTRATO SANTANDER": self.acao_santander,
        }

        print(f"Chaves disponíveis: {list(acoes.keys())}")

        if nome_banco in acoes:
            print(f"Executando ação específica para {nome_banco}.")
            if nome_banco in ["Bco.Safra", "Banco Safra", "BANCO SAFRA", "SAFRA", "Safra", "EXTRATO SAFRA"]:
                acoes[nome_banco](arquivo, respostas_safra)
            else:
                acoes[nome_banco](arquivo)
        else:
            print(f"Executando ação padrão para {nome_banco}.")

    def acao_bradesco(self, arquivo):
            print("\n=== INÍCIO DO PROCESSAMENTO: BANCO BRADESCO ===")
            print(f"Arquivo recebido: {arquivo}")
            try:
                extensao = arquivo.lower().split('.')[-1]
                print(f"Extensão detectada: {extensao}")
                dados_importados = []
                saldo_final_calculado = 0  #? inicializa a variável aqui
                
                if extensao == 'xls': #! SE O ARQUIVO É XLS
                    print("\n=== PROCESSANDO ARQUIVO XLS ===")
                    wb = xlrd.open_workbook(arquivo)
                    sheet = wb.sheet_by_index(0)
                    print(f"Planilha aberta: {sheet.name}")
                    print(f"Dimensões: {sheet.nrows} linhas x {sheet.ncols} colunas")
                    
                    print("\nBuscando saldo inicial...") #? lê o saldo inicial (F10)
                    saldo_inicial = sheet.cell_value(9, 5)
                    print(f"Valor bruto encontrado em F10: {saldo_inicial}")
                    print(f"Tipo do valor: {type(saldo_inicial)}")
                    
                    if isinstance(saldo_inicial, str): #? formata e confirma saldo inicial
                        print("Convertendo saldo inicial de string para float...")
                        saldo_inicial = float(saldo_inicial.replace(".", "").replace(",", "."))
                    saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                    print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")
                    
                    resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                    if not resposta:
                        print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                        saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                        if saldo_inicial_manual:
                            try:
                                saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                            except ValueError:
                                messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                                return
                        else:
                            messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                            return
                        
                    print("Atualizando campo de saldo inicial na interface...")
                    self.saldo_inicial_entry.delete(0, tk.END)
                    self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)
                    
                    saldo_final_calculado = saldo_inicial #? processa as linhas
                    
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
                            
                            def converter_para_float(valor):
                                if valor is None or valor == "":
                                    return 0.0
                                if isinstance(valor, str):
                                    valor = valor.replace(".", "").replace(",", ".")
                                try:
                                    return float(valor)
                                except ValueError:
                                    return 0.0
                            
                            valor_credito = converter_para_float(credito)
                            valor_debito = converter_para_float(debito)
                            valor_total = valor_credito + valor_debito
                            print(f"Valor total calculado: {valor_total}")
                            
                            if isinstance(data, float): #? formata a data se for um número
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
                    
                else:  #! SE O ARQUIVO É XLSX
                    print("\n=== PROCESSANDO ARQUIVO XLSX ===")
                    wb = openpyxl.load_workbook(arquivo, data_only=True)
                    sheet = wb.active
                    print(f"Planilha ativa: {sheet.title}")
                    
                    print("\nBuscando saldo inicial...") #? lê o saldo inicial (F10)
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
                    
                    #? inicializa o saldo final calculado com o saldo inicial
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
                            
                            def converter_para_float(valor):
                                if valor is None or valor == "":
                                    return 0.0
                                if isinstance(valor, str):
                                    valor = valor.replace(".", "").replace(",", ".")
                                try:
                                    return float(valor)
                                except ValueError:
                                    return 0.0
                            
                            valor_credito = converter_para_float(credito)
                            valor_debito = converter_para_float(debito)
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

                for i, dados in enumerate(dados_importados):
                    tag = 'linha_par' if i % 2 == 0 else 'linha_impar'
                    self.tree.insert("", "end", values=dados, tags=(tag,))
                    
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

    def acao_safra(self, arquivo, respostas_safra):
        print("\n=== INÍCIO DO PROCESSAMENTO: SAFRA ===")
        print(f"Arquivo recebido: {arquivo}")
        def formatar_valor_brasileiro(valor):
            try:
                return locale.format_string("%.2f", float(valor), grouping=True)
            except:
                return valor

        if respostas_safra:
            novo_formato = respostas_safra.get("novo_formato", False)
            conta_vinculada = respostas_safra.get("conta_vinculada", False)
            extensao = arquivo.lower().split('.')[-1]

            if extensao in ["xls", "xlsx"]:
                if novo_formato:
                    if conta_vinculada: #! LÓGICA PARA ARQUIVO XLS/XLSX, NOVO FORMATO E CONTA VINCULADA
                        print("Processando XLS/XLSX, novo formato no Safra, conta vinculada.")
                        try:
                            dados_importados = []
                            saldo_final_calculado = 0

                            if extensao == "xls": #! SE O ARQUIVO É XLS
                                import xlrd
                                print("\n=== PROCESSANDO ARQUIVO XLS ===")
                                print("Abrindo workbook...")
                                wb = xlrd.open_workbook(arquivo)
                                sheet = wb.sheet_by_index(0)

                                texto_periodo = sheet.cell_value(5, 1)
                                print(f"Texto do período: {texto_periodo}")

                                def extrair_ano(texto):
                                    match = re.search(r'\d{2}/\d{2}/(\d{4})', texto)
                                    if match:
                                        return int(match.group(1))
                                    return datetime.now().year
                                
                                ano_extrato = extrair_ano(texto_periodo)
                                print(f"Ano extraído do período: {ano_extrato}")

                                meses_pt = {
                                    'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
                                    'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
                                }

                                def converter_data_pt(data_str, ano):
                                    try:
                                        dia, mes_abrev = data_str.strip().split('/')
                                        mes = meses_pt.get(mes_abrev.lower())
                                        if mes:
                                            data = datetime(ano, mes, int(dia))
                                            return data.strftime('%d/%m/%Y')
                                    except Exception as e:
                                        print(f"Erro ao converter data '{data_str}': {e}")
                                    return data_str

                                print(f"Planilha aberta: {sheet.name}")
                                print(f"Dimensões? {sheet.nrows} linhas x {sheet.ncols} colunas")

                                print("\nBuscando saldo inicial...")
                                saldo_inicial = sheet.cell_value(3, 5)
                                print(f"Valor bruto encontrado em F4: {saldo_inicial}")
                                print(f"Tipo do valor: {type(saldo_inicial)}")

                                try:
                                    if isinstance(saldo_inicial, str):
                                        print("Convertendo saldo inicial de string para float...")
                                        saldo_inicial = float(saldo_inicial.replace(".", "").replace(",", "."))
                                    saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                                except ValueError:
                                    saldo_inicial_frmt = ""
                                    saldo_inicial = 0.0
                                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                                if not resposta:
                                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                                    if saldo_inicial_manual:
                                        try:
                                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                                        except ValueError:
                                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                                            return
                                    else:
                                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                                        return
                    
                                print("Atualizando campo de saldo inicial na interface...")
                                self.saldo_inicial_entry.delete(0, tk.END)
                                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                                saldo_final_calculado = saldo_inicial

                                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                                print(f"Total de linhas na planilha: {sheet.nrows}")

                                for row in range(8, sheet.nrows):
                                    try:
                                        print(f"\nProcessando linha {row+1}:")
                                        data = sheet.cell_value(row, 0)
                                        print(f"Data encontrada: {data} (tipo: {type(data)})")

                                        historico = sheet.cell_value(row, 1)
                                        num_doc = sheet.cell_value(row, 4)
                                        valor = sheet.cell_value(row, 5)

                                        print(f"Valores lidos:")
                                        print(f"  Histórico: {historico}")
                                        print(f"  N° Doc: {num_doc}")
                                        print(f"  Valor: {valor}")

                                        print(f"Processando descrições...")
                                        valor_celula = sheet.cell_value(row, 1)
                                        if isinstance(valor_celula, str) and valor_celula.strip() in ["SALDO POUPANCA PLUS", "CONTA VINCULADA"]:
                                            print("Saldo/conta encontrada, pulando linha...")
                                            continue

                                        def converter_para_float(valor):
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

                                        valor_total = converter_para_float(valor)
                                        valor_formatado = formatar_valor_brasileiro(valor_total)
                                        print(f"Valor total calculado: {valor_formatado}")

                                        if isinstance(data, float):
                                            data = xlrd.xldate_as_datetime(data, wb.datemode).strftime('%d/%m/%Y')
                                        elif isinstance(data, str):
                                            if re.match(r'\d{2}/\d{2}', data):  #? ex: 02/01
                                                try:
                                                    data = datetime.strptime(f"{data}/{ano_extrato}", "%d/%m/%Y").strftime("%d/%m/%Y")
                                                except Exception as e:
                                                    print(f"Erro ao converter data '{data}': {e}")
                                            elif re.match(r'\d{2}/[a-zA-Z]{3}', data):  #? ex: 02/jan
                                                data = converter_data_pt(data, ano_extrato)
                            
                                        print("Adicionando linha aos dados importados...")
                                        dados_importados.append([
                                            data, historico, num_doc, valor_formatado, "",
                                            "", "", "", "", "", "", "", ""
                                        ])

                                        saldo_final_calculado += valor_total
                                        saldo_final_calculado = round(saldo_final_calculado, 2)
                                        print(f"Novo saldo calculado: {saldo_final_calculado}")
                            
                                    except Exception as e:
                                        print(f"ERRO ao processar linha {row+1}:")
                                        print(f"Detalhes do erro: {str(e)}")
                                        traceback.print_exc()
                                        continue

                            else: #! SE O ARQUIVO É XLSX
                                print("\n=== PROCESSANDO ARQUIVO XLSX ===")
                                wb = openpyxl.load_workbook(arquivo, data_only=True)
                                sheet = wb.active
                                print(f"Planilha ativa: {sheet.title}")

                                texto_periodo = sheet["B6"].value
                                print(f"Texto do período: {texto_periodo}")

                                def extrair_ano(texto):
                                    match = re.search(r'\d{2}/\d{2}/(\d{4})', texto)
                                    if match:
                                        return int(match.group(1))
                                    return datetime.now().year
                                
                                ano_extrato = extrair_ano(texto_periodo)
                                print(f"Ano extraído do período: {ano_extrato}")

                                meses_pt = {
                                    'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
                                    'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
                                }

                                def converter_data_pt(data_str, ano):
                                    try:
                                        dia, mes_abrev = data_str.strip().split('/')
                                        mes = meses_pt.get(mes_abrev.lower())
                                        if mes:
                                            data = datetime(ano, mes, int(dia))
                                            return data.strftime('%d/%m/%Y')
                                    except Exception as e:
                                        print(f"Erro ao converter data '{data_str}': {e}")
                                    return data_str

                                print("\nBuscando saldo inicial...")
                                saldo_inicial_celula = sheet['F9'].value
                                print(f"Valor bruto encontrado em F9: {saldo_inicial_celula}")
                    
                                if isinstance(saldo_inicial_celula, str):
                                    saldo_inicial = float(saldo_inicial_celula.replace(".", "").replace(",", "."))
                                else:
                                    saldo_inicial = float(saldo_inicial_celula)
                        
                                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")
                    
                                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                                if not resposta:
                                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                                    if saldo_inicial_manual:
                                        try:
                                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                                        except ValueError:
                                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                                            return
                                    else:
                                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                                        return

                                self.saldo_inicial_entry.delete(0, tk.END)
                                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)
                    
                                saldo_final_calculado = saldo_inicial
                    
                                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                                for row in range(10, sheet.max_row + 1):
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

                                        print(f"Processando valores...")
                                        valor_celula = sheet.cell(row=row, column=6).value
                                        if valor_celula == "0,00":
                                            print("Valor 0,00 encontrado, pulando linha...")
                                            continue

                                        print(f"Processando descrições...")
                                        valor_celula = sheet.cell(row=row, column=2).value
                                        if valor_celula == "SALDO POUPANCA PLUS":
                                            print("Saldo da poupança encontrado, pulando linha...")
                                            continue
                                
                                        historico = sheet.cell(row=row, column=2).value
                                        num_doc = sheet.cell(row=row, column=5).value
                                        valor = sheet.cell(row=row, column=6).value
                            
                                        print(f"Valores lidos:")
                                        print(f"  Histórico: {historico}")
                                        print(f"  Nº Doc: {num_doc}")
                                        print(f"  Valor: {valor}")
                            
                                        def converter_para_float(valor):
                                            if valor is None or valor == "":
                                                return 0.0
                                            if isinstance(valor, str):
                                                valor = valor.replace(".", "").replace(",", ".")
                                            try:
                                                return float(valor)
                                            except ValueError:
                                                return 0.0
                            
                                        print(f"Valor total calculado: {valor_total}")

                                        if isinstance(data, datetime):
                                            data = data.strftime("%d/%m/%Y")
                                        elif isinstance(data, str):
                                            if re.match(r"\d{2}/[a-zA-Z]{3}", data):  #? ex: 02/jan
                                                data = converter_data_pt(data, ano_extrato)
                                            elif re.match(r"\d{2}/\d{2}", data):  #? ex: 02/01
                                                try:
                                                    data = datetime.strptime(f"{data}/{ano_extrato}", "%d/%m/%Y").strftime("%d/%m/%Y")
                                                except Exception as e:
                                                    print(f"Erro ao converter data '{data}': {e}")
                            
                                        dados_importados.append([
                                            data, historico, num_doc, valor_total, valor,
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

                            for i, dados in enumerate(dados_importados):
                                tag = 'linha_par' if i % 2 == 0 else 'linha_impar'
                                self.tree.insert("", "end", values=dados, tags=(tag,))
                    
                            print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
                            print(f"Total de linhas processadas: {len(dados_importados)}")

                        except Exception as e:
                            print("\n=== ERRO FATAL ===")
                            print(f"Erro: {str(e)}")
                            print("stack trace:")
                            traceback.print_exc()
                            messagebox.showerror("Erro",
                                "Erro ao processar o arquivo. Verifique se:\n\n" +
                                "1. O arquivo está no formato correto\n" +
                                "2. O arquivo não está em modo de exibição protegida\n" +
                                "3. O arquivo está fechado no Excel\n\n" +
                                f"Erro: {str(e)}")
                            return

                    else: #! LÓGICA PARA ARQUIVO XLS/XLSX, NOVO FORMATO E CONTA CORRENTE
                        print("Processando XLS/XLSX, novo formato do Safra, conta corrente.")
                        try:
                            dados_importados = []
                            saldo_final_calculado = 0

                            if extensao == "xls": #! SE O ARQUIVO É XLS
                                print("\n=== PROCESSANDO ARQUIVO XLS ===")
                                import xlrd
                                wb = xlrd.open_workbook(arquivo)
                                sheet = wb.sheet_by_index(0)
                                print(f"Planilha aberta: {sheet.name}")

                                texto_periodo = sheet.cell_value(5, 0)
                                print(f"Texto do período: {texto_periodo}")

                                def extrair_ano(texto):
                                    match = re.search(r'\d{2}/\d{2}/(\d{4})', texto)
                                    if match:
                                        return int(match.group(1))
                                    return datetime.now().year
                                
                                ano_extrato = extrair_ano(texto_periodo)
                                print(f"Ano extraído do período: {ano_extrato}")

                                meses_pt = {
                                    'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
                                    'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
                                }

                                def converter_data_pt(data_str, ano):
                                    try:
                                        dia, mes_abrev = data_str.strip().split('/')
                                        mes = meses_pt.get(mes_abrev.lower())
                                        if mes:
                                            data = datetime(ano, mes, int(dia))
                                            return data.strftime('%d/%m/%Y')
                                    except Exception as e:
                                        print(f"Erro ao converter data '{data_str}': {e}")
                                    return data_str

                                print(f"Dimensões: {sheet.nrows} linhas x {sheet.ncols} colunas")

                                print("\nBuscando saldo inicial...")
                                saldo_inicial = sheet.cell_value(12, 7)
                                print(f"Valor bruto encontrado em F4: {saldo_inicial}")
                                print(f"Tipo do valor: {type(saldo_inicial)}")

                                if isinstance(saldo_inicial, str):
                                    print("Convertendo saldo inicial de string para float...")
                                    saldo_inicial = float(saldo_inicial.replace(".", "").replace(",", "."))
                                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                                if not resposta:
                                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                                    if saldo_inicial_manual:
                                        try:
                                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                                        except ValueError:
                                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                                            return
                                    else:
                                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                                        return
                    
                                print("Atualizando campo de saldo inicial na interface...")
                                self.saldo_inicial_entry.delete(0, tk.END)
                                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                                saldo_final_calculado = saldo_inicial

                                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                                print(f"Total de linhas na planilha: {sheet.nrows}")

                                for row in range(13, sheet.nrows):
                                    try:
                                        print(f"\nProcessando linha {row+1}:")
                                        data = sheet.cell_value(row, 0)
                                        print(f"Data encontrada: {data} (tipo: {type(data)})")

                                        historico = sheet.cell_value(row, 3)
                                        num_doc = sheet.cell_value(row, 5)
                                        valor = sheet.cell_value(row, 6)

                                        print(f"Valores lidos:")
                                        print(f"  Histórico: {historico}")
                                        print(f"  N° Doc: {num_doc}")
                                        print(f"  Valor: {valor}")

                                        if valor is None or str(valor).strip() == "":
                                            print("Valor vazio, pulando linha...")
                                            continue

                                        def converter_para_float(valor):
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

                                        valor_total = converter_para_float(valor)
                                        valor_formatado = formatar_valor_brasileiro(valor_total)
                                        print(f"Valor total calculado: {valor_formatado}")

                                        if isinstance(data, float):
                                            data = xlrd.xldate_as_datetime(data, wb.datemode).strftime('%d/%m/%Y')
                                        elif isinstance(data, str):
                                            if re.match(r'\d{2}/\d{2}', data):  #? ex: 02/01
                                                try:
                                                    data = datetime.strptime(f"{data}/{ano_extrato}", "%d/%m/%Y").strftime("%d/%m/%Y")
                                                except Exception as e:
                                                    print(f"Erro ao converter data '{data}': {e}")
                                            elif re.match(r'\d{2}/[a-zA-Z]{3}', data):  #? ex: 02/jan
                                                data = converter_data_pt(data, ano_extrato)
                            
                                        print("Adicionando linha aos dados importados...")
                                        dados_importados.append([
                                            data, historico, num_doc, valor_formatado, "",
                                            "", "", "", "", "", "", "", ""
                                        ])

                                        saldo_final_calculado += valor_total
                                        print(f"Novo saldo calculado: {saldo_final_calculado}")
                            
                                    except Exception as e:
                                        print(f"ERRO ao processar linha {row+1}:")
                                        print(f"Detalhes do erro: {str(e)}")
                                        traceback.print_exc()
                                        continue

                            else: #! SE O ARQUIVO É XLSX
                                print("\n=== PROCESSANDO ARQUIVO XLSX ===")
                                wb = openpyxl.load_workbook(arquivo, data_only=True)
                                sheet = wb.active
                                print(f"Planilha ativa: {sheet.title}")

                                texto_periodo = sheet["A6"].value
                                print(f"Texto do período: {texto_periodo}")

                                def extrair_ano(texto):
                                    match = re.search(r'\d{2}/\d{2}/(\d{4})', texto)
                                    if match:
                                        return int(match.group(1))
                                    return datetime.now().year
                                
                                ano_extrato = extrair_ano(texto_periodo)
                                print(f"Ano extraído do período: {ano_extrato}")

                                meses_pt = {
                                    'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
                                    'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
                                }

                                def converter_data_pt(data_str, ano):
                                    try:
                                        dia, mes_abrev = data_str.strip().split('/')
                                        mes = meses_pt.get(mes_abrev.lower())
                                        if mes:
                                            data = datetime(ano, mes, int(dia))
                                            return data.strftime('%d/%m/%Y')
                                    except Exception as e:
                                        print(f"Erro ao converter data '{data_str}': {e}")
                                    return data_str

                                print("\nBuscando saldo inicial...")
                                saldo_inicial_celula = sheet['H13'].value
                                print(f"Valor bruto encontrado em H13: {saldo_inicial_celula}")

                                if saldo_inicial_celula is None:
                                    messagebox.showerror("Erro", "A célula do saldo inicial está vazia.")
                                    return

                                if isinstance(saldo_inicial_celula, str):
                                    saldo_inicial = float(saldo_inicial_celula.replace(".", "").replace(",", "."))
                                else:
                                    saldo_inicial = float(saldo_inicial_celula)

                                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")
                    
                                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                                if not resposta:
                                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                                    if saldo_inicial_manual:
                                        try:
                                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                                        except ValueError:
                                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                                            return
                                    else:
                                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                                        return
                                    
                                self.saldo_inicial_entry.delete(0, tk.END)
                                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                                saldo_final_calculado = saldo_inicial

                                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                                for row in range(14, sheet.max_row + 1):
                                    try:
                                        print(f"\nProcessando linha {row}:")
                                        data = sheet.cell(row=row, column=1).value
                                        print(f"Data encontrada: {data}")

                                        valor = sheet.cell(row=row, column=7).value
                                        if not valor:
                                            print("Linha vazia, pulando...")
                                        
                                        if isinstance(data, str) and "total" in data.lower():
                                            print("Encontrada linha de total, parando processamento...")
                                            break

                                        print(f"Processando valores...")
                                        valor_celula = sheet.cell(row=row, column=7).value
                                        if valor_celula == "0,00":
                                            print("Valor 0,00 encontrado, pulando linha...")
                                            continue

                                        print(f"Processando descrições...")
                                        valor_celula = sheet.cell(row=row, column=4).value
                                        if valor_celula == "SALDO CONTA CORRENTE":
                                            print("Saldo da poupança encontrado, pulando linha...")
                                            continue

                                        historico = sheet.cell(row=row, column=4).value
                                        num_doc = sheet.cell(row=row, column=6).value
                                        valor = sheet.cell(row=row, column=7).value

                                        print(f"Valores lidos:")
                                        print(f"  Histórico: {historico}")
                                        print(f"  Nº Doc: {num_doc}")
                                        print(f"  Valor: {valor}")

                                        def converter_para_float(valor):
                                            if valor is None or valor == "":
                                                return 0.0
                                            if isinstance(valor, str):
                                                valor = valor.replace(".", "").replace(",", ".")
                                            try:
                                                return float(valor)
                                            except ValueError:
                                                return 0.0
                                            
                                        valor_total = valor
                                        print(f"Valor total calculado: {valor_total}")

                                        if isinstance(data, datetime):
                                            data = data.strftime("%d/%m/%Y")
                                        elif isinstance(data, str):
                                            if re.match(r"\d{2}/[a-zA-Z]{3}", data):  #? ex: 02/jan
                                                data = converter_data_pt(data, ano_extrato)
                                            elif re.match(r"\d{2}/\d{2}", data):  #? ex: 02/01
                                                try:
                                                    data = datetime.strptime(f"{data}/{ano_extrato}", "%d/%m/%Y").strftime("%d/%m/%Y")
                                                except Exception as e:
                                                    print(f"Erro ao converter data '{data}': {e}")

                                        dados_importados.append([
                                            data, historico, num_doc, valor_total, valor,
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

                            for i, dados in enumerate(dados_importados):
                                tag = 'linha_par' if i % 2 == 0 else 'linha_impar'
                                self.tree.insert("", "end", values=dados, tags=(tag,))

                            print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
                            print(f"Total de linhas processadas: {len(dados_importados)}")

                        except Exception as e:
                            print("\n=== ERRO FATAL ===")
                            print(f"Erro: {str(e)}")
                            print("stack trace:")
                            traceback.print_exc()
                            messagebox.showerror("Erro",
                                "Erro ao processar o arquivo. Verifique se:\n\n" +
                                "1. O arquivo está no formato correto\n" +
                                "2. O arquivo não está em modo de exibição protegida\n" +
                                "3. O arquivo está fechado no Excel\n\n" +
                                f"Erro: {str(e)}")
                            return

                else: #! LÓGICA PARA ARQUIVO XLS/XLSX, FORMATO ANTIGO
                    print("Processando XLS/XLSX, formato antigo do Safra")

            elif extensao == "pdf":
                if novo_formato:
                    if conta_vinculada: #! LÓGICA PARA ARQUIVO PDF, NOVO FORMATO E CONTA VINCULADA
                        print("Processando PDF, novo formato do Safra, conta vinculada.")

                    else: #! LÓGICA PARA ARQUIVO PDF, NOVO FORMATO E CONTA CORRENTE
                        print("Processando PDF, novo formato do Safra, conta corrente.")

                else: #! LÓGICA PARA ARQUIVO PDF, FORMATO ANTIGO
                    print("Processando PDF, formato antigo do Safra.")


            else:
                print("Formato de arquivo não suportado.")
        else:
            print("Nenhuma resposta específica fornecida para Safra.")

    def acao_santander(self, arquivo):
        print("\n=== INÍCIO DO PROCESSAMENTO: SANTANDER ===")
        print(f"Arquivo Recebido: {arquivo}")
        def formatar_valor_brasileiro(valor):
            try:
                return locale.format_string("%.2f", float(valor), grouping=True)
            except:
                return valor
        try:
            extensao = arquivo.lower().split('.')[-1]
            print(f"Extensão detectada: {extensao}")
            dados_importados = []
            saldo_final_calculado = 0

            if extensao == 'xls': #! SE O ARQUIVO É XLS
                print("\n=== PROCESSANDO ARQUIVO XLS ===")
                wb = xlrd.open_workbook(arquivo)
                sheet = wb.sheet_by_index(0)
                print(f"Planilha aberta: {sheet.name}")
                print(f"Dimensões? {sheet.nrows} linhas x {sheet.ncols} colunas")

                print("\nBuscando saldo inicial...")
                saldo_inicial = sheet.cell_value(3, 5)
                print(f"Valor bruto encontrado em F4: {saldo_inicial}")
                print(f"Tipo do valor: {type(saldo_inicial)}")

                if isinstance(saldo_inicial, str):
                    print("Convertendo saldo inicial de string para float...")
                    saldo_inicial = float(saldo_inicial.replace(".", "").replace(",", "."))
                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                if not resposta:
                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                    if saldo_inicial_manual:
                        try:
                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                        except ValueError:
                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                            return
                    else:
                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                        return
                    
                print("Atualizando campo de saldo inicial na interface...")
                self.saldo_inicial_entry.delete(0, tk.END)
                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                saldo_final_calculado = saldo_inicial

                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                print(f"Total de linhas na planilha: {sheet.nrows}")

                for row in range(8, sheet.nrows):
                    try:
                        print(f"\nProcessando linha {row+1}:")
                        data = sheet.cell_value(row, 0)
                        print(f"Data encontrada: {data} (tipo: {type(data)})")

                        historico = sheet.cell_value(row, 2)
                        num_doc = sheet.cell_value(row, 3)
                        saldo = sheet.cell_value(row, 5)
                        valor = sheet.cell_value(row, 4)

                        print(f"Valores lidos:")
                        print(f"  Histórico: {historico}")
                        print(f"  N° Doc: {num_doc}")
                        print(f"  Valor: {valor}")
                        print(f"  Saldo: {saldo}")

                        if valor is None or str(valor).strip() == "":
                            print("Valor vazio, pulando linha...")
                            continue

                        def converter_para_float(valor):
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
                            
                        valor_total = converter_para_float(valor)
                        valor_formatado = formatar_valor_brasileiro(valor_total)
                        print(f"Valor total calculado: {valor_formatado}")

                        saldo_total = converter_para_float(saldo)
                        saldo_formatado = formatar_valor_brasileiro(saldo_total)
                        print(f"Saldo total calculado: {saldo_formatado}")

                        if isinstance(data, float):
                            print("Convertendo data de float para string...")
                            data = xlrd.xldate_as_datetime(data, wb.datemode).strftime('%d/%m/%Y')
                            print(f"Data convertida: {data}")
                            
                        print("Adicionando linha aos dados importados...")
                        dados_importados.append([
                            data, historico, num_doc, valor_formatado, saldo_formatado,
                            "", "", "", "", "", "", "", ""
                        ])

                        saldo_final_calculado += valor_total
                        print(f"Novo saldo calculado: {saldo_final_calculado}")
                            
                    except Exception as e:
                        print(f"ERRO ao processar linha {row+1}:")
                        print(f"Detalhes do erro: {str(e)}")
                        traceback.print_exc()
                        continue

            else: #! SE O ARQUIVO É XLSX
                print("\n=== PROCESSANDO ARQUIVO XLSX ===")
                wb = openpyxl.load_workbook(arquivo, data_only=True)
                sheet = wb.active
                print(f"Planilha ativa: {sheet.title}")

                print("\nBuscando saldo inicial...")
                saldo_inicial_celula = sheet['F4'].value
                print(f"Valor bruto encontrado em F4: {saldo_inicial_celula}")

                if isinstance(saldo_inicial_celula, str):
                    saldo_inicial = float(saldo_inicial_celula.replace(".", "").replace(",", "."))
                else:
                    saldo_inicial = float(saldo_inicial_celula)

                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                if not resposta:
                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                    if saldo_inicial_manual:
                        try:
                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                        except ValueError:
                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                            return
                    else:
                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                        return
                    
                self.saldo_inicial_entry.delete(0, tk.END)
                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                saldo_final_calculado = saldo_inicial

                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                for row in range(4, sheet.max_row + 1):
                    try:
                        print(f"\nProcessando linha {row}:")
                        data = sheet.cell(row=row, column=1).value
                        print(f"Data encontrada: {data}")

                        historico = sheet.cell(row=row, column=3).value
                        num_doc = sheet.cell(row=row, column=4).value
                        valor = sheet.cell(row=row, column=5).value
                        saldo = sheet.cell(row=row, column=6).value

                        print(f"Valores lidos:")
                        print(f"  Histórico: {historico}")
                        print(f"  N° Doc: {num_doc}")
                        print(f"  Valor: {valor}")
                        print(f"  Saldo: {saldo}")

                        def converter_para_float(valor):
                            if valor is None or valor == "":
                                return 0.0
                            if isinstance(valor, str):
                                valor = valor.replace(".", "").replace(",", ".")
                            try:
                                return float(valor)
                            except ValueError:
                                return 0.0
                            
                        valor_total = converter_para_float(valor)
                        valor_formatado = formatar_valor_brasileiro(valor_total)
                        print(f"Valor total calculado: {valor_total}")

                        saldo_total = converter_para_float(saldo)
                        saldo_formatado = formatar_valor_brasileiro(saldo_total)
                        print(f"Saldo total formatado: {saldo_total}")

                        dados_importados.append([
                            data, historico, num_doc, valor_formatado, saldo_formatado,
                            "", "", "", "", "", "", "", ""
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

            for i, dados in enumerate(dados_importados):
                tag = 'linha_par' if i % 2 == 0 else 'linha_impar'
                self.tree.insert("", "end", values=dados, tags=(tag,))
                    
            print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
            print(f"Total de linhas processadas: {len(dados_importados)}")

        except Exception as e:
            print("\n=== ERRO FATAL ===")
            print(f"Erro: {str(e)}")
            print("stack trace:")
            traceback.print_exc()
            messagebox.showerror("Erro",
                "Erro ao processar o arquivo. Verifique se:\n\n" +
                "1. O arquivo está no formato correto\n" +
                "2. O arquivo não está em modo de exibição protegida\n" +
                "3. O arquivo está fechado no Excel\n\n" +
                f"Erro: {str(e)}")
            return

    def acao_daycoval(self, arquivo):
        print("\n=== INÍCIO DO PROCESSAMENTO: DAYCOVAL ===")
        print(f"Arquivo Recebido: {arquivo}")
        try:
            extensao = arquivo.lower().split('.')[-1]
            print(f"Extensão detectada: {extensao}")
            dados_importados = []
            saldo_final_calculado = 0
            def formatar_valor_brasileiro(valor):
                try:
                    return locale.format_string("%.2f", float(valor), grouping=True)
                except:
                    return valor

            if extensao == 'xls': #! SE O ARQUIVO É XLS
                print("\n=== PROCESSANDO ARQUIVO XLS ===")
                wb = xlrd.open_workbook(arquivo)
                sheet = wb.sheet_by_index(0)
                print(f"Planilha aberta: {sheet.name}")

                meses_pt = {
                    'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
                    'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
                }

                def converter_data_pt(data_str, ano):
                    try:
                        dia, mes_abrev = data_str.strip().split('/')
                        mes = meses_pt.get(mes_abrev.lower())
                        if mes:
                            data = datetime(ano, mes, int(dia))
                            return data.strftime('%d/%m/%Y')
                    except Exception as e:
                        print(f"Erro ao converter data '{data_str}': {e}")
                        return data_str
                    
                texto_periodo = sheet.cell_value(5, 0)
                print(f"Texto do período: {texto_periodo}")
                def extrair_ano(texto):
                    match = re.search(r'\d{2}/\d{2}/(\d{4})', texto)
                    if match:
                        return datetime.now().year
                    ano_extrato = extrair_ano(texto_periodo)
                    print(f"Ano extraído do período: {ano_extrato}")

                print(f"Dimensões: {sheet.nrows} linhas x {sheet.ncols} colunas")

                print("\nBuscando saldo inicial...")
                saldo_inicial = sheet.cell_value(9, 5)
                print(f"Valor bruto encontrado em F10: {saldo_inicial}")
                print(f"Tipo do valor: {type(saldo_inicial)}")

                if isinstance(saldo_inicial, str):
                    print("Convertendo saldo inicial de string para float...")
                    saldo_inicial = float(saldo_inicial.replace(".", "").replace(",", "."))
                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                if not resposta:
                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                    if saldo_inicial_manual:
                        try:
                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                        except ValueError:
                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                            return
                    else:
                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                        return
                    
                print("Atualizando campo de saldo inicial na interface...")
                self.saldo_inicial_entry.delete(0, tk.END)
                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                saldo_final_calculado = saldo_inicial

                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                print(f"Total de linhas na planilha: {sheet.nrows}")

                for row in range(10, sheet.nrows):
                    try:
                        print(f"\nProcessando linha {row+1}:")
                        data = sheet.cell_value(row, 0)
                        print(f"Data encontrada: {data} (tipo: {type(data)})")

                        historico = sheet.cell_value(row, 2)
                        num_doc = sheet.cell_value(row, 1)
                        debito = sheet.cell_value(row, 3)
                        credito = sheet.cell_value(row, 4)
                        saldo = sheet.cell_value(row, 5)
                        valor = sheet.cell_value(row, 4)

                        print(f"Valores lidos:")
                        print(f"  Histórico: {historico}")
                        print(f"  N° Doc: {num_doc}")
                        print(f"  Débito: {debito}")
                        print(f"  Crédito: {credito}")
                        print(f"  Valor: {valor}")
                        print(f"  Saldo: {saldo}")

                        if valor is None or str(valor).strip() == "":
                            print("Valor vazio, pulando linha...")
                            continue

                        def converter_para_float(valor):
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
                            
                        valor_credito = converter_para_float(credito)
                        valor_debito = converter_para_float(debito)
                        valor_total = valor_credito + valor_debito
                        valor_formatado = formatar_valor_brasileiro(valor_total)
                        print(f"Valor total calculado: {valor_formatado}")

                        saldo_total = converter_para_float(saldo)
                        saldo_formatado = formatar_valor_brasileiro(saldo_total)
                        print(f"Saldo total calculado: {saldo_formatado}")

                        if isinstance(data, float):
                            print("Convertendo data float padrão Excel.")
                            data = xlrd.xldate_as_datetime(data, wb.datemode).strftime('%d/%m/%Y')
                            print(f"Data convertido: {data}")
                        elif isinstance(data, str):
                            print(f"Convertendo data string com mês pt-br: {data}")
                            data = converter_data_pt(data, ano_extrato)
                            print(f"Data convertida: {data}")
                            
                        print("Adicionando linha aos dados importados...")
                        dados_importados.append([
                            data, historico, num_doc, valor_formatado, saldo_formatado,
                            "", "", "", "", "", "", "", ""
                        ])

                        saldo_final_calculado += valor_total
                        print(f"Novo saldo calculado: {saldo_final_calculado}")
                            
                    except Exception as e:
                        print(f"ERRO ao processar linha {row+1}:")
                        print(f"Detalhes do erro: {str(e)}")
                        traceback.print_exc()
                        continue

            else:  #! SE O ARQUIVO É XLSX
                print("\n=== PROCESSANDO ARQUIVO XLSX ===")
                wb = openpyxl.load_workbook(arquivo, data_only=True)
                sheet = wb.active
                print(f"Planilha ativa: {sheet.title}")

                meses_pt = {
                    'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
                    'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12
                }

                def converter_data_pt(data_str, ano):
                    try:
                        dia, mes_abrev = data_str.strip().split('/')
                        mes = meses_pt.get(mes_abrev.lower())
                        if mes:
                            data = datetime(int(ano), mes, int(dia))
                            return data.strftime('%d/%m/%Y')
                    except Exception as e:
                        print(f"Erro ao converter data '{data_str}': {e}")
                    return data_str
                
                texto_periodo = sheet["A6"].value
                print(f"Texto do período: {texto_periodo}")
                def extrair_ano(texto):
                    match = re.search(r'\d{2}/\d{2}/(\d{4})', texto)
                    if match:
                        return int(match.group(1))
                    return datetime.now().year
                ano_extrato = extrair_ano(texto_periodo)
                print(f"Ano extraído do período: {ano_extrato}")

                print("\nBuscando saldo inicial...")
                saldo_inicial_celula = sheet['F10'].value
                print(f"Valor bruto encontrado em F10: {saldo_inicial_celula}")

                if isinstance(saldo_inicial_celula, str):
                    saldo_inicial = float(saldo_inicial_celula.replace(".", "").replace(",", "."))
                else:
                    saldo_inicial = float(saldo_inicial_celula)

                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                if not resposta:
                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                    if saldo_inicial_manual:
                        try:
                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                        except ValueError:
                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                            return
                    else:
                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                        return
                    
                self.saldo_inicial_entry.delete(0, tk.END)
                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                saldo_final_calculado = saldo_inicial

                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                for row in range(11, sheet.max_row + 1):
                    try:
                        print(f"\nProcessando linha {row}:")
                        data = sheet.cell(row=row, column=1).value
                        print(f"Data encontrada: {data}")

                        historico = sheet.cell(row=row, column=3).value
                        num_doc = sheet.cell(row=row, column=2).value
                        debito = sheet.cell(row=row, column=4).value
                        credito = sheet.cell(row=row, column=5).value
                        valor = sheet.cell(row=row, column=5).value
                        saldo = sheet.cell(row=row, column=6).value

                        print(f"Valores lidos:")
                        print(f"  Histórico: {historico}")
                        print(f"  N° Doc: {num_doc}")
                        print(f"  Débito: {debito}")
                        print(f"  Crédito: {credito}")
                        print(f"  Valor: {valor}")
                        print(f"  Saldo: {saldo}")

                        def converter_para_float(valor):
                            if valor is None or valor == "":
                                return 0.0
                            if isinstance(valor, str):
                                valor = valor.replace(".", "").replace(",", ".")
                            try:
                                return float(valor)
                            except ValueError:
                                return 0.0
                            
                        valor_credito = converter_para_float(credito)
                        valor_debito = converter_para_float(debito)
                        valor_total = valor_credito + valor_debito
                        valor_formatado = formatar_valor_brasileiro(valor_total)
                        print(f"Valor total calculado: {valor_formatado}")

                        saldo_total = converter_para_float(saldo)
                        saldo_formatado = formatar_valor_brasileiro(saldo_total)
                        print(f"Saldo total formatado: {saldo_total}")

                        if isinstance(data, datetime):
                            print("Data já está no formato datetime...")
                            data = data.strftime("%d/%m/%Y")
                        elif isinstance(data, str):
                            print(f"Convertendo data abreviada em pt-br: {data}")
                            data = converter_data_pt(data, ano_extrato)
                            print(f"Data convertido: {data}")

                        dados_importados.append([
                            data, historico, num_doc, valor_formatado, saldo_formatado,
                            "", "", "", "", "", "", "", ""
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

            for i, dados in enumerate(dados_importados):
                tag = 'linha_par' if i % 2 == 0 else 'linha_impar'
                self.tree.insert("", "end", values=dados, tags=(tag,))

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

    def acao_itau(self, arquivo):
        print("\n=== INÍCIO DO PROCESSAMENTO: ITAÚ ===")
        print(f"Arquivo Recebido: {arquivo}")
        def formatar_valor_brasileiro(valor):
            try:
                return locale.format_string("%.2f", float(valor), grouping=True)
            except:
                return valor
        try:
            extensao = arquivo.lower().split('.')[-1]
            print(f"Extensão detectada: {extensao}")
            dados_importados = []
            saldo_final_calculado = 0

            if extensao == "xls": #! SE O ARQUIVO É XLS
                print("\n=== PROCESSANDO ARQUIVO XLS ===")
                wb = xlrd.open_workbook(arquivo)
                sheet = wb.sheet_by_index(0)
                print(f"Planilha aberta: {sheet.name}")
                print(f"Dimensões? {sheet.nrows} linhas x {sheet.ncols} colunas")

                print("\nBuscando saldo inicial...")
                saldo_inicial = sheet.cell_value(7, 7)
                print(f"Valor bruto encontrado em H8: {saldo_inicial}")
                print(f"Tipo do valor: {type(saldo_inicial)}")

                if isinstance(saldo_inicial, str):
                    print("Convertendo saldo inicial de string para float...")
                    saldo_inicial = float(saldo_inicial.replace(".", "").replace(",", "."))
                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                if not resposta:
                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                    if saldo_inicial_manual:
                        try:
                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                        except ValueError:
                                messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                                return
                    else:
                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                        return
                    
                print("Atualizando campo de saldo inicial na interface...")
                self.saldo_inicial_entry.delete(0, tk.END)
                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                saldo_final_calculado = saldo_inicial

                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                print(f"Total de linhas na planilha: {sheet.nrows}")

                for row in range(8, sheet.nrows):
                    try:
                        print(f"\nProcessando linha {row+1}:")
                        data = sheet.cell_value(row, 1)
                        print(f"Data encontrada: {data} (tipo: {type(data)})")

                        historico = sheet.cell_value(row, 4)
                        valor = sheet.cell_value(row, 6)

                        print(f"Valores lidos:")
                        print(f"  Histórico: {historico}")
                        print(f"  Valor: {valor}")

                        if valor is None or str(valor).strip() == "":
                            print("Valor vazio, pulando linha...")
                            continue

                        def converter_para_float(valor):
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
                        
                        valor_total = valor
                        print(f"Valor total calculado: {valor_total}")

                        if isinstance(data, float):
                            print("Convertendo data de float para string...")
                            data = xlrd.xldate_as_datetime(data, wb.datemode).strftime('%d/%m/%Y')
                            print(f"Data convertida: {data}")
                            
                        print("Adicionando linha aos dados importados...")
                        dados_importados.append([
                            data, historico, "", valor_total, valor,
                            "", "", "", "", "", "", "", ""
                        ])

                        saldo_final_calculado += valor_total
                        print(f"Novo saldo calculado: {saldo_final_calculado}")
                            
                    except Exception as e:
                        print(f"ERRO ao processar linha {row+1}:")
                        print(f"Detalhes do erro: {str(e)}")
                        traceback.print_exc()
                        continue

            else: #! SE O ARQUIVO É XLSX
                print("\n=== PROCESSANDO ARQUIVO XLSX ===")
                wb = openpyxl.load_workbook(arquivo, data_only=True)
                sheet = wb.active
                print(f"Planilha ativa: {sheet.title}")

                print("\nBuscando saldo inicial...")
                saldo_inicial_celula = sheet['H8'].value
                print(f"Valor bruto encontrado em H8: {saldo_inicial_celula}")

                if isinstance(saldo_inicial_celula, str):
                    saldo_inicial = float(saldo_inicial_celula.replace(".", "").replace(",", "."))
                else:
                    saldo_inicial = float(saldo_inicial_celula)

                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                if not resposta:
                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                    if saldo_inicial_manual:
                        try:
                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                        except ValueError:
                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                            return
                    else:
                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                        return
                    
                self.saldo_inicial_entry.delete(0, tk.END)
                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                saldo_final_calculado = saldo_inicial

                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                for row in range(9, sheet.max_row + 1):
                    try:
                        print(f"\nProcessando linha {row}:")
                        data = sheet.cell(row=row, column=2).value

                        if isinstance(data, datetime):
                            data = data.date()

                        elif not isinstance(data, date):
                            print("Data inválida ou vazia, pulando...")
                            continue

                        data_formatada = data.strftime("%d/%m/%Y")
                        print(f"Data formatada encontrada: {data_formatada}")

                        historico = sheet.cell(row=row, column=5).value
                        valor = sheet.cell(row=row, column=7).value

                        print(f"Valores lidos:")
                        print(f"  Histórico: {historico}")
                        print(f"  Valor: {valor}")

                        if valor is None or str(valor).strip() == "":
                            print("Valor vazio, pulando linha...")
                            continue

                        def converter_para_float(valor):
                            if valor is None or valor == "":
                                return 0.0
                            if isinstance(valor, str):
                                valor = valor.replace(".", "").replace(",", ".")
                            try:
                                return float(valor)
                            except ValueError:
                                return 0.0
                            
                        valor_total = converter_para_float(valor)
                        valor_formatado = formatar_valor_brasileiro(valor_total)
                        print(f"Valor total calculado: {valor_total}")
                            
                        dados_importados.append([
                            data_formatada, historico, "", valor_formatado, "",
                            "", "", "", "", "", "", "", ""
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

            for i, dados in enumerate(dados_importados):
                tag = 'linha_par' if i % 2 == 0 else 'linha_impar'
                self.tree.insert("", "end", values=dados, tags=(tag,))
                    
            print("\n=== PROCESSAMENTO CONCLUÍDO COM SUCESSO ===")
            print(f"Total de linhas processadas: {len(dados_importados)}")

        except Exception as e:
            print("\n=== ERRO FATAL ===")
            print(f"Erro: {str(e)}")
            print("stack trace:")
            traceback.print_exc()
            messagebox.showerror("Erro",
                "Erro ao processar o arquivo. Verifique se:\n\n" +
                "1. O arquivo está no formato correto\n" +
                "2. O arquivo não está em modo de exibição protegida\n" +
                "3. O arquivo está fechado no Excel\n\n" +
                f"Erro: {str(e)}")
            return

    def acao_brasil(self, arquivo):
        print("\n=== INÍCIO DO PROCESSAMENTO: BANCO DO BRASIL ===")
        print(f"Arquivo Recebido: {arquivo}")
        def formatar_valor_brasileiro(valor):
            try:
                return locale.format_string("%.2f", float(valor), grouping=True)
            except:
                return valor
        try:
            extensao = arquivo.lower().split('.')[-1]
            print(f"Extensão detectada: {extensao}")
            dados_importados = []
            saldo_final_calculado = 0

            if extensao == 'xls': #! SE O ARQUIVO É XLS
                print("\n=== PROCESSANDO ARQUIVO XLS ===")
                wb = xlrd.open_workbook(arquivo)
                sheet = wb.sheet_by_index(0)
                print(f"Planilha aberta: {sheet.name}")
                print(f"Dimensões? {sheet.nrows} linhas x {sheet.ncols} colunas")

                print("\nBuscando saldo inicial...")
                saldo_inicial = sheet.cell_value(3, 8)
                print(f"Valor bruto encontrado em F4: {saldo_inicial}")
                print(f"Tipo do valor: {type(saldo_inicial)}")

                if isinstance(saldo_inicial, str):
                    print("Convertendo saldo inicial de string para float...")
                    saldo_inicial = float(saldo_inicial.replace(".", "").replace(",", "."))
                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                if not resposta:
                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                    if saldo_inicial_manual:
                        try:
                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                        except ValueError:
                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                            return
                    else:
                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                        return
                    
                print("Atualizando campo de saldo inicial na interface...")
                self.saldo_inicial_entry.delete(0, tk.END)
                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                saldo_final_calculado = saldo_inicial

                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                print(f"Total de linhas na planilha: {sheet.nrows}")

                for row in range(4, sheet.nrows):
                    try:
                        print(f"\nProcessando linha {row+1}:")
                        data = sheet.cell_value(row, 0)
                        print(f"Data encontrada: {data} (tipo: {type(data)})")

                        historico = sheet.cell_value(row, 7)
                        num_doc = sheet.cell_value(row, 5)
                        valor = sheet.cell_value(row, 8)

                        print(f"Valores lidos:")
                        print(f"  Histórico: {historico}")
                        print(f"  N° Doc: {num_doc}")
                        print(f"  Valor: {valor}")

                        if valor is None or str(valor).strip() == "":
                            print("Valor vazio, pulando linha...")
                            continue

                        def converter_para_float(valor):
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
                            
                        valor_total = converter_para_float(valor)
                        valor_formatado = formatar_valor_brasileiro(valor_total)
                        print(f"Valor total calculado: {valor_formatado}")

                        if isinstance(data, float):
                            print("Convertendo data de float para string...")
                            data = xlrd.xldate_as_datetime(data, wb.datemode).strftime('%d/%m/%Y')
                            print(f"Data convertida: {data}")
                            
                        print("Adicionando linha aos dados importados...")
                        dados_importados.append([
                            data, historico, num_doc, valor_formatado, "",
                            "", "", "", "", "", "", "", ""
                        ])

                        saldo_final_calculado += valor_total
                        print(f"Novo saldo calculado: {saldo_final_calculado}")
                            
                    except Exception as e:
                        print(f"ERRO ao processar linha {row+1}:")
                        print(f"Detalhes do erro: {str(e)}")
                        traceback.print_exc()
                        continue

            else: #! SE O ARQUIVO É XLSX
                print("\n=== PROCESSANDO ARQUIVO XLSX ===")
                wb = openpyxl.load_workbook(arquivo, data_only=True)
                sheet = wb.active
                print(f"Planilha ativa: {sheet.title}")

                print("\nBuscando saldo inicial...")
                saldo_inicial_celula = sheet['I4'].value
                print(f"Valor bruto encontrado em I4: {saldo_inicial_celula}")

                if isinstance(saldo_inicial_celula, str):
                    saldo_inicial = float(saldo_inicial_celula.replace(".", "").replace(",", "."))
                else:
                    saldo_inicial = float(saldo_inicial_celula)

                saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                print(f"Saldo inicial formatado: R${saldo_inicial_frmt}")

                resposta = messagebox.askyesno("Confirmação de saldo", f"O saldo inicial é de R${saldo_inicial_frmt}?")

                if not resposta:
                    print("Usuário não confirmou o saldo inicial, solicitando entrada manual.")
                    saldo_inicial_manual = simpledialog.askstring("Entrada de Saldo", "Insira o saldo inicial correto:")
                    if saldo_inicial_manual:
                        try:
                            saldo_inicial = float(saldo_inicial_manual.replace(".", "").replace(",", "."))
                            saldo_inicial_frmt = locale.format_string("%.2f", saldo_inicial, grouping=True)
                        except ValueError:
                            messagebox.showerror("Erro", "Valor de saldo inicial inválido.")
                            return
                    else:
                        messagebox.showinfo("Aviso", "Processo cancelado pelo usuário.")
                        return
                    
                self.saldo_inicial_entry.delete(0, tk.END)
                self.saldo_inicial_entry.insert(0, saldo_inicial_frmt)

                print("\n=== INICIANDO PROCESSAMENTO DAS LINHAS ===")
                for row in range(5, sheet.max_row + 1):
                    try:
                        print(f"\nProcessando linha {row}:")
                        data = sheet.cell(row=row, column=1).value
                        print(f"Data encontrada: {data}")

                        historico = sheet.cell(row=row, column=8).value
                        num_doc = sheet.cell(row=row, column=6).value
                        valor = sheet.cell(row=row, column=9).value

                        print(f"Valores lidos:")
                        print(f"  Histórico: {historico}")
                        print(f"  N° Doc: {num_doc}")
                        print(f"  Valor: {valor}")

                        def converter_para_float(valor):
                            if valor is None or valor == "":
                                return 0.0
                            if isinstance(valor, str):
                                valor = valor.replace(".", "").replace(",", ".")
                            try:
                                return float(valor)
                            except ValueError:
                                return 0.0
                            
                        valor_total = converter_para_float(valor)
                        valor_formatado = formatar_valor_brasileiro(valor_total)
                        print(f"Valor total calculado: {valor_total}")

                        dados_importados.append([
                            data, historico, num_doc, valor_formatado, "",
                            "", "", "", "", "", "", "", ""
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

            for i, dados in enumerate(dados_importados):
                tag = 'linha_par' if i % 2 == 0 else 'linha_impar'
                self.tree.insert("", "end", values=dados, tags=(tag,))

        except Exception as e:
            print("\n=== ERRO FATAL ===")
            print(f"Erro: {str(e)}")
            print("stack trace:")
            traceback.print_exc()
            messagebox.showerror("Erro",
                "Erro ao processar o arquivo. Verifique se:\n\n" +
                "1. O arquivo está no formato correto\n" +
                "2. O arquivo não está em modo de exibição protegida\n" +
                "3. O arquivo está fechado no Excel\n\n" +
                f"Erro: {str(e)}")
            return

    def confirmar_limpar_dados(self):
        resposta = messagebox.askyesno("Atenção", "Tem certeza que deseja limpar todos os dados?")
        if resposta:
            self.limpar_dados()

    def limpar_dados(self):
        self.saldo_inicial_entry.delete(0, tk.END)
        self.saldo_final_entry.delete(0, tk.END)
        self.saldo_final_calculado_entry.delete(0, tk.END)
        self.empresa_entry.delete(0, tk.END)
        self.conta_entry.delete(0, tk.END)
        self.diferenca_entry.delete(0, tk.END)
        self.banco_entry.delete(0, tk.END)
        self.agencia_conta_entry.delete(0, tk.END)
        self.saldo_final_contabil_entry.delete(0, tk.END)
        self.diferenca_extrato_bancario_entry.delete(0, tk.END)
        for i in self.tree.get_children():
            self.tree.delete(i)

    def exportar_dados(self):
        wb = Workbook()
        ws = wb.active

        #* -------------------- DEFINE OS CABEÇALHOS DAS COLUNAS -------------------- #
        colunas_exportar = ["LancamentoLC", "DataLC", "DebitoLC", "D-C/CLC", "CreditoLC",
                      "C-C/CLC", "CNPJLC", "HistoricoLC", "ValorLC"]
        
        headers = {
            "LancamentoLC": "Lançamento",
            "DataLC": "Data",
            "DebitoLC": "Débito",
            "D-C/CLC": "D-C/C",
            "CreditoLC": "Crédito",
            "C-C/CLC": "C-C/C",
            "CNPJLC": "CNPJ",
            "HistoricoLC": "Histórico",
            "ValorLC": "Valor"
        }

        #* -------------------- ITERA SOBRE OS ITENS DA TREEVIEW E ADICIONA À PLANILHA -------------------- #
        dados = []
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            dados.append([values[self.tree["columns"].index(col)] for col in colunas_exportar])

        df = pd.DataFrame(dados, columns=[headers[col] for col in colunas_exportar])

        #* ------------------- SALVA O ARQUIVO EXCEL -------------------- #
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Exportar", "Dados exportados com sucesso!")

    def limpar_1p(self):
        resposta = messagebox.askyesno("Atenção", "Tem certeza que deseja limpar os dados da classificação?")
        if resposta:
            self.limpar_classificacao()

    def limpar_classificacao(self):
        colunas_para_limpar = [
            "LancamentoLC", "DataLC", "DebitoLC", "CreditoLC",
            "HistoricoLC", "ValorLC"
        ]
        
        for item in self.tree.get_children():
            values = list(self.tree.item(item, 'values'))
            for coluna in colunas_para_limpar:
                idx = self.tree["columns"].index(coluna)
                values[idx] = ""
            self.tree.item(item, values=values)

    def fechar_janela(self):
        if hasattr(self, 'tela_selecao_conta') and self.tela_selecao_conta.winfo_exists():
            self.tela_selecao_conta.root.destroy()
        self.root.destroy()

class TelaSelecaoConta:
    def __init__(self, root, callback):
        self.root = root
        self.callback = callback
        self.root.title("Importador")
        self.root.geometry("480x210")
        self.root.config(bg='#f4f4f4')
        self.root.iconbitmap(r"C:\Users\regina.santos\Desktop\Automacao\Judite\icon.ico")

        #* -------------------- TÍTULO PRINCIPAL -------------------- #
        self.title_label = tk.Label(root, text="Qual conta bancária irá importar?", font=("Roboto", 17, "bold"), bg='#f4f4f4') 
        self.title_label.grid(row=1, column=0, columnspan=1, padx=10, pady=10, sticky="w")

        #* -------------------- CAMPOS DE INFORMAÇÕES -------------------- #
        self.label_empresa = tk.Label(root, text="Empresa:", font=("Roboto", 10), bg='#f4f4f4') 
        self.label_empresa.grid(row=2, column=0, pady=1, padx=10, sticky="w")

        self.empresas = self.carregar_empresas()
        self.combobox_empresa = ttk.Combobox(root, values=self.empresas, font=("Roboto", 10), width=63)
        self.combobox_empresa.grid(row=3, column=0, pady=1, padx=10, sticky="w")
        self.combobox_empresa.bind("<<ComboboxSelected>>", self.atualizar_contas_contabeis)
        self.combobox_empresa.bind("<KeyRelease>", self.atualizar_contas_contabeis)

        self.label_conta_contabil = tk.Label(root, text="Conta Contábil:", font=("Roboto", 10), bg='#f4f4f4')
        self.label_conta_contabil.grid(row=4, column=0, pady=1, padx=10, sticky="w")

        self.combobox_conta_contabil = ttk.Combobox(root, font=("Roboto", 10), width=63)
        self.combobox_conta_contabil.grid(row=5, column=0, pady=1, padx=10, sticky="w")

        #* -------------------- BOTÕES PRINCIPAIS -------------------- #
        self.btn_nova_empresa = tk.Button(root, text="Nova Empresa", command=self.abrir_tela_nova_empresa, font=("Roboto", 10), bg="#4CAF50", fg="white")
        self.btn_nova_empresa.grid(row=7, column=0, pady=15, padx=10, sticky="w")

        self.btn_nova_conta = tk.Button(root, text="Nova Conta", command=self.abrir_tela_nova_conta, font=("Roboto", 10), bg="#4CAF50", fg="white")
        self.btn_nova_conta.grid(row=7, column=0, pady=15, padx=109, sticky="w")

        self.btn_alterar_conta = tk.Button(root, text="Alterar Conta", command=self.alterar_conta, font=("Roboto", 10), bg="#FFC107", fg="black")
        self.btn_alterar_conta.grid(row=7, column=0, pady=15, padx=190, sticky="w")

        self.btn_ok = tk.Button(root, text="OK", command=self.confirmar, font=("Roboto", 10), bg="#4CAF50", fg="white")
        self.btn_ok.grid(row=7, column=0, columnspan=2, pady=15, padx=83, sticky="e")

        self.btn_cancelar = tk.Button(root, text="Cancelar", command=self.cancelar, font=("Roboto", 10), bg="#FF5733", fg="white")
        self.btn_cancelar.grid(row=7, column=0, columnspan=2, pady=15, padx=15, sticky="e")

    def atualizar_contas_contabeis(self, event=None):
        empresa = self.combobox_empresa.get()
        if empresa:
            codigo_empresa = empresa.split(" - ")[0]
            contas = self.carregar_contas_contabeis(codigo_empresa)
            if contas:
                self.combobox_conta_contabil['values'] = contas
            else:
                self.combobox_conta_contabil['values'] = []
                resposta = messagebox.askyesno("Cadastrar Conta", "Nenhuma conta encontrada para esta empresa. Deseja cadastrar uma nova conta?")
                if resposta:
                    self.abrir_tela_nova_conta()
                else:
                    self.root.lift() #? deixa a janela importador como foco dps
        else:
            self.combobox_conta_contabil['values'] = []

    def carregar_empresas(self):
        empresas = []
        if os.path.exists('empresas.csv'):
            with open('empresas.csv', mode='r', newline='') as file:
                reader = csv.reader(file)
                for row in reader:
                    if row:
                        empresas.append(f"{row[0]} - {row[1]}")
        return empresas

    def carregar_contas_contabeis(self, codigo_empresa):
        contas = []
        if os.path.exists('contas.csv'):
            with open('contas.csv', mode='r', newline='') as file:
                reader = csv.reader(file)
                next(reader, None)
                for row in reader:
                    if row and re.search(rf',{codigo_empresa}$', ','.join(row)):
                        contas.append(f"{row[0]} - {row[1]} - {row[2]} - {row[3]} - {row[4]} - {row[5]}")
        return contas

    def abrir_tela_nova_empresa(self):
        if not hasattr(self, 'tela_nova_empresa') or not self.tela_nova_empresa.winfo_exists():
            self.tela_nova_empresa = tk.Toplevel(self.root)
            app = TelaNovaEmpresa(self.tela_nova_empresa, self.salvar_nova_empresa)

    def abrir_tela_nova_conta(self):
        if not hasattr(self, 'tela_nova_conta') or not self.tela_nova_conta.winfo_exists():
            self.tela_nova_conta = tk.Toplevel(self.root)
            app = TelaNovaConta(self.tela_nova_conta, self.salvar_nova_conta)

    def salvar_nova_empresa(self, codigo, razao_social):
        with open('empresas.csv', mode='a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([codigo, razao_social])
        messagebox.showinfo("Nova Empresa", "Empresa adicionada com sucesso!")
        self.combobox_empresa['values'] = self.carregar_empresas()

    def salvar_nova_conta(self, codigo_empresa, banco, agencia, conta_bancaria, conta_ativo, conta_passivo):
        with open('contas.csv', mode='a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([codigo_empresa, banco, agencia, conta_bancaria, conta_ativo, conta_passivo])
        messagebox.showinfo("Nova Conta", "Conta adicionada com sucesso!")
        self.combobox_conta_contabil['values'] = self.carregar_contas_contabeis()

    def verificar_empresa(self, event):
        empresa = self.combobox_empresa.get()
        if empresa:
            codigo = empresa.split(" - ")[0]
            if not any(codigo in e for e in self.empresas):
                self.perguntar_cadastrar_empresa(codigo)
            else:
                self.combobox_conta_contabil.focus_set()

    def verificar_empresa_digitada(self, event):
        empresa = self.combobox_empresa.get()
        if empresa:
            codigo = empresa.split(" - ")[0]
            if not any(codigo in e for e in self.empresas):
                self.perguntar_cadastrar_empresa(codigo)

    def perguntar_cadastrar_empresa(self, codigo):
        resposta = messagebox.askyesno("Cadastrar Empresa", f"A empresa com código {codigo} não está cadastrada. Deseja cadastrá-la?")
        if resposta:
            self.abrir_tela_nova_empresa()

    def confirmar(self):
        empresa = self.combobox_empresa.get()
        conta = self.combobox_conta_contabil.get()
        if empresa and conta:
            arquivo = None
            respostas_safra = {}  #? dicionário para armazenar as respostas
            
            #? verificar se é conta Safra
            if "Safra" in conta:
                resp1 = messagebox.askyesno("Banco Safra", "O extrato está no formato PDF?")
                if resp1:
                    messagebox.showinfo("Aviso", "Funcionalidade de importação de PDF ainda não implementada.")
                else:
                    resp2 = messagebox.askyesno("Banco Safra", "O extrato está no novo formato do Banco Safra?")
                    if resp2:
                        resp3 = messagebox.askyesno("Banco Safra", "O extrato é conta vinculada?")
                        arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls;*.xlsx")])
                        #? armazenar as respostas para usar na acao_safra
                        respostas_safra = {
                            "novo_formato": True,
                            "conta_vinculada": resp3
                        }
                    else:
                        arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls;*.xlsx")])
                        respostas_safra = {
                            "novo_formato": False,
                            "conta_vinculada": False
                        }
            else:
                pdf_resposta = messagebox.askyesno("Formato do Extrato", "O extrato está em PDF?")
                if pdf_resposta:
                    messagebox.showinfo("Aviso", "Funcionalidade de importação de PDF ainda não implementada.")
                else:
                    xlsx_resposta = messagebox.askyesno("Formato do Extrato", "O extrato está em XLS ou XLSX?")
                    if xlsx_resposta:
                        arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls;*.xlsx")])
                    else:
                        messagebox.showinfo("Aviso", "Por favor, selecione um arquivo válido.")

            if arquivo:
                #? passar as respostas junto com os outros parâmetros
                self.callback(empresa, conta, arquivo, respostas_safra)

            self.root.destroy()
        else:
            messagebox.showwarning("Aviso", "Por favor, preencha ambos os campos.")

    def cancelar(self):
        self.root.destroy()

    def alterar_conta(self):
        messagebox.showinfo("Alterar Conta", "Funcionalidade de Alterar Conta ainda não implementada.")
        self.root.lift()
        
class TelaNovaConta:
    def __init__(self, root, callback):
        self.root = root
        self.callback = callback
        self.root.title("Nova Conta")
        self.root.geometry("480x240")
        self.root.config(bg='#f4f4f4')
        self.root.iconbitmap(r"C:\Users\regina.santos\Desktop\Automacao\Judite\icon.ico")

        #* -------------------- TÍTULO PRINCIPAL -------------------- #
        self.title_label = tk.Label(root, text="Informações da conta bancária", font=("Roboto", 17, "bold"), bg='#f4f4f4') 
        self.title_label.grid(row=1, column=0, columnspan=1, padx=10, pady=10, sticky="w")

        #* -------------------- CAMPOS DE INFORMAÇÕES -------------------- #
        self.label_codigo_empresa = tk.Label(root, text="Código da Empresa:", font=("Roboto", 10), bg='#f4f4f4') 
        self.label_codigo_empresa.grid(row=2, column=0, pady=1, padx=10, sticky="w")
        self.entry_codigo_empresa = tk.Entry(root, font=("Roboto", 10), width=50)
        self.entry_codigo_empresa.grid(row=3, column=0, pady=1, padx=10, sticky="w")

        self.label_bancos = tk.Label(root, text="Banco:", font=("Roboto", 10), bg='#f4f4f4')
        self.label_bancos.grid(row=4, column=0, pady=1, padx=10, sticky="w")
        self.bancos = ["001 - Banco do Brasil", "077 - Banco Inter", "104 - Banco Caixa Eletrônica", 
                       "237 - Banco Bradesco", "274 - Banco Grafeno", "290 - Banco Pagseguro", 
                       "336 - Banco C6 Bank", "341 - Banco Itau", "353 - Banco Santander", 
                       "399 - Banco HSBC", "422 - Banco Safra", "505 - Credit Suisse", 
                       "707 - Banco Daycoval", "748 - Banco Sicredi"]
        self.combobox_banco = ttk.Combobox(root, values=self.bancos, font=("Roboto", 10), width=62)
        self.combobox_banco.grid(row=5, column=0, pady=1, padx=10, sticky="w")

        self.label_agencia = tk.Label(root, text="Agência:", font=("Roboto", 10), bg='#f4f4f4')
        self.label_agencia.grid(row=6, column=0, pady=1, padx=10, sticky="w")
        self.entry_agencia = tk.Entry(root, font=("Roboto", 10), width=13)
        self.entry_agencia.grid(row=7, column=0, pady=1, padx=10, sticky="w")

        self.label_conta_bancaria = tk.Label(root, text="Conta Bancária:", font=("Roboto", 10), bg='#f4f4f4')
        self.label_conta_bancaria.grid(row=6, column=0, pady=1, padx=116, sticky="w")
        self.entry_conta_bancaria = tk.Entry(root, font=("Roboto", 10), width=15)
        self.entry_conta_bancaria.grid(row=7, column=0, pady=1, padx=116, sticky="w")

        self.label_conta_ativo = tk.Label(root, text="N° Conta Ativo:", font=("Roboto", 10), bg='#f4f4f4')
        self.label_conta_ativo.grid(row=6, column=0, pady=1, padx=240, sticky="w")
        self.entry_conta_ativo = tk.Entry(root, font=("Roboto", 10), width=15)
        self.entry_conta_ativo.grid(row=7, column=0, pady=1, padx=240, sticky="w")

        self.label_conta_passivo = tk.Label(root, text="N° Conta Passivo:", font=("Roboto", 10), bg='#f4f4f4')
        self.label_conta_passivo.grid(row=6, column=0, pady=1, padx=120, sticky="e")
        self.entry_conta_passivo = tk.Entry(root, font=("Roboto", 10), width=15)
        self.entry_conta_passivo.grid(row=7, column=0, pady=1, padx=120, sticky="e")

        #* -------------------- BOTÕES PRINCIPAIS -------------------- #
        self.btn_salvar = tk.Button(root, text="Salvar", command=self.salvar, font=("Roboto", 10), bg="#4CAF50", fg="white")
        self.btn_salvar.grid(row=8, column=0, sticky="e", padx=192, pady=8)

        self.btn_cancelar = tk.Button(root, text="Cancelar", command=self.cancelar, font=("Roboto", 10), bg="#FF5733", fg="white")
        self.btn_cancelar.grid(row=8, column=0, sticky="e", padx=122, pady=8)

    def salvar(self):
        codigo_empresa = self.entry_codigo_empresa.get()
        bancos = self.combobox_banco.get()
        agencia = self.entry_agencia.get()
        conta_bancaria = self.entry_conta_bancaria.get()
        conta_ativo = self.entry_conta_ativo.get()
        conta_passivo = self.entry_conta_passivo.get()
        if codigo_empresa and bancos and agencia and conta_bancaria and conta_ativo and conta_passivo:
            self.root.destroy()
            self.callback(conta_ativo, bancos, agencia, conta_bancaria, conta_passivo, codigo_empresa)
        else:
            messagebox.showwarning("Aviso", "Por favor, preencha todos os campos.")

    def cancelar(self):
        self.root.destroy()

class TelaNovaEmpresa:
    def __init__(self, root, callback):
        self.root = root
        self.callback = callback
        self.root.title("Nova Empresa")
        self.root.geometry("480x210")
        self.root.config(bg='#f4f4f4')
        self.root.iconbitmap(r"C:\Users\regina.santos\Desktop\Automacao\Judite\icon.ico")

        #* -------------------- TÍTULO PRINCIPAL -------------------- #
        self.title_label = tk.Label(root, text="Informe os dados da empresa", font=("Roboto", 17, "bold"), bg='#f4f4f4') 
        self.title_label.grid(row=1, column=0, columnspan=1, padx=10, pady=10, sticky="w")

        #* -------------------- CAMPOS DE INFORMAÇÕES -------------------- #
        self.label_codigo = tk.Label(root, text="Código:", font=("Roboto", 10), bg='#f4f4f4') 
        self.label_codigo.grid(row=2, column=0, pady=1, padx=10, sticky="w")
        self.entry_codigo = tk.Entry(root, font=("Roboto", 10), width=65) 
        self.entry_codigo.grid(row=3, column=0, pady=1, padx=10, sticky="w")

        self.label_razao_social = tk.Label(root, text="Razão Social:", font=("Roboto", 10), bg='#f4f4f4') 
        self.label_razao_social.grid(row=4, column=0, pady=1, padx=10, sticky="w")
        self.entry_razao_social = tk.Entry(root, font=("Roboto", 10), width=65) 
        self.entry_razao_social.grid(row=5, column=0, pady=1, padx=10, sticky="w")

        #* -------------------- BOTÕES PRINCIPAIS -------------------- #
        self.btn_salvar = tk.Button(root, text="Salvar", command=self.salvar, font=("Roboto", 10), bg="#4CAF50", fg="white")
        self.btn_salvar.grid(row=7, column=0, columnspan=2, pady=15, padx=77, sticky="e")

        self.btn_cancelar = tk.Button(root, text="Cancelar", command=self.cancelar, font=("Roboto", 10), bg="#FF5733", fg="white")
        self.btn_cancelar.grid(row=7, column=0, columnspan=2, pady=15, padx=10, sticky="e")

    def salvar(self):
        codigo = self.entry_codigo.get()
        razao_social = self.entry_razao_social.get()
        if codigo and razao_social:
            self.root.destroy()
            self.callback(codigo, razao_social)
        else:
            messagebox.showwarning("Aviso", "Por favor, preencha ambos os campos.")

    def cancelar(self):
        self.root.destroy()

if __name__ == "__main__":
    open_main_window()