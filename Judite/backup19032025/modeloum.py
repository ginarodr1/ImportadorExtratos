import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl import Workbook
import pandas as pd
import csv
import os
import bcrypt
import re

class Paginas:
    def __init__(self, root):
        self.root = root
        self.root.title("Importador de Extratos Bancários")
        self.root.geometry("1000x600")
        self.root.config(bg='#D9D9D9')
        self.root.state("zoomed")
        self.root.iconbitmap(r"C:\Users\regina.santos\Desktop\Automação\Judite\icon.ico")
        csv_file_path = r"C:\Users\regina.santos\Desktop\Automação\Judite\bancodedadosrv01.csv"
        self.df_banco_dados = pd.read_csv(csv_file_path, delimiter=';')

        # Cria um widget Notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # Cria as duas abas usando classes internas
        self.tab1 = self.ImportadorPG(self.notebook)
        self.tab2 = self.ClassificadorPG(self.notebook)

        # Adiciona as abas ao Notebook
        self.notebook.add(self.tab1, text="  Importador  ")
        self.notebook.add(self.tab2, text="  Classificador  ")

        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)



    class ImportadorPG(tk.Frame):
        def __init__(self, parent):
            super().__init__(parent)

            # Configura o grid para 5 colunas
            for i in range(5):
                self.grid_columnconfigure(i, weight=1)
            for i in range(6):
                self.grid_rowconfigure(i, weight=1)

            # Instancia a classe ImportadorExtratos e adiciona à aba 1
            self.importador = ImportadorExtratos(self)
            self.importador.grid(row=0, column=0, pady=20, sticky="nsew", columnspan=5)



    class ClassificadorPG(tk.Frame):
        def __init__(self, parent):
            super().__init__(parent)

            # Configura o grid para 5 colunas
            for i in range(5):
                self.grid_columnconfigure(i, weight=1)
            for i in range(6):
                self.grid_rowconfigure(i, weight=1)

            # Instancia a classe Classificador e adiciona à aba 2
            self.classificador = ClassificadorBanco(self)
            self.classificador.grid(row=0, column=0, pady=20, sticky="nsew", columnspan=5)









class ImportadorExtratos(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.config(bg='#D9D9D9')

        # Simulação de carregamento de dados
        csv_file_path = r"C:\Users\regina.santos\Desktop\Automação\Judite\bancodedadosrv01.csv"
        self.df_banco_dados = pd.read_csv(csv_file_path, delimiter=';')

        for i in range(5):
            self.grid_columnconfigure(i, weight=1)
        for i in range(6):
            self.grid_rowconfigure(i, weight=1)

        self.janelas_filhas = []  # Armazenar todas as janelas filhas

        #! -------------------- TÍTULO PRINCIPAL -------------------- #

        self.title_label = tk.Label(self, text="Importador de Extratos", font=("Roboto", 17, "bold"), bg='#D9D9D9')
        self.title_label.grid(row=1, column=0, padx=10, sticky="w")

        #! -------------------- BOTÕES PRINCIPAIS -------------------- #

        self.btn_importar = tk.Button(self, text="Importar", command=self.mostrar_selecao_conta, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_importar.grid(row=2, column=0, padx=10, sticky="w")

        self.btn_limpar = tk.Button(self, text="Limpar Tudo", command=self.confirmar_limpar_dados, font=("Roboto", 10), bg="#FF5733", fg="white", relief="groove", padx=5, pady=3)
        self.btn_limpar.grid(row=2, column=0, padx=80, sticky="w")

        self.btn_classificar1p = tk.Button(self, text="Classificar 1P", command=self.classificar_dados, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_classificar1p.grid(row=0, column=3, padx=90, sticky="e")

        self.btn_limpar1p = tk.Button(self, text="Limpar 1P", command=self.limpar_1p, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_limpar1p.grid(row=0, column=3, padx=5, sticky="e")

        self.btn_exportar = tk.Button(self, text="Exportar", command=self.exportar_dados, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_exportar.grid(row=0, column=4, sticky="w")

        #! -------------------- CAMPOS DE SALDO -------------------- #
        #* -------------- LANÇAMENTO EXTRATO BANCÁRIO -------------- #

        self.saldo_inicial_label = tk.Label(self, text="Saldo Inicial Importado>", font=("Roboto", 11), bg='#D9D9D9', anchor='e')
        self.saldo_inicial_label.grid(row=0, column=1, padx=1, sticky="w")
        self.saldo_inicial_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_inicial_entry.grid(row=0, column=1, columnspan=2, padx=160, sticky="w")

        self.saldo_final_label = tk.Label(self, text="Saldo Final Importado>", font=("Roboto", 11), bg='#D9D9D9', anchor='e')
        self.saldo_final_label.grid(row=1, column=1, padx=3, sticky="w")
        self.saldo_final_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_final_entry.grid(row=1, column=1, columnspan=2, padx=160, sticky="w")

        self.saldo_final_calculado_label = tk.Label(self, text="Saldo Final Calculado>", font=("Roboto", 11), bg='#D9D9D9', anchor='e')
        self.saldo_final_calculado_label.grid(row=2, column=1, padx=3, sticky="w")
        self.saldo_final_calculado_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_final_calculado_entry.grid(row=2, column=1, columnspan=2, padx=160, sticky="w")

        self.diferenca_label = tk.Label(self, text="Diferença>", font=("Roboto", 11), bg='#D9D9D9', anchor='e')
        self.diferenca_label.grid(row=3, column=1, padx=82, sticky="w")
        self.diferenca_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.diferenca_entry.grid(row=3, column=1, columnspan=2, padx=160, sticky="w")


        self.empresa_label = tk.Label(self, text="Empresa>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.empresa_label.grid(row=1, column=2, padx=51, sticky="w")
        self.empresa_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.empresa_entry.grid(row=1, column=2, padx=125, sticky="w")

        self.conta_label = tk.Label(self, text="Conta>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.conta_label.grid(row=2, column=2, padx=72, sticky="w")
        self.conta_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.conta_entry.grid(row=2, column=2, padx=125, sticky="w")

        self.centro_custo_label = tk.Label(self, text="C/Custo>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.centro_custo_label.grid(row=3, column=2, padx=54, sticky="w")
        self.centro_custo_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.centro_custo_entry.grid(row=3, column=2, padx=125, sticky="w")


        self.banco_label = tk.Label(self, text="Banco>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.banco_label.grid(row=2, column=2, padx=1, sticky="e")
        self.banco_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=20)
        self.banco_entry.grid(row=2, column=3, padx=1, sticky="w")
        
        self.agencia_conta_label = tk.Label(self, text="Agência/Conta>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.agencia_conta_label.grid(row=3, column=2, padx=1, sticky="e")
        self.agencia_conta_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=20)
        self.agencia_conta_entry.grid(row=3, column=3, padx=1, sticky="w")


        self.saldo_final_contabil_label = tk.Label(self, text="Saldo Final Contábil>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.saldo_final_contabil_label.grid(row=2, column=3, columnspan=4, padx=95, sticky="e")
        self.saldo_final_contabil_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.saldo_final_contabil_entry.grid(row=2, column=3, columnspan=4, padx=10, sticky="e")

        self.diferenca_extrato_bancario_label = tk.Label(self, text="Diferença c/Extrato Bancário>", font=("Roboto", 11), bg='#D9D9D9', anchor='w')
        self.diferenca_extrato_bancario_label.grid(row=3, column=3, columnspan=4, padx=95, sticky="e")
        self.diferenca_extrato_bancario_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.diferenca_extrato_bancario_entry.grid(row=3, column=3, columnspan=4, padx=10, sticky="e")

        #self.linhadeajuda_label = tk.Label(self, text="Coluna 0", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        #self.linhadeajuda_label.grid(row=1, column=0)
        #self.linhadeajuda_label = tk.Label(self, text="Coluna 1", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        #self.linhadeajuda_label.grid(row=1, column=1)
        #self.linhadeajuda_label = tk.Label(self, text="Coluna 2", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        #self.linhadeajuda_label.grid(row=1, column=2)
        #self.linhadeajuda_label = tk.Label(self, text="Coluna 3", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        #self.linhadeajuda_label.grid(row=1, column=3)
        #self.linhadeajuda_label = tk.Label(self, text="Coluna 4", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        #self.linhadeajuda_label.grid(row=1, column=4)
        #self.linhadeajuda_label = tk.Label(self, text="Coluna 5", fg='#D9D9D9', font=("Roboto", 10), bg='#D9D9D9')
        #self.linhadeajuda_label.grid(row=1, column=5)

    def mostrar_selecao_conta(self):
        print("Teste!")

    def confirmar_limpar_dados(self):
        print("Teste!")

    def classificar_dados(self):
        print("Teste!")

    def limpar_1p(self):
        print("Teste!")

    def exportar_dados(self):
        print("Teste!")


        










class ClassificadorBanco(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        # Carrega os dados do CSV
        csv_file_path = r"C:\Users\regina.santos\Desktop\Automação\Judite\bancodedadosrv01.csv"
        self.df_banco_dados = pd.read_csv(csv_file_path, delimiter=';')

        self.title_label = tk.Label(self, text="Classificador", font=("Roboto", 17, "bold"), bg='#D9D9D9')
        self.title_label.grid(row=3, column=0, padx=10, sticky="w")

        # Cria a Treeview para exibir os dados
        self.tree = ttk.Treeview(self, columns=list(self.df_banco_dados.columns), show="headings")

        # Define os cabeçalhos das colunas
        for col in self.df_banco_dados.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        # Adiciona os dados à Treeview
        for index, row in self.df_banco_dados.iterrows():
            self.tree.insert("", "end", values=list(row))

        # Adiciona a Treeview ao layout
        self.tree.grid(row=0, column=0, sticky="nsew")

        # Configura o grid para expandir corretamente
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)


if __name__ == "__main__":
    root = tk.Tk()
    app = Paginas(root)
    root.mainloop()
