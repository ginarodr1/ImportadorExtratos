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
        self.root.config(bg='#F0F0F0')
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
            self.config(bg='#F0F0F0')  # Define a cor de fundo do frame

            # Configura o grid para 5 colunas
            for i in range(5):
                self.grid_columnconfigure(i, weight=1)
            for i in range(6):
                self.grid_rowconfigure(i, weight=1)

            # Instancia a classe ImportadorExtratos e adiciona à aba 1
            self.importador = ImportadorExtratos(self)
            self.importador.grid(row=0, column=0, pady=10, sticky="nsew", columnspan=5)

    class ClassificadorPG(tk.Frame):
        def __init__(self, parent):
            super().__init__(parent)
            self.config(bg='#F0F0F0')  # Define a cor de fundo do frame

            # Configura o grid para 5 colunas
            for i in range(5):
                self.grid_columnconfigure(i, weight=1)
            for i in range(6):
                self.grid_rowconfigure(i, weight=1)

            # Instancia a classe Classificador e adiciona à aba 2
            self.classificador = ClassificadorBanco(self)
            self.classificador.grid(row=0, column=0, pady=10, sticky="nsew", columnspan=5)

class ImportadorExtratos(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.config(bg='#F0F0F0')  # Define a cor de fundo do frame

        # Simulação de carregamento de dados
        csv_file_path = r"C:\Users\regina.santos\Desktop\Automação\Judite\bancodedadosrv01.csv"
        self.df_banco_dados = pd.read_csv(csv_file_path, delimiter=';')

        for i in range(5):
            self.grid_columnconfigure(i, weight=1)
        for i in range(6):
            self.grid_rowconfigure(i, weight=1)

        self.janelas_filhas = []  # Armazenar todas as janelas filhas

        # Campo de saldo inicial na primeira linha
        self.saldo_inicial_label = tk.Label(self, text="Saldo Inicial Importado>", font=("Roboto", 11), bg='#F0F0F0', anchor='e')
        self.saldo_inicial_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")

        self.saldo_inicial_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_inicial_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        # Título principal na segunda linha
        self.title_label = tk.Label(self, text="Importador de Extratos", font=("Roboto", 17, "bold"), bg='#F0F0F0')
        self.title_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        # Botões principais
        self.btn_importar = tk.Button(self, text="Importar", command=self.mostrar_selecao_conta, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_importar.grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.btn_limpar = tk.Button(self, text="Limpar Tudo", command=self.confirmar_limpar_dados, font=("Roboto", 10), bg="#FF5733", fg="white", relief="groove", padx=5, pady=3)
        self.btn_limpar.grid(row=2, column=0, padx=80, pady=5, sticky="w")

    def mostrar_selecao_conta(self):
        print("Teste!")

    def confirmar_limpar_dados(self):
        print("Teste!")

class ClassificadorBanco(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.config(bg='#F0F0F0')  # Define a cor de fundo do frame

        # Carrega os dados do CSV
        csv_file_path = r"C:\Users\regina.santos\Desktop\Automação\Judite\bancodedadosrv01.csv"
        self.df_banco_dados = pd.read_csv(csv_file_path, delimiter=';')

if __name__ == "__main__":
    root = tk.Tk()
    app = Paginas(root)
    root.mainloop()
