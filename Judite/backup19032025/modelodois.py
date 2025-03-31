import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import csv
import re
from openpyxl import Workbook

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # Configurações da janela
        self.title("Importador de Extratos Bancários")
        self.geometry("1000x600")
        self.config(bg='#F0F0F0')
        self.state("zoomed")
        self.iconbitmap(r"C:\Users\regina.santos\Desktop\Automação\Judite\icon.ico")

        # Container principal
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        # Dicionário para armazenar frames
        self.frames = {}

        # Criação das páginas
        for F in (ImportPage, ClassifPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("ImportPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

class ImportPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.config(style="TFrame")

        # Instância do ImportadorExtratos
        self.importador = ImportadorExtratos(self)
        self.importador.pack(fill="both", expand=True)

        # Botão para navegar para a página de classificação
        switch_button = ttk.Button(self, text="Go to Classif Page",
                                    command=lambda: controller.show_frame("ClassifPage"))
        switch_button.pack(pady=20)

class ClassifPage(ttk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.config(style="TFrame")

        label = ttk.Label(self, text="This is the Classif Page", background='#E0E0E0')
        label.pack(pady=20)

        switch_button = ttk.Button(self, text="Go to Import Page",
                                    command=lambda: controller.show_frame("ImportPage"))
        switch_button.pack()

class ImportadorExtratos(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.config(style="TFrame")

        # Carregar dados do banco de dados
        csv_file_path = r"C:\Users\regina.santos\Desktop\Automação\Judite\bancodedadosrv01.csv"
        self.df_banco_dados = pd.read_csv(csv_file_path, delimiter=';')

        # Configuração da interface
        self.setup_ui()

    def setup_ui(self):
        # Título principal
        self.title_label = tk.Label(self, text="Importador de Extratos", font=("Roboto", 17, "bold"), bg='#F0F0F0')
        self.title_label.grid(row=2, column=0, padx=10, sticky="w")

        # Botões principais
        self.btn_importar = ttk.Button(self, text="Importar", command=self.mostrar_selecao_conta, style="AccentButton")
        self.btn_importar.grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.btn_limpar = tk.Button(self, text="Limpar Tudo", command=self.confirmar_limpar_dados, font=("Roboto", 10), bg="#FF5733", fg="white", relief="groove", padx=5, pady=3)
        self.btn_limpar.grid(row=3, column=0, padx=80, sticky="w")

        self.btn_classificar1p = tk.Button(self, text="Classificar 1P", command=self.classificar_dados, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_classificar1p.grid(row=0, column=3, padx=90, sticky="e")

        self.btn_limpar1p = tk.Button(self, text="Limpar 1P", command=self.limpar_1p, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_limpar1p.grid(row=0, column=3, padx=5, sticky="e")

        self.btn_exportar = tk.Button(self, text="Exportar", command=self.exportar_dados, font=("Roboto", 10), bg="#4CAF50", fg="white", relief="groove", padx=5, pady=3)
        self.btn_exportar.grid(row=0, column=6, sticky="w")

        # Campos de saldo
        self.setup_saldo_fields()

        # Treeview para exibir os dados importados
        self.setup_treeview()

        # Configuração de layout
        self.configure_grid()

    def setup_saldo_fields(self):
        # Campos de saldo
        self.saldo_inicial_label = tk.Label(self, text="Saldo Inicial Importado>", font=("Roboto", 11), bg='#F0F0F0', anchor='e')
        self.saldo_inicial_label.grid(row=0, column=1, padx=1, sticky="w")
        self.saldo_inicial_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_inicial_entry.grid(row=0, column=1, columnspan=2, padx=160, sticky="w")

        self.saldo_final_label = tk.Label(self, text="Saldo Final Importado>", font=("Roboto", 11), bg='#F0F0F0', anchor='e')
        self.saldo_final_label.grid(row=2, column=1, padx=3, sticky="w")
        self.saldo_final_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_final_entry.grid(row=2, column=1, columnspan=2, padx=160, sticky="w")

        self.saldo_final_calculado_label = tk.Label(self, text="Saldo Final Calculado>", font=("Roboto", 11), bg='#F0F0F0', anchor='e')
        self.saldo_final_calculado_label.grid(row=3, column=1, padx=3, sticky="w")
        self.saldo_final_calculado_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.saldo_final_calculado_entry.grid(row=3, column=1, columnspan=2, padx=160, sticky="w")

        self.diferenca_label = tk.Label(self, text="Diferença>", font=("Roboto", 11), bg='#F0F0F0', anchor='e')
        self.diferenca_label.grid(row=4, column=1, padx=82, sticky="w")
        self.diferenca_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=17)
        self.diferenca_entry.grid(row=4, column=1, columnspan=2, padx=160, sticky="w")

        # Outros campos
        self.setup_other_fields()

    def setup_other_fields(self):
        self.empresa_label = tk.Label(self, text="Empresa>", font=("Roboto", 11), bg='#F0F0F0', anchor='w')
        self.empresa_label.grid(row=2, column=2, padx=51, sticky="w")
        self.empresa_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.empresa_entry.grid(row=2, column=2, padx=125, sticky="w")

        self.conta_label = tk.Label(self, text="Conta>", font=("Roboto", 11), bg='#F0F0F0', anchor='w')
        self.conta_label.grid(row=3, column=2, padx=72, sticky="w")
        self.conta_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.conta_entry.grid(row=3, column=2, padx=125, sticky="w")

        self.centro_custo_label = tk.Label(self, text="C/Custo>", font=("Roboto", 11), bg='#F0F0F0', anchor='w')
        self.centro_custo_label.grid(row=4, column=2, padx=54, sticky="w")
        self.centro_custo_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.centro_custo_entry.grid(row=4, column=2, padx=125, sticky="w")

        self.banco_label = tk.Label(self, text="Banco>", font=("Roboto", 11), bg='#F0F0F0', anchor='w')
        self.banco_label.grid(row=3, column=2, padx=1, sticky="e")
        self.banco_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=20)
        self.banco_entry.grid(row=3, column=3, padx=1, sticky="w")

        self.agencia_conta_label = tk.Label(self, text="Agência/Conta>", font=("Roboto", 11), bg='#F0F0F0', anchor='w')
        self.agencia_conta_label.grid(row=4, column=2, padx=1, sticky="e")
        self.agencia_conta_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=20)
        self.agencia_conta_entry.grid(row=4, column=3, padx=1, sticky="w")

        self.saldo_final_contabil_label = tk.Label(self, text="Saldo Final Contábil>", font=("Roboto", 11), bg='#F0F0F0', anchor='w')
        self.saldo_final_contabil_label.grid(row=3, column=3, columnspan=4, padx=95, sticky="e")
        self.saldo_final_contabil_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.saldo_final_contabil_entry.grid(row=3, column=3, columnspan=4, padx=10, sticky="e")

        self.diferenca_extrato_bancario_label = tk.Label(self, text="Diferença c/Extrato Bancário>", font=("Roboto", 11), bg='#F0F0F0', anchor='w')
        self.diferenca_extrato_bancario_label.grid(row=4, column=3, columnspan=4, padx=95, sticky="e")
        self.diferenca_extrato_bancario_entry = tk.Entry(self, font=("Roboto", 10), bd=1, relief="solid", width=10)
        self.diferenca_extrato_bancario_entry.grid(row=4, column=3, columnspan=4, padx=10, sticky="e")

    def setup_treeview(self):
        # Treeview para exibir os dados importados
        self.tree = ttk.Treeview(self, columns=(
            "DataLEB", "DescriçãoLEB", "NDocLEB", "ValorLEB", "SaldoLEB",
            "LançamentoLC", "DataLC", "DébitoLC", "D-C/CLC", "CréditoLC",
            "C-C/CLC", "CNPJLC", "HistóricoLC", "ValorLC"
        ), show="headings")

        # Definir cabeçalhos das colunas
        self.tree.heading("DataLEB", text="Data")
        self.tree.heading("DescriçãoLEB", text="Descrição")
        self.tree.heading("NDocLEB", text="N° Doc")
        self.tree.heading("ValorLEB", text="Valor")
        self.tree.heading("SaldoLEB", text="Saldo")
        self.tree.heading("LançamentoLC", text="Lançamento")
        self.tree.heading("DataLC", text="Data")
        self.tree.heading("DébitoLC", text="Débito")
        self.tree.heading("D-C/CLC", text="D-C/C")
        self.tree.heading("CréditoLC", text="Crédito")
        self.tree.heading("C-C/CLC", text="C-C/C")
        self.tree.heading("CNPJLC", text="CNPJ")
        self.tree.heading("HistóricoLC", text="Histórico")
        self.tree.heading("ValorLC", text="Valor")

        # Ajustar o tamanho das colunas
        self.tree.column("DataLEB", width=60)
        self.tree.column("DescriçãoLEB", width=150)
        self.tree.column("NDocLEB", width=60)
        self.tree.column("ValorLEB", width=60)
        self.tree.column("SaldoLEB", width=60)
        self.tree.column("LançamentoLC", width=150)
        self.tree.column("DataLC", width=60)
        self.tree.column("DébitoLC", width=60)
        self.tree.column("D-C/CLC", width=50)
        self.tree.column("CréditoLC", width=60)
        self.tree.column("C-C/CLC", width=50)
        self.tree.column("CNPJLC", width=80)
        self.tree.column("HistóricoLC", width=100)
        self.tree.column("ValorLC", width=80)
        self.tree.grid(row=6, column=0, columnspan=7, pady=5, padx=10, sticky="nsew")

    def configure_grid(self):
        # Configuração de layout
        for i in range(7):
            self.grid_columnconfigure(i, weight=1)
        for i in range(7):
            self.grid_rowconfigure(i, weight=1)

    def mostrar_selecao_conta(self):
        # Implementação da lógica para mostrar a seleção de conta
        pass

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

    def classificar_dados(self):
        self.copiar_coluna("DataLEB", "DataLC")
        self.copiar_coluna("ValorLEB", "ValorLC")

        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            descricao = values[1]

            correspondencia = self.df_banco_dados[self.df_banco_dados['Descricao'] == descricao]

            if not correspondencia.empty:
                descricao_banco = correspondencia.iloc[0, 5]
                debito = correspondencia.iloc[0, 6]
                credito = correspondencia.iloc[0, 9]

                novos_valores = list(values)
                novos_valores[5] = descricao_banco
                novos_valores[7] = debito
                novos_valores[9] = credito

                self.tree.item(item, values=novos_valores)

        messagebox.showinfo("Classificar Dados", "Dados classificados com sucesso!")

    def copiar_coluna(self, coluna_origem, coluna_destino):
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            valor_origem = values[self.tree["columns"].index(coluna_origem)]
            self.tree.set(item, coluna_destino, valor_origem)

    def exportar_dados(self):
        wb = Workbook()
        ws = wb.active

        colunas_exportar = ["LançamentoLC", "DataLC", "DébitoLC", "D-C/CLC", "CréditoLC",
                            "C-C/CLC", "CNPJLC", "HistóricoLC", "ValorLC"]

        dados = []
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            dados.append([values[self.tree["columns"].index(col)] for col in colunas_exportar])

        df = pd.DataFrame(dados, columns=colunas_exportar)

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Exportar", "Dados exportados com sucesso!")

    def limpar_1p(self):
        messagebox.showinfo("Limpar 1P", "Essa função não está disponível no momento!")

if __name__ == "__main__":
    app = App()
    app.mainloop()
