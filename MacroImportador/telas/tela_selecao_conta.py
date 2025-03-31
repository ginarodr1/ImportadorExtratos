import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import os
import re

class TelaSelecaoConta:
    def __init__(self, root, callback):
        self.root = root
        self.callback = callback
        self.root.title("Importador")
        self.root.geometry("480x210")
        self.root.config(bg='#f4f4f4')
        self.root.iconbitmap(r"C:\Users\regina.santos\Desktop\Automação\Judite\icon.ico")

        self.title_label = tk.Label(root, text="Qual conta bancária irá importar?", font=("Roboto", 17, "bold"), bg='#f4f4f4')
        self.title_label.grid(row=1, column=0, columnspan=1, padx=10, pady=10, sticky="w")

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
                    self.root.lift()
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
        from telas.tela_nova_empresa import TelaNovaEmpresa
        if not hasattr(self, 'tela_nova_empresa') or not self.tela_nova_empresa.winfo_exists():
            self.tela_nova_empresa = tk.Toplevel(self.root)
            app = TelaNovaEmpresa(self.tela_nova_empresa, self.salvar_nova_empresa)

    def abrir_tela_nova_conta(self):
        from telas.tela_nova_conta import TelaNovaConta
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

    def confirmar(self):
        empresa = self.combobox_empresa.get()
        conta = self.combobox_conta_contabil.get()
        if empresa and conta:
            pdf_resposta = messagebox.askyesno("Formato do Extrato", "O extrato está em PDF?")

            arquivo = None
            if pdf_resposta:
                arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
            else:
                xlsx_resposta = messagebox.askyesno("Formato do Extrato", "O extrato está em XLSX?")

                if xlsx_resposta:
                    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls;*.xlsx")])
                else:
                    messagebox.showinfo("Aviso", "Por favor, selecione um arquivo válido.")

            if arquivo:
                self.callback(empresa, conta, arquivo)

            self.root.destroy()
        else:
            messagebox.showwarning("Aviso", "Por favor, preencha ambos os campos.")

    def cancelar(self):
        self.root.destroy()

    def alterar_conta(self):
        messagebox.showinfo("Alterar Conta", "Funcionalidade de Alterar Conta ainda não implementada.")
        self.root.lift()
