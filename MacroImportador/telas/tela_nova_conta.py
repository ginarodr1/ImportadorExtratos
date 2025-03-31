import tkinter as tk
from tkinter import messagebox, ttk

class TelaNovaConta:
    def __init__(self, root, callback):
        self.root = root
        self.callback = callback
        self.root.title("Nova Conta")
        self.root.geometry("480x240")
        self.root.config(bg='#f4f4f4')
        self.root.iconbitmap(r"C:\Users\regina.santos\Desktop\Automação\Judite\icon.ico")

        self.title_label = tk.Label(root, text="Informações da conta bancária", font=("Roboto", 17, "bold"), bg='#f4f4f4')
        self.title_label.grid(row=1, column=0, columnspan=1, padx=10, pady=10, sticky="w")

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
