import tkinter as tk
from tkinter import messagebox

class TelaNovaEmpresa:
    def __init__(self, root, callback):
        self.root = root
        self.callback = callback
        self.root.title("Nova Empresa")
        self.root.geometry("480x210")
        self.root.config(bg='#f4f4f4')
        self.root.iconbitmap(r"C:\Users\regina.santos\Desktop\Automação\Judite\icon.ico")

        self.title_label = tk.Label(root, text="Informe os dados da empresa", font=("Roboto", 17, "bold"), bg='#f4f4f4')
        self.title_label.grid(row=1, column=0, columnspan=1, padx=10, pady=10, sticky="w")

        self.label_codigo = tk.Label(root, text="Código:", font=("Roboto", 10), bg='#f4f4f4')
        self.label_codigo.grid(row=2, column=0, pady=1, padx=10, sticky="w")
        self.entry_codigo = tk.Entry(root, font=("Roboto", 10), width=65)
        self.entry_codigo.grid(row=3, column=0, pady=1, padx=10, sticky="w")

        self.label_razao_social = tk.Label(root, text="Razão Social:", font=("Roboto", 10), bg='#f4f4f4')
        self.label_razao_social.grid(row=4, column=0, pady=1, padx=10, sticky="w")
        self.entry_razao_social = tk.Entry(root, font=("Roboto", 10), width=65)
        self.entry_razao_social.grid(row=5, column=0, pady=1, padx=10, sticky="w")

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
