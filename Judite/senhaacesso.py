import tkinter as tk
from tkinter import ttk, messagebox, filedialog

def open_main_window():
    root = tk.Tk()
    app = ImportadorExtratos(root)
    root.mainloop()

class SenhaLogin:
    def __init__(self, root, callback):
        self.root = root
        self.callback = callback
        self.root.title("Login")
        self.root.geometry("247x100")
        self.root.config(bg='#f4f4f4')
        self.root.iconbitmap(r"C:\Users\regina.santos\Desktop\Automacao\Judite\icon.ico")

        self.titulosenha = tk.Label(root, text="Importador de Extratos está protegido.", font=("Roboto", 10), bg='#f4f4f4')
        self.titulosenha.grid(row=2, column=0, columnspan=2, pady=5, padx=10, sticky="w")

        self.senha_label = tk.Label(root, text="Senha:", font=("Roboto", 10), bg='#f4f4f4')
        self.senha_label.grid(row=3, column=0, columnspan=2, pady=5, padx=10, sticky="w")

        self.senha_entry = tk.Entry(root, show="*", font=("Roboto", 10), width=20)
        self.senha_entry.grid(row=3, column=0, padx=60, sticky="w")

        self.login_button = tk.Button(root, text="Entrar", command=self.verificar_senha, font=("Roboto", 10), bg="#4CAF50", fg="white")
        self.login_button.grid(row=4, column=0, columnspan=1, padx=27, sticky="e")

        self.senha_entry.bind("<Return>", lambda event: self.verificar_senha())

        self.root.after(0, lambda: self.root.focus_force())
        self.senha_entry.focus_set()

        self.senha_hash = b'$2b$12$ZlVoXX9jSkWwL.jkPDVgdOlaN0vSecKPWIl8WoEfiUJLA..1A24WS' #! senha em hash gerada e armazenada

    def verificar_senha(self):
        senha = self.senha_entry.get()
        if bcrypt.checkpw(senha.encode(), self.senha_hash):
            self.root.destroy()
            self.callback()
        else:
            messagebox.showerror("Erro", "Senha incorreta!")
            self.senha_entry.delete(0, tk.END)


if __name__ == "__main__":
    login_root = tk.Tk()
    login_app = SenhaLogin(login_root, open_main_window)
    login_root.mainloop()