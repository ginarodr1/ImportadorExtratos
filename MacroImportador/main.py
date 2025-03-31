import tkinter as tk
from autenticacao.senha_login import SenhaLogin
from importacao.importador_extratos import ImportadorExtratos

def open_main_window():
    root = tk.Tk()
    app = ImportadorExtratos(root)
    root.mainloop()

if __name__ == "__main__":
    login_root = tk.Tk()
    login_app = SenhaLogin(login_root, open_main_window)
    login_root.mainloop()
