import tkinter as tk
from tkinter import ttk

class SimpleApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # Configurações da janela
        self.title("Aplicação Simples")
        self.geometry("400x200")
        self.config(bg='#F0F0F0')

        # Configuração de estilo para o botão
        style = ttk.Style()

        # Configura o estilo TButton diretamente
        style.configure("TButton",
                        font=("Roboto", 12, "bold"),
                        padding=2,
                        relief="flat",
                        background="#292935",
                        foreground="white",
                        borderwidth=0,
                        focusthickness=0,
                        highlightthickness=0)

        style.map("TButton",
                  background=[('active', '#45a049')],
                  foreground=[('active', 'white')])

        # Botão "Importar" usando ttk
        self.btn_importar = ttk.Button(self, text="Importar", command=self.on_import_click, style="TButton")
        self.btn_importar.pack(pady=50, padx=20, fill='x')

    def on_import_click(self):
        # Lógica para quando o botão for clicado
        print("Botão Importar clicado!")

if __name__ == "__main__":
    app = SimpleApp()
    app.mainloop()
