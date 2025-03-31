import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("INGRID")

message = tk.Label(root, text="Qual será o mês salvo?")
message.pack()

window_width = 300 #tamanho da janela
window_height = 100

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
center_x = int(screen_width/2 - window_width / 2)
center_y = int(screen_height/2 - window_height / 2)
root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}') #meio da tela
root.attributes('-topmost', 1) #sempre no topo
root.resizable(False, False)
root.iconbitmap('icon.ico') #icone





#meses = [
    #"Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
#]

#mes_selecionado = tk.StringVar()

#combo = ttk.Combobox(root, textvariable=mes_selecionado, values=meses, state="readonly")
#combo.pack(pady=10)
#combo.current(0)





root.mainloop()