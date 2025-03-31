import pandas as pd
import tkinter as tk
from tkinter import filedialog

def detectar_banco(nome_arquivo):
    bancos = ["Bco.Brasil", "Banco do Brasil", "BANCO DO BRASIL", "BB", "BRASIL", "Brasil"
              "Bco.Inter", "Banco Inter", "BANCO INTER", "Inter", "INTER",
              "Bco.Caixa", "Banco Caixa Eletrônica", "BANCO CAIXA ELETRÔNICA", "Caixa Eletrônica", "Caixa", 
              "Bco.Bradesco", "Banco Bradesco", "BANCO BRADESCO", "Bradesco",
              "Bco.Grafeno", "Banco Grafeno", "BANCO GRAFENO", "GRAFENO", "Grafeno",
              "Bco.Pagseguro", "Banco Pagseguro", "BANCO PAGSEGURO", "PAGSEGURO", "Pagseguro",
              "Bco.C6Bank", "Banco C6 Bank", "C6 Bank", "BANCO C6 BANK", "C6 BANK", "C6BANK", "C6",
              "Bco.Itaú", "Banco Itaú", "BANCO ITAÚ", "ITAÚ", "Itaú",
              "Bco.Santander", "Banco Santander", "BANCO SANTANDER", "SANTANDER", "Santander",
              "Bco.HSBC", "Banco HSBC", "BANCO HSBC", "HSBC",
              "Bco.Safra", "Banco Safra", "BANCO SAFRA", "SAFRA", "Safra",
              "Bco.Suisse", "Banco Suisse", "Banco Credit Suisse", "Credit Suisse", "CREDIT SUISSE", "SUISSE", "BANCO SUISSE",
              "Bco.Daycoval", "Banco Daycoval", "BANCO DAYCOVAL", "DAYCOVAL", "Daycoval",
              "Bco.Itaú", "Banco Itaú", "BANCO ITAÚ", "ITAÚ", "Itaú", "Itau", "ITAU",
              ]
    
    for banco in bancos:
        if banco in nome_arquivo:
            print(f"Banco detectado: {banco}")
            executar_acao_para_banco(banco)
            break
    else:
        print("Nenhum banco detectado no nome do arquivo.")

def executar_acao_para_banco(nome_banco):
    def acao_brasil():
        print("Executando ação específica para o Banco do Brasil.")

    def acao_inter():
        print("Executando ação específica para o Inter.")

    def acao_caixa():
        print("Executando ação específica para a Caixa Eletrônica.")

    def acao_bradesco():
        print("Executando ação específica para o Bradesco.")

    def acao_grafeno():
        print("Executando ação específica para o Grafeno.")

    def acao_pagseguro():
        print("Executando ação específica para o Pagseguro.")

    def acao_c6bank():
        print("Executando ação específica para o C6 Bank.")

    def acao_itau():
        print("Executando ação específica para o Itaú.")

    def acao_santander():
        print("Executando ação específica para o Santander.")

    def acao_hsbc():
        print("Executando ação específica para o HSBC.")

    def acao_safra():
        print("Executando ação específica para o Safra.")

    def acao_suisse():
        print("Executando ação específica para o Credit Suisse.")

    def acao_daycoval():
        print("Executando ação específica para o Daycoval.")

    def acao_sicredi():
        print("Executando ação específica para o Sicredi.")

    acoes = {
        "Bco.Brasil": acao_brasil, "Banco do Brasil": acao_brasil, "BANCO DO BRASIL": acao_brasil, "BB": acao_brasil, "BRASIL": acao_brasil, "Brasil": acao_brasil,
        "Bco.Inter": acao_inter, "Banco Inter": acao_inter, "BANCO INTER": acao_inter, "Inter": acao_inter, "INTER": acao_inter,
        "Bco.Caixa": acao_caixa, "Banco Caixa Eletrônica": acao_caixa, "BANCO CAIXA ELETRÔNICA": acao_caixa, "Caixa Eletrônica": acao_caixa, "Caixa": acao_caixa, 
        "Bco.Bradesco": acao_bradesco, "Banco Bradesco": acao_bradesco, "BANCO BRADESCO": acao_bradesco, "Bradesco": acao_bradesco,
        "Bco.Grafeno": acao_grafeno, "Banco Grafeno": acao_grafeno, "BANCO GRAFENO": acao_grafeno, "GRAFENO": acao_grafeno, "Grafeno": acao_grafeno,
        "Bco.Pagseguro": acao_pagseguro, "Banco Pagseguro": acao_pagseguro, "BANCO PAGSEGURO": acao_pagseguro, "PAGSEGURO": acao_pagseguro, "Pagseguro": acao_pagseguro,
        "Bco.C6Bank": acao_c6bank, "Banco C6 Bank": acao_c6bank, "C6 Bank": acao_c6bank, "BANCO C6 BANK": acao_c6bank, "C6 BANK": acao_c6bank, "C6BANK": acao_c6bank, "C6": acao_c6bank, "c6": acao_c6bank,
        "Bco.Itaú": acao_itau, "Banco Itaú": acao_itau, "BANCO ITAÚ": acao_itau, "ITAÚ": acao_itau, "Itaú": acao_itau,
        "Bco.Santander": acao_santander, "Banco Santander": acao_santander, "BANCO SANTANDER": acao_santander, "SANTANDER": acao_santander, "Santander": acao_santander,
        "Bco.HSBC": acao_hsbc, "Banco HSBC": acao_hsbc, "BANCO HSBC": acao_hsbc, "HSBC": acao_hsbc,
        "Bco.Safra": acao_safra, "Banco Safra": acao_safra, "BANCO SAFRA": acao_safra, "SAFRA": acao_safra, "Safra": acao_safra,
        "Bco.Suisse": acao_suisse, "Banco Suisse": acao_suisse, "Banco Credit Suisse": acao_suisse, "Credit Suisse": acao_suisse, "CREDIT SUISSE": acao_suisse, "SUISSE": acao_suisse, "BANCO SUISSE": acao_suisse,
        "Bco.Daycoval": acao_daycoval, "Banco Daycoval": acao_daycoval, "BANCO DAYCOVAL": acao_daycoval, "DAYCOVAL": acao_daycoval, "Daycoval": acao_daycoval,
        "Bco.Itaú": acao_itau, "Banco Itaú": acao_itau, "BANCO ITAÚ": acao_itau, "ITAÚ": acao_itau, "Itaú": acao_itau, "Itau": acao_itau, "ITAU": acao_itau,
    }

    if nome_banco in acoes:
        acoes[nome_banco]()
    else:
        print(f"Executando ação padrão para {nome_banco}.")

def abrir_explorador_e_detectar_banco():
    root = tk.Tk()
    root.withdraw()

    nome_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
    )

    if nome_arquivo:
        detectar_banco(nome_arquivo)
    else:
        print("Nenhum arquivo selecionado.")


abrir_explorador_e_detectar_banco()