import os

def verificar_pastas_recursivamente(diretorio):
    for pasta in os.listdir(diretorio):  # Percorre apenas as pastas diretas dentro do diretório principal
        caminho_pasta = os.path.join(diretorio, pasta)
        
        if os.path.isdir(caminho_pasta):  # Confirma que é uma pasta
            verificar_pasta(caminho_pasta)  # Verifica se a pasta tem arquivos
            
            # Agora verifica as subpastas dessa pasta
            for subpasta in os.listdir(caminho_pasta):
                caminho_subpasta = os.path.join(caminho_pasta, subpasta)
                if os.path.isdir(caminho_subpasta):  # Confirma que é uma subpasta
                    verificar_pasta(caminho_subpasta)  # Verifica se a subpasta tem arquivos

def verificar_pasta(caminho):
    arquivos_na_pasta = [f for f in os.listdir(caminho) if os.path.isfile(os.path.join(caminho, f))]
    
    if arquivos_na_pasta:
        print(f"A pasta '{caminho}' contém arquivos.")
    else:
        print(f"A pasta '{caminho}' está vazia.")

# Exemplo de uso com o caminho real
verificar_pastas_recursivamente(r"C:\Users\regina.santos\Desktop\Pessoal")



# armazenar




def verificar_pastas_recursivamente(diretorio):
    pastas_provisoes_vazias = []  # Lista para armazenar as pastas "Provisões" vazias

    for pasta in os.listdir(diretorio):  # Percorre apenas as pastas diretas dentro do diretório principal
        caminho_pasta = os.path.join(diretorio, pasta)
        
        if os.path.isdir(caminho_pasta):  # Confirma que é uma pasta
            for subpasta in os.listdir(caminho_pasta):  # Verifica subpastas dentro da pasta principal
                caminho_subpasta = os.path.join(caminho_pasta, subpasta)
                if os.path.isdir(caminho_subpasta) and subpasta.lower() == "provisões":  # Verifica se a subpasta se chama "Provisões"
                    verificar_pasta_provisoes(caminho_subpasta, pastas_provisoes_vazias)

    # Exibir todas as pastas "Provisões" vazias no final
    print("\n📂 Pastas 'Provisões' vazias encontradas:")
    for pasta in pastas_provisoes_vazias:
        print(f" - {pasta}")

def verificar_pasta_provisoes(caminho, pastas_provisoes_vazias):
    arquivos_na_pasta = [f for f in os.listdir(caminho) if os.path.isfile(os.path.join(caminho, f))]
    
    if arquivos_na_pasta:
        print(f"A pasta '{caminho}' contém arquivos.")
    else:
        print(f"A pasta '{caminho}' está vazia.")
        pastas_provisoes_vazias.append(caminho)  # Armazena a pasta "Provisões" vazia na lista

# Exemplo de uso com o caminho real
verificar_pastas_recursivamente(r"C:\Users\regina.santos\Desktop\Pessoal")



            #if arquivos_na_pasta:
                #print(f"A pasta '{caminho_completo}' contém arquivos.")
            #else:
                #print(f"A pasta '{caminho_completo}' está vazia.")


#def verificar_pastas(diretorio):
    #for pasta in os.listdir(diretorio):
        #caminho_completo = os.path.join(diretorio, pasta)
        #if os.path.isdir(caminho_completo):
            #arquivos_na_pasta = [f for f in os.listdir(caminho_completo) if os.path.isfile(os.path.join(caminho_completo, f))]
        
            #if arquivos_na_pasta:
               # print(f"A pasta '{caminho_completo} contém arquivos.")
           # else:
                #print(f"A pasta '{caminho_completo}' está vazia.")

#verificar_pastas("C:\\Users\\regina.santos\\Desktop\\Pessoal")