import os

def verificar_pastas_recursivamente(diretorio):
    empresas_vazias = []  # Lista para armazenar os códigos das empresas com "Provisões" vazia

    for pasta in os.listdir(diretorio):  # Percorre apenas as pastas diretas dentro do diretório principal
        caminho_pasta = os.path.join(diretorio, pasta)
        
        if os.path.isdir(caminho_pasta):  # Confirma que é uma pasta
            for subpasta in os.listdir(caminho_pasta):  # Verifica subpastas dentro da pasta principal
                caminho_subpasta = os.path.join(caminho_pasta, subpasta)
                if os.path.isdir(caminho_subpasta) and subpasta.lower() == "provisões":  # Verifica se a subpasta se chama "Provisões"
                    if verificar_pasta_provisoes(caminho_subpasta):  # Se estiver vazia, armazena o código da empresa
                        empresas_vazias.append(pasta)

    # Exibir todas as empresas cujas pastas "Provisões" estão vazias
    print("\n🏢 Empresas com a pasta 'Provisões' vazia:")
    for empresa in empresas_vazias:
        print(f" - {empresa}")

    return empresas_vazias  # Retorna a lista de empresas com "Provisões" vazia

def verificar_pasta_provisoes(caminho):
    arquivos_na_pasta = [f for f in os.listdir(caminho) if os.path.isfile(os.path.join(caminho, f))]
    
    if arquivos_na_pasta:
        print(f"A pasta '{caminho}' contém arquivos.")
        return False  # Não está vazia
    else:
        print(f"A pasta '{caminho}' está vazia.")
        return True  # Está vazia

# Exemplo de uso com o caminho real
empresas_vazias = verificar_pastas_recursivamente(r"C:\Users\regina.santos\Desktop\Pessoal")

# Agora a variável empresas_vazias contém os códigos das empresas com a pasta "Provisões" vazia
print("\n📌 Lista final de empresas vazias:", empresas_vazias)





for item in empresas_vazias:
    print("Empresa vazia!")