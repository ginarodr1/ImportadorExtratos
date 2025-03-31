import bcrypt

def hash_senha(senha):
    salt = bcrypt.gensalt() #? gera um salt aleatorio
    hashed = bcrypt.hashpw(senha.encode(), salt) #? cria o hash da senha com o salt
    return hashed

senha = "paxta"
senha_hash = hash_senha(senha)
print("Hash da senha:", senha_hash)
