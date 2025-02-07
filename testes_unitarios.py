from fuzzywuzzy import fuzz

def comparar_empresas(nome_empresa1, nome_empresa2):
    # Calcula a similaridade entre as duas empresas usando fuzzywuzzy
    similaridade = fuzz.ratio(nome_empresa1, nome_empresa2)

    # Retorna o valor de similaridade
    return similaridade

# Testando a função
nome_empresa1 = "07611951000146_Silva E Lima Moveis Ltda"
nome_empresa2 = "07611951000227_Silva gianluquinhas"

similaridade = comparar_empresas(nome_empresa1, nome_empresa2)
print(f"A similaridade entre as duas empresas é: {similaridade}%")

