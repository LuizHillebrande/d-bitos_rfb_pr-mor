import pandas as pd
from fuzzywuzzy import fuzz

# Função para normalizar as strings
def normalizar_string(texto):
    texto = texto.lower()  # Converte para minúsculas
    texto = texto.replace(".", "")  # Remove pontos
    texto = texto.replace(",", "")  # Remove vírgulas
    texto = texto.replace("e", "")  # Remove a palavra "e"
    texto = texto.replace("de", "")  # Remove a palavra "de"
    texto = texto.strip()  # Remove espaços extras
    return texto

# Carregar o arquivo Excel
file_path = "EMPRESAS FGTS.xlsx"
df = pd.read_excel(file_path)

# Verificar se a coluna "Razão Social" existe
if "Razão Social" not in df.columns:
    print("A coluna 'Razão Social' não existe no arquivo.")
else:
    # Normalizar as strings da coluna 'Razão Social'
    df['Razão Social Normalizada'] = df['Razão Social'].apply(normalizar_string)

    # Comparar todas as empresas
    for i, nome1 in enumerate(df['Razão Social Normalizada']):
        for j, nome2 in enumerate(df['Razão Social Normalizada']):
            # Ignorar comparações de uma empresa com ela mesma
            if i != j:
                # Calcula a similaridade entre os nomes
                similaridade = fuzz.ratio(nome1, nome2)
                
                # Se a similaridade for maior que 80%, há um erro
                if similaridade > 80:
                    print(f"Erro: As empresas '{df['Razão Social'][i]}' e '{df['Razão Social'][j]}' têm {similaridade}% de similaridade.")

    print("Comparação finalizada!")
