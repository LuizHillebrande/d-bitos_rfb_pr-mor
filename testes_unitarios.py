import os
import re

def extrair_nome_empresa_e_cnpj(nome_arquivo):
    """
    Extrai o nome da empresa e o CNPJ do nome do arquivo PDF.
    O nome do arquivo segue o formato 'situacao_fiscal--CNPJ-Nome_Arquivo.pdf'.
    """
    # Expressão regular para capturar o CNPJ
    cnpj = re.search(r"situacao_fiscal--(\d{14})-", nome_arquivo)
    if cnpj:
        cnpj = cnpj.group(1)  # Extrai o CNPJ

    # Remove o prefixo 'situacao_fiscal--CNPJ-' e qualquer código no final
    nome_limpo = re.sub(r"situacao_fiscal--\d{14}-", "", nome_arquivo)
    nome_limpo = re.sub(r"_[0-9]+\.pdf$", "", nome_limpo)  # Remove código final (se existir)
    
    return nome_limpo.strip(), cnpj

def renomear_pdfs_com_cnpj(pasta):
    """
    Itera sobre todos os PDFs da pasta e renomeia, usando o CNPJ como ID único.
    """
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".pdf"):
            caminho_antigo = os.path.join(pasta, arquivo)
            nome_empresa, cnpj = extrair_nome_empresa_e_cnpj(arquivo)
            
            # Criar o novo nome do arquivo com CNPJ
            if cnpj:
                novo_nome = f"{cnpj}_{nome_empresa}.pdf"
                caminho_novo = os.path.join(pasta, novo_nome)
                
                # Renomeia o arquivo
                os.rename(caminho_antigo, caminho_novo)
                print(f"Renomeado: {arquivo} → {novo_nome}")
            else:
                print(f"Não foi possível extrair o CNPJ de: {arquivo}")

# Defina a pasta onde estão os PDFs
pasta_pdfs = "debitos"

renomear_pdfs_com_cnpj(pasta_pdfs)