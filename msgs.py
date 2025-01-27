import os
import pandas as pd

# Caminho para a pasta "resultados"
diretorio_resultados = os.path.join(os.getcwd(), 'resultados')

def criar_msgs():
    # Percorre todos os arquivos Excel na pasta
    for arquivo in os.listdir(diretorio_resultados):
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):  # Verifica se é um arquivo Excel
            caminho_arquivo = os.path.join(diretorio_resultados, arquivo)
            
            # Lê o arquivo Excel
            df = pd.read_excel(caminho_arquivo)
            
            # Garante que as colunas necessárias estão no DataFrame
            if {'EMPRESA', 'DÍVIDA ATIVA', 'NUMERO DO PROCESSO', 'SITUAÇÃO'}.issubset(df.columns):
                nome_empresa = df['EMPRESA'].iloc[0]  # Considera que todas as linhas têm o mesmo nome de empresa
                
                # Agrupa os processos pela mesma situação
                situacoes = df.groupby('SITUAÇÃO')['NUMERO DO PROCESSO'].apply(list).to_dict()
                
                # Gera a mensagem personalizada para a empresa
                mensagem = f"Olá {nome_empresa},\n"
                for situacao, processos in situacoes.items():
                    processos_formatados = ', '.join(processos)  # Junta os números dos processos
                    mensagem += f"Você tem os processos {processos_formatados} em '{situacao}'.\n"
                
                print(f"Mensagem para {nome_empresa}:\n{mensagem}\n")
            else:
                print(f"O arquivo {arquivo} não possui as colunas esperadas.")

import os
import pandas as pd
import re

# Caminhos para as pastas e arquivos
diretorio_codigos = os.path.join(os.getcwd(), 'resultados_codigos')
arquivo_tabelas = os.path.join(os.getcwd(), 'TABELASCDIGOSDERECEITA.xlsx')

# Carrega as tabelas de códigos
tabela_depto_pessoal = pd.read_excel(arquivo_tabelas, sheet_name='Depto Pessoal')
tabela_fiscal = pd.read_excel(arquivo_tabelas, sheet_name='Fiscal')

import re

def criar_msgs_codigos():
    for arquivo in os.listdir(diretorio_codigos):
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
            caminho_arquivo = os.path.join(diretorio_codigos, arquivo)

            # Lê o arquivo Excel
            df = pd.read_excel(caminho_arquivo)

            # Verifica se as colunas necessárias existem
            if {'Empresa', 'Código Fiscal', 'PA - Exercício', 'Saldo Devedor Consignado'}.issubset(df.columns):
                nome_empresa = df['Empresa'].iloc[0]  # Nome da empresa

                # Gera mensagem personalizada para cada linha
                mensagem = f"Olá {nome_empresa},\n"
                for _, row in df.iterrows():
                    codigo_fiscal_completo = row['Código Fiscal']
                    pa_exercicio = row['PA - Exercício']
                    saldo_devedor = row['Saldo Devedor Consignado']

                    # Substitui o hífen por barra e remove a descrição
                    codigo_fiscal_formatado = re.sub(r'(\d+)-(\d+)\s*-\s*.*', r'\1/\2', codigo_fiscal_completo)

                    # Imprime para verificar o código fiscal formatado
                    print(f"Código fiscal formatado: {codigo_fiscal_formatado}")
                    
                    # Verifica em qual tabela o código está presente
                    if codigo_fiscal_formatado in tabela_depto_pessoal['Código de receita'].astype(str).values:
                        tipo_debito = "relacionado ao departamento pessoal"
                    elif codigo_fiscal_formatado in tabela_fiscal['Código de receita'].astype(str).values:
                        tipo_debito = "relacionado ao fiscal"
                    else:
                        # Usa a descrição após o código como mensagem
                        descricao = re.sub(r'^\d+[-/]\d+\s-\s', '', codigo_fiscal_completo)  # Remove o código e o traço/barra
                        tipo_debito = f"relacionado a {descricao}"

                    # Adiciona à mensagem
                    mensagem += (
                        f"Você tem um débito no valor de {saldo_devedor} com vencimento em {pa_exercicio}, "
                        f"{tipo_debito}.\n"
                    )

                print(f"Mensagem para {nome_empresa}:\n{mensagem}\n")
            else:
                print(f"O arquivo {arquivo} não possui as colunas esperadas.")

# Chamada da função
criar_msgs_codigos()

