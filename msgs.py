import os
import pandas as pd
import re
from fuzzywuzzy import process
from datetime import datetime

# Caminho para a pasta "resultados"
diretorio_resultados = os.path.join(os.getcwd(), 'resultados')
# Caminhos para as pastas e arquivos
diretorio_codigos = os.path.join(os.getcwd(), 'resultados_codigos')
arquivo_tabelas = os.path.join(os.getcwd(), 'TABELASCDIGOSDERECEITA.xlsx')
tabela_depto_pessoal = pd.read_excel(arquivo_tabelas, sheet_name='Depto Pessoal')
tabela_fiscal = pd.read_excel(arquivo_tabelas, sheet_name='Fiscal')



def salvar_mensagem(df_existente, nome_empresa, nova_mensagem, caminho_saida):
    # Lista de empresas já existentes no arquivo
    nomes_existentes = df_existente['Empresa'].tolist()

    # Encontrar o nome mais parecido
    nome_mais_proximo, score = process.extractOne(nome_empresa, nomes_existentes) if nomes_existentes else (None, 0)

    # Se encontrou uma correspondência confiável, usa o nome existente
    if nome_mais_proximo and score >= 80:
        nome_empresa = nome_mais_proximo

    # Se a empresa já existir no arquivo, concatena a mensagem
    if nome_empresa in df_existente['Empresa'].values:
        df_existente.loc[df_existente['Empresa'] == nome_empresa, 'Mensagem'] += f"\n{nova_mensagem}"
    else:
        nova_linha = pd.DataFrame({"Empresa": [nome_empresa], "Mensagem": [nova_mensagem]})
        df_existente = pd.concat([df_existente, nova_linha], ignore_index=True)

    return df_existente


import os
import pandas as pd
from datetime import datetime
import re

def criar_msgs(caminho_saida):
    data_atual = datetime.now().strftime("%d/%m/%y")
    
    # Percorre todos os arquivos Excel na pasta
    for arquivo in os.listdir(diretorio_resultados):

        if os.path.exists(caminho_saida):
            df_existente = pd.read_excel(caminho_saida)
        else:
            df_existente = pd.DataFrame(columns=["Empresa", "Mensagem"])

        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):  # Verifica se é um arquivo Excel
            caminho_arquivo = os.path.join(diretorio_resultados, arquivo)
            
            # Lê o arquivo Excel
            df = pd.read_excel(caminho_arquivo)
            
            # Garante que as colunas necessárias estão no DataFrame
            if {'EMPRESA', 'DÍVIDA ATIVA', 'NUMERO DO PROCESSO', 'SITUAÇÃO'}.issubset(df.columns):
                
                # Tenta extrair o CNPJ limpo (14 dígitos) da coluna "EMPRESA"
                cnpj = re.search(r'(\d{14})', str(df['EMPRESA'].iloc[0]))  # Supondo que o CNPJ esteja na primeira linha
                if cnpj:
                    cnpj = cnpj.group(1)  # Extrai o CNPJ limpo
                    
                    # Remover o CNPJ do nome da empresa para utilizá-lo na mensagem
                    nome_empresa_sem_cnpj = df['EMPRESA'].iloc[0].replace(cnpj + "_", "")  # Remove o CNPJ do início do nome
                    
                    print(f"🔍 Buscando pelo CNPJ: {cnpj}")
                    
                    # Agrupa os processos pela mesma situação
                    situacoes = df.groupby('SITUAÇÃO')['NUMERO DO PROCESSO'].apply(list).to_dict()
                    
                    # Gera a mensagem personalizada para a empresa (usando o nome sem o CNPJ)
                    mensagem = f"{nome_empresa_sem_cnpj}, com base em consulta no dia {data_atual}, \n"
                    for situacao, processos in situacoes.items():
                        processos_formatados = ', '.join(processos)  # Junta os números dos processos
                        mensagem += f"Você tem os processos {processos_formatados} em '{situacao}'.\n"

                    df_existente = salvar_mensagem(df_existente, cnpj, mensagem.strip(), caminho_saida)
                    
                    print(f"Mensagem para {nome_empresa_sem_cnpj}:\n{mensagem}\n")
                else:
                    print(f"⚠️ CNPJ não encontrado para a empresa '{df['EMPRESA'].iloc[0]}'.")
            else:
                print(f"O arquivo {arquivo} não possui as colunas esperadas.")
            
        df_existente.to_excel(caminho_saida, index=False)
        print("Mensagens salvas com sucesso!")



def criar_msgs_codigos(diretorio_codigos, tabela_depto_pessoal, tabela_fiscal, caminho_saida):
    data_atual = datetime.now().strftime("%d/%m/%y")
    if os.path.exists(caminho_saida):
        df_existente = pd.read_excel(caminho_saida)
    else:
        df_existente = pd.DataFrame(columns=["Empresa", "Mensagem"])

    for arquivo in os.listdir(diretorio_codigos):
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
            caminho_arquivo = os.path.join(diretorio_codigos, arquivo)
            df = pd.read_excel(caminho_arquivo)

            if {'Empresa', 'Código Fiscal', 'PA - Exercício', 'Saldo Devedor Consignado'}.issubset(df.columns):
                # Tenta extrair o CNPJ limpo (14 dígitos) da coluna "Empresa"
                cnpj = re.search(r'(\d{14})', str(df['Empresa'].iloc[0]))  # Supondo que o CNPJ esteja na primeira linha
                if cnpj:
                    cnpj = cnpj.group(1)  # Extrai o CNPJ limpo
                    
                    # Remover o CNPJ do nome da empresa para utilizá-lo na mensagem
                    nome_empresa_sem_cnpj = df['Empresa'].iloc[0].replace(cnpj + "_", "")  # Remove o CNPJ do início do nome
                    
                    print(f"🔍 Buscando pelo CNPJ: {cnpj}")
                    
                    # Função para ajustar o formato do PA - Exercício
                    def formatar_pa_exercicio(pa_exercicio):
                        try:
                            if len(str(pa_exercicio).split('/')) == 3:  # Caso DDD/MM/YYYY
                                return '/'.join(str(pa_exercicio).split('/')[1:])
                            return str(pa_exercicio)
                        except Exception as e:
                            print(f"Erro ao formatar PA - Exercício: {pa_exercicio}, erro: {e}")
                            return None

                    df['PA - Exercício'] = df['PA - Exercício'].apply(formatar_pa_exercicio)

                    # Agrupando os dados pelo PA - Exercício
                    meses_agrupados = df.groupby('PA - Exercício')

                    mensagem = f"Olá {nome_empresa_sem_cnpj},\nSegue o resumo dos seus débitos consultados no dia {data_atual}:\n\n"

                    for pa_exercicio, grupo in meses_agrupados:
                        mensagem += f"**Referente a {pa_exercicio}:**\n"
                        debitos_por_tipo = {}

                        for _, row in grupo.iterrows():
                            codigo_fiscal_completo = str(row['Código Fiscal']).strip()
                            saldo_devedor = str(row['Saldo Devedor Consignado']).replace(',', '.')

                            try:
                                saldo_devedor = float(saldo_devedor)
                            except ValueError:
                                saldo_devedor = 0.0  # Caso o valor não seja numérico, considera como zero

                            if saldo_devedor <= 0:
                                continue  # Ignora débitos zerados

                            match = re.match(r'(\d+)[-/](\d+)', codigo_fiscal_completo)
                            if match:
                                codigo_fiscal_formatado_original = f"{match.group(1)}-{match.group(2)}"
                                codigo_fiscal_com_variacao = f"{match.group(1)}/{match.group(2)}"
                            else:
                                codigo_fiscal_formatado_original = codigo_fiscal_completo
                                codigo_fiscal_com_variacao = codigo_fiscal_completo

                            # Verifica em qual tabela o código está presente
                            if (codigo_fiscal_formatado_original in tabela_depto_pessoal['Código de receita'].astype(str).values or
                                codigo_fiscal_com_variacao in tabela_depto_pessoal['Código de receita'].astype(str).values):
                                tipo_debito = "Departamento Pessoal"
                            elif (codigo_fiscal_formatado_original in tabela_fiscal['Código de receita'].astype(str).values or
                                  codigo_fiscal_com_variacao in tabela_fiscal['Código de receita'].astype(str).values):
                                tipo_debito = "Fiscal"
                            else:
                                descricao = re.sub(r'^\d+[-/]\d+\s-\s', '', codigo_fiscal_completo)
                                tipo_debito = f"outros ({descricao})"

                            # Soma os valores por tipo de débito
                            if tipo_debito in debitos_por_tipo:
                                debitos_por_tipo[tipo_debito] += saldo_devedor
                            else:
                                debitos_por_tipo[tipo_debito] = saldo_devedor

                        # Adiciona os valores somados à mensagem
                        for tipo, valor in debitos_por_tipo.items():
                            mensagem += f"  - {tipo}: R$ {valor:.2f}\n"
                        
                        mensagem += "\n"  # Separação entre meses

                    df_existente = salvar_mensagem(df_existente, cnpj, mensagem.strip(), caminho_saida)
                    print(f"Mensagem gerada para {nome_empresa_sem_cnpj}:\n{mensagem}\n")
                else:
                    print(f"⚠️ CNPJ não encontrado para a empresa '{df['Empresa'].iloc[0]}'.")
            else:
                print(f"O arquivo {arquivo} não possui as colunas esperadas.")

    df_existente.to_excel(caminho_saida, index=False)
    print("Mensagens salvas com sucesso!")


# Chamada da função
criar_msgs_codigos(diretorio_codigos, tabela_depto_pessoal, tabela_fiscal, caminho_saida = 'mensagens.xlsx')
criar_msgs(caminho_saida="mensagens.xlsx")
