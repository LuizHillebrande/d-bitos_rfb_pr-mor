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

import os
import pandas as pd
import re

import pandas as pd
import re

def criar_msgs_codigos(diretorio_codigos, tabela_depto_pessoal, tabela_fiscal):
    for arquivo in os.listdir(diretorio_codigos):
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
            caminho_arquivo = os.path.join(diretorio_codigos, arquivo)

            # Lê o arquivo Excel
            df = pd.read_excel(caminho_arquivo)

            # Verifica se as colunas necessárias existem
            if {'Empresa', 'Código Fiscal', 'PA - Exercício', 'Saldo Devedor Consignado'}.issubset(df.columns):
                nome_empresa = df['Empresa'].iloc[0]  # Nome da empresa

                # Função para ajustar o formato do PA - Exercício
                def formatar_pa_exercicio(pa_exercicio):
                    try:
                        # Se o formato for DDD/MM/YYYY, transforma em MM/YYYY
                        if len(str(pa_exercicio).split('/')) == 3:  # Caso DDD/MM/YYYY
                            return '/'.join(str(pa_exercicio).split('/')[1:])
                        # Se for MM/YYYY, deixa como está
                        return str(pa_exercicio)
                    except Exception as e:
                        print(f"Erro ao formatar PA - Exercício: {pa_exercicio}, erro: {e}")
                        return None

                # Aplica a função para formatar a coluna 'PA - Exercício'
                df['PA - Exercício'] = df['PA - Exercício'].apply(formatar_pa_exercicio)

                # Filtra os dados para agrupar os que têm o mesmo mês/ano
                meses_agrupados = df.groupby('PA - Exercício')

                # Verificar e imprimir os agrupamentos
                for pa_exercicio, grupo in meses_agrupados:
                    print(f"\nGrupo para PA - Exercício {pa_exercicio}:")
                    print(grupo[['Empresa', 'PA - Exercício', 'Saldo Devedor Consignado']])

                # Gera mensagem personalizada para cada linha
                mensagem = f"Olá {nome_empresa},\n"
                for _, row in df.iterrows():
                    codigo_fiscal_completo = str(row['Código Fiscal']).strip()
                    pa_exercicio = row['PA - Exercício']
                    
                    # Converte saldo_devedor, substituindo vírgula por ponto e tratando como float
                    saldo_devedor = str(row['Saldo Devedor Consignado']).replace(',', '.')
                    try:
                        saldo_devedor = float(saldo_devedor)
                    except ValueError:
                        saldo_devedor = 0.0  # Caso o valor não seja numérico, considera como zero

                    # Garante que o código esteja no formato esperado
                    match = re.match(r'(\d+)[-/](\d+)', codigo_fiscal_completo)
                    if match:
                        codigo_fiscal_formatado_original = f"{match.group(1)}-{match.group(2)}"
                        codigo_fiscal_com_variacao = f"{match.group(1)}/{match.group(2)}"
                    else:
                        codigo_fiscal_formatado_original = codigo_fiscal_completo
                        codigo_fiscal_com_variacao = codigo_fiscal_completo

                    # Verifica em qual tabela o código está presente (em qualquer formato)
                    if (codigo_fiscal_formatado_original in tabela_depto_pessoal['Código de receita'].astype(str).values or
                        codigo_fiscal_com_variacao in tabela_depto_pessoal['Código de receita'].astype(str).values):
                        tipo_debito = "relacionado ao departamento pessoal"
                    elif (codigo_fiscal_formatado_original in tabela_fiscal['Código de receita'].astype(str).values or
                          codigo_fiscal_com_variacao in tabela_fiscal['Código de receita'].astype(str).values):
                        tipo_debito = "relacionado ao fiscal"
                    else:
                        descricao = re.sub(r'^\d+[-/]\d+\s-\s', '', codigo_fiscal_completo)
                        tipo_debito = f"relacionado a {descricao}"

                    # Adiciona à mensagem apenas se o saldo devedor for diferente de zero
                    if saldo_devedor > 0:
                        mensagem += (
                            f"Você tem um débito no valor de {saldo_devedor:.2f} com vencimento em {pa_exercicio}, "
                            f"{tipo_debito}.\n"
                        )

                print(f"Mensagem para {nome_empresa}:\n{mensagem}\n")
            else:
                print(f"O arquivo {arquivo} não possui as colunas esperadas.")



# Chamada da função
criar_msgs_codigos(diretorio_codigos, tabela_depto_pessoal, tabela_fiscal)
