from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pyautogui
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import os
import re
import pdfplumber
import pandas as pd
import customtkinter as ctk
import json
from time import sleep
import zipfile
import os
import shutil
from thefuzz import process 
from selenium.webdriver.common.by import By

import os
import pandas as pd
import re
from fuzzywuzzy import process
from datetime import datetime
import textwrap

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

    # Mostrar os nomes das empresas existentes no DataFrame
    print("Empresas existentes no arquivo:", nomes_existentes)

    # Encontrar o nome mais parecido
    nome_mais_proximo, score = process.extractOne(nome_empresa, nomes_existentes) if nomes_existentes else (None, 0)

    # Mostrar o nome mais próximo e o score
    print(f"Procurando pelo nome: {nome_empresa}")
    print(f"Nome mais próximo encontrado: {nome_mais_proximo}, Score: {score}")

    # Se encontrou uma correspondência confiável, usa o nome existente
    if nome_mais_proximo and score >= 80:
        nome_empresa = nome_mais_proximo
        print(f"Usando nome mais próximo: {nome_empresa}")

    # Se a empresa já existir no arquivo, concatena a mensagem
    if nome_empresa in df_existente['Empresa'].values:
        print(f"Empresa '{nome_empresa}' encontrada no arquivo, concatenando mensagem...")
        df_existente.loc[df_existente['Empresa'] == nome_empresa, 'Mensagem'] += f"\n{nova_mensagem}"
    else:
        print(f"Empresa '{nome_empresa}' não encontrada no arquivo, criando nova linha...")
        nova_linha = pd.DataFrame({"Empresa": [nome_empresa], "Mensagem": [nova_mensagem]})
        df_existente = pd.concat([df_existente, nova_linha], ignore_index=True)

    return df_existente



import os
import pandas as pd
from datetime import datetime
import re

diretorio_processos_sief = os.path.join(os.getcwd(), 'processos sief')

def criar_msgs_processos_sief(caminho_saida, diretorio_processos_sief):
    from datetime import datetime
    import os
    import pandas as pd

    data_atual = datetime.now().strftime("%d/%m/%y")
    
    # Verifica se já existe um arquivo com mensagens e carrega os dados
    if os.path.exists(caminho_saida):
        df_existente = pd.read_excel(caminho_saida)
    else:
        df_existente = pd.DataFrame(columns=["Empresa", "Mensagem"])

    # Percorre todos os arquivos Excel na pasta
    for arquivo in os.listdir(diretorio_processos_sief):
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
            # Exemplo do nome do arquivo: "23098061000139_C R V ESTERO E CIA LTDA.xlsx"
            # Extrai o CNPJ que está antes do primeiro '_'
            cnpj = arquivo.split('_')[0]
            print(f"🔍 CNPJ extraído do nome do arquivo: {cnpj}")

            # Monta o caminho completo do arquivo
            caminho_arquivo = os.path.join(diretorio_processos_sief, arquivo)
            
            # Lê o arquivo Excel
            df = pd.read_excel(caminho_arquivo)
            
            # Verifica se a coluna necessária existe
            if 'Processos SIEF' in df.columns:
                # Lista de processos (removendo valores vazios)
                processos = df["Processos SIEF"].dropna().astype(str).tolist()

                match = re.match(r'^\d+_(.*)\.xlsx$', arquivo)
                if match:
                    nome_empresa_sem_cnpj = match.group(1)
                else:
                    nome_empresa_sem_cnpj = "Nome não encontrado"
                
                # Gera a mensagem personalizada usando somente o CNPJ como identificador
                mensagem = f"\n\nA empresa {nome_empresa_sem_cnpj} possui os seguintes débitos referentes a Processos SIEF:\n\n"
                mensagem += ', '.join(processos)
                
                # Salva ou concatena a mensagem no DataFrame existente, usando o CNPJ como chave
                df_existente = salvar_mensagem(df_existente, cnpj, mensagem.strip(), caminho_saida)
                
                print(f"✅ Mensagem gerada para {cnpj}:\n{mensagem}\n")
            else:
                print(f"⚠️ O arquivo {arquivo} não possui a coluna 'Processos SIEF' esperada.")

    # Salva as mensagens geradas no arquivo Excel
    df_existente.to_excel(caminho_saida, index=False)
    print("✅ Mensagens salvas com sucesso!")

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
                    mensagem = f"A empresa possui os seguintes débitos referente a parcelamentos: \n"
                    for situacao, processos in situacoes.items():
                        processos_formatados = ', '.join(processos)  # Junta os números dos processos
                        mensagem += f"{situacao}'.\n"

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

                    mensagem = f"Olá {nome_empresa_sem_cnpj},\n"
                    mensagem += "Identificamos que sua empresa possui algumas pendências em aberto junto à Receita Federal.\n"
                    mensagem += "Essas pendências podem gerar multas, juros e complicações mais sérias se não forem regularizadas em tempo hábil.\n\n"
                    mensagem += "Segue o resumo dos seus débitos:\n\n"
                    mensagem = textwrap.dedent(mensagem)




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
import pandas as pd

def criar_msg_fgts():
    # Carregar os arquivos
    fgts_df = pd.read_excel("debitos_fgts.xlsx")
    mensagens_df = pd.read_excel("mensagens.xlsx")

    # Criar um dicionário para agrupar os débitos por empresa
    fgts_dict = {}
    for _, row in fgts_df.iterrows():
        nome_completo = row["Nome da Empresa"]
        cnpj, nome_empresa = nome_completo.split("_", 1)
        mes_ref = row["Mês Ref."]
        valor = row["Valor Débitos"]
        
        if cnpj not in fgts_dict:
            fgts_dict[cnpj] = {"nome": nome_empresa, "debitos": {}}
        
        if mes_ref not in fgts_dict[cnpj]["debitos"]:
            fgts_dict[cnpj]["debitos"][mes_ref] = 0
        
        fgts_dict[cnpj]["debitos"][mes_ref] += valor

    # Criar ou atualizar as mensagens
    for cnpj, data in fgts_dict.items():
        nome_empresa = data["nome"]
        debitos_texto = ", ".join([f"{mes}: R$ {valor:.2f}" for mes, valor in data["debitos"].items()])
        
        if cnpj in mensagens_df["Empresa"].astype(str).values:
            print('tinha o cnpj', cnpj)
            mensagem_fgts = f"{nome_empresa}, você também possui débitos de FGTS: " + ", ".join(
                [f"{mes} no valor de R$ {valor:.2f}" for mes, valor in data['debitos'].items()]
            ) + "."
            mensagens_df.loc[mensagens_df["Empresa"].astype(str) == cnpj, "Mensagem"] += f" {mensagem_fgts}"
        else:
            mensagem = f"{nome_empresa}, segue resumo dos seus débitos de FGTS: {debitos_texto}."
            mensagens_df = pd.concat([mensagens_df, pd.DataFrame({"Empresa": [cnpj], "Mensagem": [mensagem]})], ignore_index=True)

    # Salvar o arquivo atualizado
    mensagens_df.to_excel("mensagens.xlsx", index=False)

    print("Mensagens de FGTS geradas e salvas com sucesso!")


def criar_msg_final():
    # Carregar o arquivo de mensagens
    mensagens_df = pd.read_excel("mensagens.xlsx")

    # Definir a mensagem final
    data_atual = datetime.now().strftime("%d/%m/%y")
    mensagem_final = (
        f"\nOs valores informados são válidos na data de envio deste e-mail ({data_atual}) e podem sofrer alterações.\n"
        "Caso tenha interesse em regularizar essas pendências, entre em contato com o nosso time "
        "para mais detalhes e orientações sobre os próximos passos.\n"
        "Ficamos à disposição para qualquer dúvida ou informação adicional!\n\n"
        "Atenciosamente,\n"
        "Prímor Contábil\n"
        "(44) 98462-9927 / atendimento@contabilprimor.com.br"
    )

    # Garantir que a mensagem final seja a última coisa adicionada a cada linha
    mensagens_df["Mensagem"] = mensagens_df["Mensagem"].astype(str) + mensagem_final

    # Salvar as mensagens atualizadas
    mensagens_df.to_excel("mensagens.xlsx", index=False)

    print("Mensagem final adicionada com sucesso!")





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

def consultar_pdf_da_empresa(nome_empresa, numeros_procurados):
    # Caminho da pasta onde os PDFs foram baixados/descompactados
    pasta_debitos = os.path.join(os.getcwd(), 'debitos')

    # Listar todos os arquivos PDF na pasta
    arquivos_pdf = [f for f in os.listdir(pasta_debitos) if f.endswith('.pdf')]

    # Tentar extrair o CNPJ limpo (sem pontos, barras ou traços) do nome da empresa
    cnpj = re.search(r'(\d{14})', nome_empresa)
    
    if cnpj:
        cnpj = cnpj.group(1)  # Extrai o CNPJ limpo
        print(f"🔍 Buscando pelo CNPJ: {cnpj}")
        
        # Buscar diretamente o PDF com base no CNPJ limpo
        nome_pdf_proximo = None
        for arquivo in arquivos_pdf:
            if cnpj in arquivo:  # Verifica se o CNPJ está no nome do arquivo PDF
                nome_pdf_proximo = arquivo
                break

        if nome_pdf_proximo:
            caminho_pdf = os.path.join(pasta_debitos, nome_pdf_proximo)
            print(f"🔍 PDF encontrado: {nome_pdf_proximo}")

            # Aqui você pode usar uma biblioteca para abrir o PDF, por exemplo, pdfplumber
            abrir_pdf(caminho_pdf, numeros_procurados, nome_empresa)
        else:
            print(f"⚠️ Nenhum PDF encontrado para o CNPJ '{cnpj}'.")
    else:
        print(f"⚠️ CNPJ não encontrado no nome da empresa '{nome_empresa}'.")

#busca os numeros de dividas ativas aqui


def abrir_pdf(caminho_pdf, numeros_procurados, nome_empresa, pasta_destino="resultados"):
    """
    Abre o PDF, procura por números específicos, combina texto de todas as páginas 
    e captura a situação associada ao número.
    """
    resultados = []  # Lista para armazenar os dados processados

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            # Combinar o texto de todas as páginas do PDF
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"

            # Procurar os números no texto combinado
            for numero in numeros_procurados:
                if str(numero) in texto_completo:
                    print(f"⚠️ Número encontrado no PDF: {numero}")

                    # Verificar qualquer situação associada ao número
                    padrao = rf"{numero}.*?Situação:\s+([^\n]+)"
                    match = re.search(padrao, texto_completo, re.DOTALL)
                    
                    if match:
                        situacao = match.group(1).strip()
                        print(f"✅ Situação do número {numero}: {situacao}")
                        
                        # Adiciona os dados na lista
                        resultados.append({
                            "EMPRESA": nome_empresa,
                            "DÍVIDA ATIVA": "SIM",
                            "NUMERO DO PROCESSO": numero,
                            "SITUAÇÃO": situacao
                        })
                    else:
                        print(f"⚠️ Situação do número {numero} não encontrada.")
        
        # Criar a pasta de destino se não existir
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)
        
        # Salvar resultados no Excel
        if resultados:
            caminho_excel = os.path.join(pasta_destino, f"{nome_empresa}_resultados.xlsx")
            df = pd.DataFrame(resultados)
            df.to_excel(caminho_excel, index=False)
            print(f"✅ Resultados salvos em: {caminho_excel}")
        else:
            print(f"⚠️ Nenhum dado encontrado para a empresa: {nome_empresa}")

    except Exception as e:
        print(f"Erro ao abrir ou processar o arquivo PDF {caminho_pdf}: {e}")



def processar_excel_e_abrir_pdf():
    """
    Processa os arquivos Excel na pasta 'dividas ativas', extrai os números das dívidas e
    busca por esses números nos PDFs relacionados.
    """
    pasta_destino = "dividas ativas"
    
    for excel_file in os.listdir(pasta_destino):
        if excel_file.endswith('.xlsx'):
            # Caminho completo do arquivo Excel
            caminho_excel = os.path.join(pasta_destino, excel_file)

            # Ler os números do Excel
            df = pd.read_excel(caminho_excel)
            numeros_procurados = df["Inscrição da Dívida"].astype(str).tolist()
            
            # Extrair o nome da empresa a partir do nome do arquivo Excel
            nome_empresa = os.path.splitext(excel_file)[0].replace(" LTDA", "")
            print(f"Procurando números no PDF para a empresa: {nome_empresa}")

            # Localizar o PDF correspondente
            consultar_pdf_da_empresa(nome_empresa, numeros_procurados)

#CONSULTA PDF DOS CODIGOS FISCAIS
def consultar_pdf_da_empresa_codigos(nome_empresa, numeros_procurados):
    """
    Procura o PDF correspondente ao nome da empresa e chama a função para extrair os valores.
    """
    # Caminho da pasta onde os PDFs estão
    pasta_debitos = os.path.join(os.getcwd(), 'debitos')

    # Listar todos os arquivos PDF na pasta
    arquivos_pdf = [f for f in os.listdir(pasta_debitos) if f.endswith('.pdf')]

    # Tentar extrair o CNPJ limpo (sem pontos, barras ou traços) do nome da empresa
    cnpj = re.search(r'(\d{14})', nome_empresa)
    
    if cnpj:
        cnpj = cnpj.group(1)  # Extrai o CNPJ limpo
        print(f"🔍 Buscando pelo CNPJ: {cnpj}")
        
        # Buscar diretamente o PDF com base no CNPJ limpo
        nome_pdf_proximo = None
        for arquivo in arquivos_pdf:
            if cnpj in arquivo:  # Verifica se o CNPJ está no nome do arquivo PDF
                nome_pdf_proximo = arquivo
                break

        if nome_pdf_proximo:
            caminho_pdf = os.path.join(pasta_debitos, nome_pdf_proximo)
            print(f"🔍 PDF encontrado: {nome_pdf_proximo}")

            # Chamar a função para processar o PDF
            abrir_pdf_codigos(caminho_pdf, numeros_procurados, nome_empresa)
        else:
            print(f"⚠️ Nenhum PDF encontrado para o CNPJ '{cnpj}'.")
    else:
        print(f"⚠️ CNPJ não encontrado no nome da empresa '{nome_empresa}'.")

def abrir_pdf_codigos(caminho_pdf, numeros_procurados, nome_empresa):
    """
    Abre o PDF, busca os códigos fiscais e PA - EXERC. e extrai o saldo devedor consignado.
    """
    resultados = []  # Lista para armazenar os resultados

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            # Combinar o texto de todas as páginas do PDF
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"

            # Iterar pelos códigos fiscais e PA-EXERC. fornecidos
            for codigo, pa_exerc in numeros_procurados:
                # Regex para encontrar a linha correspondente no PDF
                padrao = rf"{codigo}\s+{pa_exerc}.*?([\d.,]+)\s*DEVEDOR$"
                match = re.search(padrao, texto_completo, re.MULTILINE)

                if match:
                    saldo_devedor = match.group(1).strip()
                    print(f"✅ Código: {codigo}, PA-EXERC.: {pa_exerc}, Saldo: {saldo_devedor}")
                    
                    # Adicionar o resultado
                    resultados.append({
                        "Empresa": nome_empresa,
                        "Código Fiscal": codigo,
                        "PA - Exercício": pa_exerc,
                        "Saldo Devedor Consignado": saldo_devedor
                    })
                else:
                    print(f"⚠️ Nenhuma linha encontrada para Código: {codigo}, PA-EXERC.: {pa_exerc}")

        # Salvar os resultados em um Excel
        if resultados:
            salvar_resultados_excel(resultados, nome_empresa)
        else:
            print(f"⚠️ Nenhum dado extraído do PDF para a empresa: {nome_empresa}")

    except Exception as e:
        print(f"Erro ao abrir ou processar o PDF {caminho_pdf}: {e}")



def salvar_resultados_excel(resultados, nome_empresa):
    """
    Salva os resultados extraídos em um arquivo Excel.
    """
    # Criar um DataFrame com os resultados
    df = pd.DataFrame(resultados)

    # Caminho para salvar os resultados
    pasta_destino = os.path.join(os.getcwd(), 'resultados_codigos')
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    caminho_excel = os.path.join(pasta_destino, f"{nome_empresa}_codigos_fiscais.xlsx")
    df.to_excel(caminho_excel, index=False)
    print(f"✅ Resultados salvos em: {caminho_excel}")


def processar_empresas_codigos():
    """
    Itera sobre os arquivos Excel na pasta 'codigos fiscais' e processa os PDFs correspondentes.
    """
    pasta_codigos = os.path.join(os.getcwd(), 'codigos fiscais')

    for arquivo_excel in os.listdir(pasta_codigos):
        if arquivo_excel.endswith('.xlsx'):
            caminho_excel = os.path.join(pasta_codigos, arquivo_excel)
            nome_empresa = os.path.splitext(arquivo_excel)[0]

            # Ler o Excel para obter os códigos fiscais e PA-EXERC.
            df = pd.read_excel(caminho_excel)

            numeros_procurados = list(zip(df["Codigos Fiscais"], df["PA - EXERC."]))

            print(f"🔍 Processando empresa: {nome_empresa}")
            consultar_pdf_da_empresa_codigos(nome_empresa, numeros_procurados)


#salva os debitos de divida ativa
def salvar_numeros_em_excel(lista_numeros, nome_arquivo, pasta_destino):
    # Salvar a lista de números em um arquivo Excel
    df = pd.DataFrame(lista_numeros, columns=["Inscrição da Dívida"])
    caminho_arquivo = os.path.join(pasta_destino, f"{nome_arquivo}.xlsx")
    df.to_excel(caminho_arquivo, index=False)
    print(f"Arquivo salvo em {caminho_arquivo}")

import pandas as pd
import os

def salvar_processos_em_excel(lista_processos, nome_arquivo, pasta_destino, nome_empresa):
    # Criar o DataFrame
    df = pd.DataFrame(lista_processos, columns=["Processos SIEF"])

    # Adicionar a coluna "Nome Empresa" na primeira posição
    df.insert(0, "Nome Empresa", nome_empresa)  

    # Criar caminho do arquivo
    caminho_arquivo = os.path.join(pasta_destino, f"{nome_arquivo}.xlsx")
    
    # Salvar no Excel
    df.to_excel(caminho_arquivo, index=False)
    print(f"Arquivo salvo em {caminho_arquivo}")

# Chamando a função corretamente


#codigos fiscais salvando

def salvar_codigos_em_excel(lista_numeros, lista_pa_exercicio, nome_arquivo, pasta_codigos):
    # Verifique se as listas têm o mesmo comprimento
    if len(lista_numeros) != len(lista_pa_exercicio):
        raise ValueError("As listas de números e de PA - EXERC. devem ter o mesmo comprimento.")

    # Verificar se o diretório existe, se não, criar o diretório
    if not os.path.exists(pasta_codigos):
        os.makedirs(pasta_codigos)
    
    # Salvar a lista de números e PA em um DataFrame
    df = pd.DataFrame({
        "Codigos Fiscais": lista_numeros,  # Lista de números
        "PA - EXERC.": lista_pa_exercicio  # Lista de PA - Exercicio
    })
    
    # Definir o caminho do arquivo Excel
    caminho_arquivo = os.path.join(pasta_codigos, f"{nome_arquivo}.xlsx")
    
    # Salvar o DataFrame no arquivo Excel
    df.to_excel(caminho_arquivo, index=False)
    
    print(f"Arquivo salvo em {caminho_arquivo}")

def descompactar_arquivo_zip(download_folder, driver):

    zip_file = None

    while not zip_file:
        # Verifica se há um arquivo ZIP na pasta
        for file in os.listdir(download_folder):
            if file.endswith('.zip'):
                zip_file = os.path.join(download_folder, file)
                break  # Sai do loop assim que encontrar o primeiro ZIP
                
        if not zip_file:
            pyautogui.press('f5')
            sleep(2)
            baixar_nuvem = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(@id, '__BV_toggle_') and contains(@class, 'nav-link dropdown-toggle')]"))
            )
            baixar_nuvem.click()
            print("Nenhum arquivo ZIP encontrado. Tentando novamente em 5 segundos...")
            pyautogui.click(667, 300, duration=1)  # Clica no botão
            time.sleep(10)  # Espera 10 segundos antes de tentar novamente

    # Se encontrou um ZIP, descompacta
    try:
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(download_folder)
        print(f"Arquivo ZIP {zip_file} descompactado.")

        # Excluir o arquivo ZIP
        os.remove(zip_file)
        print(f"Arquivo ZIP {zip_file} excluído.")
    except Exception as e:
        print(f"Erro ao descompactar ou excluir o arquivo ZIP: {e}")

# Caminho da pasta onde os PDFs foram descompactados
pasta_debitos = os.path.join(os.getcwd(), 'debitos')

# Função para carregar a tabela de códigos fiscais
def carregar_codigos_fiscais(caminho_arquivo_excel):
    df_depto = pd.read_excel(caminho_arquivo_excel, sheet_name="Depto Pessoal")
    df_fiscal = pd.read_excel(caminho_arquivo_excel, sheet_name="Fiscal")

    # Concatenar os dois dataframes
    df = pd.concat([df_depto, df_fiscal])

    # Criar um dicionário {codigo: descricao}
    codigos_fiscais = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1]))
    
    return codigos_fiscais

import shutil

def limpar_pastas():
    pastas = ['debitos', 'resultados', 'resultados_codigos', 'codigos fiscais', 'dividas ativas', 'processos sief']
    
    for pasta in pastas:
        if os.path.exists(pasta):
            for arquivo in os.listdir(pasta):
                caminho_arquivo = os.path.join(pasta, arquivo)
                if os.path.isfile(caminho_arquivo) or os.path.islink(caminho_arquivo):
                    os.unlink(caminho_arquivo)  # Remove arquivos e links simbólicos
                elif os.path.isdir(caminho_arquivo):
                    shutil.rmtree(caminho_arquivo)  # Remove subpastas
            print(f"Conteúdo da pasta '{pasta}' foi excluído.")
        else:
            print(f"A pasta '{pasta}' não existe.")



def login():
    limpar_pastas()
    download_folder = os.path.join(os.getcwd(), 'debitos')

    # Configuração para o Chrome salvar os PDFs diretamente na pasta especificada
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_folder,  # Caminho para a pasta onde o PDF será salvo
        "download.prompt_for_download": False,  # Desativa o prompt para confirmar download
        "plugins.always_open_pdf_externally": True  # Abre os PDFs diretamente sem pedir para abrir
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)
    driver.get('https://app.monitorcontabil.com.br/login')
    driver.maximize_window()

    try:
        confirm_ok = WebDriverWait(driver,3).until(
            EC.element_to_be_clickable((By.XPATH,"//button[@class='swal2-confirm swal-btn-green swal2-styled']"))
        )
        confirm_ok.click()
    except:
        print('N tinha botao de ok')

    sleep(2)

    logar_email = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='email']"))
    )
    logar_email.send_keys('taina@contabilprimor.com.br')

    logar_senha = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='senhaInput']"))
    )
    logar_senha.send_keys('Primor1214')

    logar = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@type='submit']"))
    )
    logar.click()

    sleep(3)
    pyautogui.press('esc')

    driver.get("https://app.monitorcontabil.com.br/situacao-fiscal/visualizar?busca=")

    sleep(2)

    
    filtro_irregulares = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, "//select[@name='filtroSelecao']"))
    )

    # Abre a lista de opções do select
    filtro_irregulares.click()

    # Seleciona a opção que contém "Irregulares"
    select = Select(filtro_irregulares)

    # Aqui vamos selecionar a opção "Irregulares"
    for option in select.options:
        if "Irregulares" in option.text:
            option.click()
            sleep(1)
            pyautogui.press('esc')
            break
        
    sleep(3)
    
    baixar_relatorios = WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Baixar situações']]"))
    )
    baixar_relatorios.click()


    baixar_popup = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@class='btn btn-sm btn-success']"))
    )
    baixar_popup.click()

    sleep(8)
    pyautogui.press('F5')
    sleep(2)

    baixar_nuvem = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(@id, '__BV_toggle_') and contains(@class, 'nav-link dropdown-toggle')]"))
    )
    baixar_nuvem.click()


    WebDriverWait(driver, 5).until(
        EC.invisibility_of_element_located((By.XPATH, "//small[@class='notification-text' and text()='Em Execução']"))
    )

    # Agora que o "Em Execução" não está mais visível, podemos clicar no botão de download
    #baixar_definitivamente = WebDriverWait(driver, 10).until(
        #EC.element_to_be_clickable((By.XPATH, "//div[@class='media mr-1']//div[@class='col-2']//svg[contains(@class, 'feather-download')]"))
    #)
    #baixar_definitivamente.click()


    pyautogui.click(667,300, duration = 1)

    sleep(5)
    descompactar_arquivo_zip(download_folder, driver)
    sleep(2)
    renomear_pdfs_com_cnpj(pasta_pdfs)
    pyautogui.press('Esc')
    sleep(2)

    filtro_a_z = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH, "//th[contains(@class, 'vgt-left-align') and contains(@class, 'sortable')]/span[text()='Razão social']"))
    )
    filtro_a_z.click()
    sleep(2)


    pasta_cnpjs = os.path.join(os.getcwd(), 'cnpj_empresas')
    if not os.path.exists(pasta_cnpjs):
        os.makedirs(pasta_cnpjs)

    # Nome do arquivo Excel
    arquivo_excel_cnpjs = 'empresas_cnpj.xlsx'
    caminho_excel_cnpjs = os.path.join(pasta_cnpjs, arquivo_excel_cnpjs)

    # Lista para armazenar os dados das empresas
    dados_cnpjs = []

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr.clickable")))

    # Encontrar todas as linhas da tabela
    while True:
        linhas = driver.find_elements(By.CSS_SELECTOR, "tr.clickable")
        start = len(linhas) - 2 
        for i, linha in enumerate(linhas[start:], start=start):  # Define o início da iteração
        #for i, linha in enumerate(linhas):
            linhas = driver.find_elements(By.CSS_SELECTOR, "tr.clickable")
            print("Numero de linhas:",len(linhas))
            print(f'Processando linha {i + 1} de {len(linhas)}')
            try:
                print(f'Processando linha {linhas.index(linha) + 1} de {len(linhas)}')
                # Extrair o CNPJ
                cnpj_element = linha.find_element(By.XPATH, ".//td[contains(@class, 'col-tamanho-cnpj')]//span/span/span")
                cnpj = re.sub(r'[^0-9]', '', cnpj_element.text)

                # Encontrar o Nome da Empresa dentro da linha específica
                nome_empresa_element = linha.find_element(By.XPATH, ".//td[contains(@class, 'vgt-left-align')]//span/span/span")
                nome_empresa = nome_empresa_element.text

                print(f"Empresaaa: {nome_empresa} | CNPJ: {cnpj}")

                # Encontrar e clicar na lupa correspondente
                lupa = linha.find_element(By.XPATH, ".//button[@class='btn btn btn-none rounded-pill m-0 icone-acao p-0 btn-none btn-none'][2]")
                
                # Usar JavaScript para garantir o clique, caso necessário
                driver.execute_script("arguments[0].click();", lupa)

                # Aguarde o carregamento da nova informação, se houver
                WebDriverWait(driver, 5).until(EC.staleness_of(lupa))  # Espera a lupa desaparecer/mudar
            except Exception as e:
                print(f"Erro ao processar uma linha: {e}")

                #tentando clicar em débitos(sief)
                try:    
                    debitos_sief = WebDriverWait(driver,5).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'list-group-item') and contains(@class, 'collapsed')]//span[text()='Débito (Sief)']"))
                    )
                    debitos_sief.click()
                    print('CLICADO EM DEBITOS SIEF')
                    numeros = driver.find_elements(By.XPATH, "//tr//td[@aria-colindex='1']//div[@class='ml-50']")
                    lista_numeros = [numero.text for numero in numeros]

                    print("Códigos fiscais:", lista_numeros, "\n")

                    pa_exercicio = driver.find_elements(By.XPATH, "//tr//td[@aria-colindex='2']//div[@class='ml-50']")
                    lista_pa_exercicio = [pa.text for pa in pa_exercicio]
                    print("Pa - exercício = ", lista_pa_exercicio, "\n")

                    nome_arquivo = f"{cnpj}_{nome_empresa}".replace("/", "_").replace(".", "_")  # Substituindo caracteres não permitidos
                    pasta_debitos = os.path.join(os.getcwd(), 'debitos')
                    pasta_codigos = "codigos fiscais"
                    salvar_codigos_em_excel(lista_numeros, lista_pa_exercicio, nome_arquivo, pasta_codigos)
                    processar_empresas_codigos()

                except Exception as e:
                    print('n cliquei em debitos sie F')
                    print(f'Erro{e}')

                #tentando clicar em dívidas ativas
                try: 
                    divida_ativa = WebDriverWait(driver,5).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'list-group-item') and contains(@class, 'collapsed')]//span[text()='Dividas ativas']"))
                    )
                    divida_ativa.click()

                    # Extrair todos os números de todos os <div class="ml-50"> dentro da coluna específica
                    numeros = driver.find_elements(By.XPATH, "//tr//td[@aria-colindex='1']//div[@class='ml-50']")
                    lista_numeros = [numero.text for numero in numeros]

        
                    # Gerar nome do arquivo com CNPJ e nome da empresa
                    nome_arquivo = f"{cnpj}_{nome_empresa}".replace("/", "_").replace(".", "_")  # Substituindo caracteres não permitidos

                    pasta_destino = "dividas ativas"
                    if not os.path.exists(pasta_destino):
                        os.makedirs(pasta_destino)
                    
                    salvar_numeros_em_excel(lista_numeros, nome_arquivo, pasta_destino)
                    sleep(1)
                    processar_excel_e_abrir_pdf()
                    #PEGAR os numeros da dívida, passar pra um excel, ai a partir disso ir no pdf e extrair
                    #Todos que estiverem com ''Pendencia - inscrição'', situação ''Ativa em cobrança'' ou ''Ativa a ser cobrada'', precisa colocar pois são pendências em divida ativa que não foram negociadas ainda
                    #Poderia colocar na mensagem algo como: Pendência em Inscrição em dívida ativa na Procuradoria-Geral da Fazenda Nacional:- colocar os números das inscrições e data que foi inscrito (obs: quando estiver parcelamento rescindido não aparecerá data da inscrição)
                    #Os que estiverem em ''Inscrição com Exigibilidade Suspensa'' E ''Parcelamento com Exigibilidade Suspensa'' não precisa informar nada, pois as vidas já estão negociadas e parceladas

                except Exception as e:
                    print(f"Erro ao clicar em 'Dividas Ativas' ou extrair os números: {e}")

                #tentando clicar em processo fiscal sief
                try: 
                    processo_sief = WebDriverWait(driver,2).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'list-group-item') and contains(@class, 'collapsed')]//span[text()='Processo Fiscal (Sief)']"))
                    )
                    processo_sief.click()

                    # Extrair todos os números de todos os <div class="ml-50"> dentro da coluna específica
                    processos = driver.find_elements(By.XPATH, "//tr//td[@aria-colindex='1']//div[@class='ml-50']")
                    lista_processos = [processo.text for processo in processos]

                    # Gerar nome do arquivo com CNPJ e nome da empresa
                    nome_arquivo = f"{cnpj}_{nome_empresa}".replace("/", "_").replace(".", "_")  # Substituindo caracteres não permitidos

                    pasta_destino_sief = "processos sief"
                    if not os.path.exists(pasta_destino_sief):
                        os.makedirs(pasta_destino_sief)
                    
                    salvar_processos_em_excel(lista_processos, nome_arquivo, pasta_destino_sief, nome_empresa)
                    sleep(1)
                    #PEGAR os numeros da dívida, passar pra um excel, ai a partir disso ir no pdf e extrair
                    #Todos que estiverem com ''Pendencia - inscrição'', situação ''Ativa em cobrança'' ou ''Ativa a ser cobrada'', precisa colocar pois são pendências em divida ativa que não foram negociadas ainda
                    #Poderia colocar na mensagem algo como: Pendência em Inscrição em dívida ativa na Procuradoria-Geral da Fazenda Nacional:- colocar os números das inscrições e data que foi inscrito (obs: quando estiver parcelamento rescindido não aparecerá data da inscrição)
                    #Os que estiverem em ''Inscrição com Exigibilidade Suspensa'' E ''Parcelamento com Exigibilidade Suspensa'' não precisa informar nada, pois as vidas já estão negociadas e parceladas

                except Exception as e:
                    print(f"Erro ao clicar em 'Dividas Ativas' ou extrair os números: {e}")
                
                

                pyautogui.press('esc')
                sleep(2)

                

                if i + 1 == len(linhas):
                    try:
                        sleep(1)
                        pula_pagina = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Go to next page']"))
                        )
                        pula_pagina.click()
                        sleep(3)
                        print("Avançando para a próxima página...")
                        linhas = driver.find_elements(By.CSS_SELECTOR, "tr.clickable")
                        print("Numero de linhas:",len(linhas))
                    except Exception as e:
                        print("Processando mensagens finais...")
                        criar_msgs_codigos(diretorio_codigos, tabela_depto_pessoal, tabela_fiscal, caminho_saida="mensagens.xlsx")
                        criar_msgs(caminho_saida="mensagens.xlsx")
                        criar_msgs_processos_sief(caminho_saida="mensagens.xlsx", diretorio_processos_sief=diretorio_processos_sief)
                        criar_msg_final()
                        print(f"Erro ao clicar no botão de avançar página: {e}")
                        driver.quit()
        
        
                        
        pasta_debitos = os.path.join(os.getcwd(), 'debitos')

from PIL import Image, ImageTk


def atualizar_informacoes(frame):
    """
    Simula a atualização de informações dinâmicas no frame central.
    """
    for widget in frame.winfo_children():
        widget.destroy()  # Limpa o conteúdo anterior

    info_label = ctk.CTkLabel(
        frame,
        text="Atualizando informações... 🚀",
        font=("Arial", 16),
        anchor="center",
    )
    info_label.pack(pady=10)

    # Exemplo de atualização futura
    info_atualizada = ctk.CTkLabel(
        frame,
        text="Extração de débitos em andamento...",
        font=("Arial", 14),
        anchor="center",
    )
    info_atualizada.pack(pady=5)

import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
import schedule
import time
import threading


# Função para configurar o horário
def agendar_robo():
    horario = horario_entry.get()  # Pegando o valor do campo de entrada
    try:
        # Validar o formato de hora
        horario_formatado = datetime.strptime(horario, "%H:%M")
        
        # Agendar a execução do login para o horário escolhido
        horario_str = horario_formatado.strftime("%H:%M")
        schedule.every().day.at(horario_str).do(login)  # Executa a função login todos os dias nesse horário

        messagebox.showinfo("Horário agendado", f"O robô será executado às {horario_formatado.strftime('%H:%M')}")
        
        # Iniciar o agendamento em uma thread separada para não bloquear a interface
        threading.Thread(target=run_schedule).start()

    except ValueError:
        messagebox.showerror("Erro", "Formato de hora inválido! Use o formato HH:MM.")

# Função para rodar o schedule em uma thread separada
def run_schedule():
    while True:
        schedule.run_pending()  # Verifica se alguma tarefa agendada precisa ser executada
        time.sleep(1)  # Espera 1 segundo antes de verificar novamente


# Configuração inicial da interface
ctk.set_appearance_mode("dark")  # "System", "Light" ou "Dark"
ctk.set_default_color_theme("blue")  # "blue", "green" ou "dark-blue"

# Janela principal
app = ctk.CTk()
app.geometry("900x600")
app.title("Sistema de Extração de Débitos - Primor")

# Menu lateral
menu_frame = ctk.CTkFrame(app, width=200, corner_radius=0)
menu_frame.pack(side="left", fill="y")

menu_label = ctk.CTkLabel(
    menu_frame,
    text="Menu",
    font=("Arial", 20, "bold"),
    anchor="w",
)
menu_label.pack(pady=(20, 10), padx=10, anchor="w")

# Botão no menu
extrair_button = ctk.CTkButton(
    menu_frame,
    text="Extrair débitos",
    command=login,
    font=("Arial", 16, "bold"),
)
extrair_button.pack(pady=20, padx=10)

# Área principal
main_frame = ctk.CTkFrame(app, corner_radius=10)
main_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

main_label = ctk.CTkLabel(
    main_frame,
    text="Bem-vindo ao PrimorFiscal Messenger",
    font=("Arial", 24, "bold"),
    anchor="center",
)
main_label.pack(pady=20)

# Seleção de horário
horario_label = ctk.CTkLabel(
    app, text="Escolha o horário para o robô rodar (formato HH:MM):", font=("Arial", 16)
)
horario_label.pack(pady=10)

# Campo de entrada para o horário
horario_entry = ctk.CTkEntry(app, width=150, font=("Arial", 14))
horario_entry.insert(0, "09:00")  # Definir o horário padrão
horario_entry.pack(pady=10)

# Botão para agendar
agendar_button = ctk.CTkButton(
    main_frame,
    text="Agendar Robô",
    command=agendar_robo,
    font=("Arial", 16, "bold"),
)
agendar_button.pack(pady=10)

# Adicionar imagens e mensagem de parceria
try:
    from PIL import Image, ImageTk  # Certifique-se de que o Pillow está instalado
    primor_image = ImageTk.PhotoImage(Image.open("imgs/primor.png").resize((200, 120)))
    luiz_image = ImageTk.PhotoImage(Image.open("imgs/luiz.png").resize((200, 120)))

    images_frame = ctk.CTkFrame(main_frame)
    images_frame.pack(pady=20)

    primor_label = ctk.CTkLabel(images_frame, image=primor_image, text="")
    primor_label.grid(row=0, column=0, padx=10)

    luiz_label = ctk.CTkLabel(images_frame, image=luiz_image, text="")
    luiz_label.grid(row=0, column=1, padx=10)

    # Texto de parceria
    partnership_label = ctk.CTkLabel(
        main_frame,
        text="Uma parceria entre Primor e Luiz Fernando Hillebrande",
        font=("Arial", 16, "italic"),
        anchor="center",
    )
    partnership_label.pack(pady=(10, 20))

except Exception as e:
    print(f"Erro ao carregar imagens: {e}")

# Frame dinâmico para atualizações
info_frame = ctk.CTkFrame(main_frame, height=200)
info_frame.pack(fill="both", expand=True, padx=10, pady=(10, 20))

# Atualizar as informações no frame (exemplo)
# atualizar_informacoes(info_frame)  # Descomente se necessário

# Footer (centralizado)
footer_frame = ctk.CTkFrame(app, height=50, corner_radius=0)
footer_frame.pack(side="bottom", fill="x")

footer_label = ctk.CTkLabel(
    footer_frame,
    text="Luiz Fernando Hillebrande",
    font=("Arial", 14),
    anchor="center",
)
footer_label.pack(pady=10)

# Iniciar a interface
app.mainloop()
