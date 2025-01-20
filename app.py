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

def login():
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

    sleep(2)

    logar_email = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='email']"))
    )
    logar_email.send_keys('luiz.logika@gmail.com')

    logar_senha = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='senhaInput']"))
    )
    logar_senha.send_keys('Luiz123')

    logar = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@type='submit']"))
    )
    logar.click()

    sleep(3)
    pyautogui.press('esc')

    '''
    AQUI ACABA O LOGIN
    '''


    '''
    COMECANDO A BAIXAR OS PDF
    '''

    #TENHO QUE ENTRAR NESSA ABA AGORA, https://app.monitorcontabil.com.br/situacao-fiscal/visualizar?busca=

    driver.get("https://app.monitorcontabil.com.br/situacao-fiscal/visualizar?busca=")

    #atualizar_lote = WebDriverWait(driver,5).until(
        #EC.element_to_be_clickable((By.XPATH,"//button[@title='A atualização busca a Situação fiscal na data atual para todas as empresas selecionadas. #Cada atualização consumirá um crédito do saldo em conta.']"))
    #)

    #atualizar_lote.click()

    baixar_relatorios = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@title='O download será feito conforme os filtros atualmente selecionados']"))
    )
    baixar_relatorios.click()

    baixar_popup = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@class='btn btn-sm btn-success']"))
    )
    baixar_popup.click()

    sleep(8)
    pyautogui.press('F5')

    baixar_nuvem = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@class='nav-link dropdown-toggle' and @id='__BVID__197__BV_toggle_']"))
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


    zip_file = None
    for file in os.listdir(download_folder):
        if file.endswith('.zip'):
            zip_file_path = os.path.join(download_folder, file)
            if not zip_file or os.path.getmtime(zip_file_path) > os.path.getmtime(zip_file):
                zip_file = zip_file_path

    if zip_file:
        # Descompactar o arquivo ZIP
        try:
            with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                zip_ref.extractall(download_folder)
            print(f"Arquivo ZIP {zip_file} descompactado.")

            # Excluir o arquivo ZIP
            os.remove(zip_file)
            print(f"Arquivo ZIP {zip_file} excluído.")
        except Exception as e:
            print(f"Erro ao descompactar ou excluir o arquivo ZIP: {e}")
    else:
        print("Nenhum arquivo ZIP encontrado na pasta 'debitos'.")


    sleep(3)
    pasta_debitos = os.path.join(os.getcwd(), 'debitos')
    nomes_empresas = extrair_nome_empresa(pasta_debitos)
    salvar_nome_empresa_excel(nomes_empresas, 'nomes_empresas.xlsx')

        
    sleep(2)
    driver.quit()

def extrair_nome_empresa(pasta_debitos):
    # Lista para armazenar os nomes das empresas
    nomes_empresas = []

    # Percorrer todos os arquivos na pasta
    for arquivo in os.listdir(pasta_debitos):
        if arquivo.endswith(".pdf"):  # Verifica se é um arquivo PDF
            # Nome do arquivo: situacao_fiscal--CNPJ-NOME DA EMPRESA.pdf
            # Padrão de regex para extrair o nome da empresa
            match = re.match(r'situacao_fiscal--\d{14}-(.*)\.pdf', arquivo)
            if match:
                # Extrai o nome da empresa
                nome_empresa = match.group(1)
                nomes_empresas.append(nome_empresa)

    return nomes_empresas

# Função para salvar os nomes em um arquivo Excel
def salvar_nome_empresa_excel(nomes_empresas, caminho_arquivo_excel):
    # Criar um DataFrame com os nomes
    df = pd.DataFrame(nomes_empresas, columns=["Nome da Empresa"])
    # Salvar no Excel
    df.to_excel(caminho_arquivo_excel, index=False)

# Caminho da pasta onde os PDFs foram descompactados
pasta_debitos = os.path.join(os.getcwd(), 'debitos')


def carregar_codigos_fiscais(caminho_arquivo_excel):
    # Carregar os dados dos dois sheets
    df_depto = pd.read_excel(caminho_arquivo_excel, sheet_name="Depto Pessoal")
    df_fiscal = pd.read_excel(caminho_arquivo_excel, sheet_name="Fiscal")

    # Concatenar os dois dataframes
    df = pd.concat([df_depto, df_fiscal])

    # Converter para um dicionário {codigo: descricao}
    codigos_fiscais = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
    
    return codigos_fiscais

# Função para extrair o texto de todos os PDFs na pasta
def extrair_texto_pdfs(pasta_debitos):
    textos_pdfs = {}
    
    for arquivo in os.listdir(pasta_debitos):
        if arquivo.endswith(".pdf"):
            caminho_pdf = os.path.join(pasta_debitos, arquivo)
            with pdfplumber.open(caminho_pdf) as pdf:
                texto_completo = ""
                for pagina in pdf.pages:
                    texto_completo += pagina.extract_text()
                textos_pdfs[arquivo] = texto_completo
    
    return textos_pdfs

# Função para buscar os códigos fiscais nos textos extraídos dos PDFs
def buscar_codigos_fiscais(textos_pdfs, codigos_fiscais):
    resultados = []

    # Loop por cada PDF e seu texto
    for nome_pdf, texto in textos_pdfs.items():
        # Procurar pelos códigos fiscais no texto do PDF
        for codigo, descricao in codigos_fiscais.items():
            codigo_str = str(codigo)
            # Regex para encontrar o código e seu valor associado
            pattern = r"(" + re.escape(codigo_str) + r")(.*?)(\d+[\.,]?\d{2})"
            matches = re.findall(pattern, texto)

            for match in matches:
                codigo_fiscal = match[0]
                descricao_encontrada = descricao
                valor_total = match[2].replace(",", ".")
                
                # Salvar os resultados com a empresa e o valor
                resultado = {
                    "Nome da Empresa": nome_pdf.split('--')[1].split('-')[1],  # Extrair nome da empresa do nome do arquivo
                    "Código Fiscal": codigo_fiscal,
                    "Descrição": descricao_encontrada,
                    "Valor Total": valor_total
                }
                resultados.append(resultado)
    
    return resultados

# Função para salvar os resultados em um arquivo Excel
def salvar_resultados_excel(resultados, caminho_arquivo_excel):
    df_resultados = pd.DataFrame(resultados)
    df_resultados.to_excel(caminho_arquivo_excel, index=False)

# Caminho do arquivo Excel com os códigos fiscais
caminho_tabela_codigos = 'TABELASCDIGOSDERECEITA.xlsx'

# Caminho da pasta 'debitos' onde estão os PDFs
pasta_debitos = os.path.join(os.getcwd(), 'debitos')

# Carregar os códigos fiscais
codigos_fiscais = carregar_codigos_fiscais(caminho_tabela_codigos)

# Extrair o texto de todos os PDFs
textos_pdfs = extrair_texto_pdfs(pasta_debitos)

# Buscar os códigos fiscais nos textos
resultados = buscar_codigos_fiscais(textos_pdfs, codigos_fiscais)

# Salvar os resultados no Excel
salvar_resultados_excel(resultados, 'resultados_fiscais.xlsx')

print("Resultados fiscais extraídos e salvos com sucesso!")





