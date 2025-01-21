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

def consultar_pdf_da_empresa(nome_empresa):
    # Caminho da pasta onde os PDFs foram baixados/descompactados
    pasta_debitos = os.path.join(os.getcwd(), 'debitos')

    # Buscar arquivos PDF na pasta de débitos que contenham o nome da empresa (sem LTDA)
    arquivos_pdf = [f for f in os.listdir(pasta_debitos) if f.endswith('.pdf') and nome_empresa in f]

    if arquivos_pdf:
        # Se encontrar arquivos PDF que correspondem ao nome da empresa
        for pdf_file in arquivos_pdf:
            caminho_pdf = os.path.join(pasta_debitos, pdf_file)
            print(f"PDF encontrado: {pdf_file}")
            
            # Aqui você pode usar uma biblioteca para abrir o PDF, por exemplo, pdfplumber
            abrir_pdf(caminho_pdf)
    else:
        print(f"Nenhum PDF encontrado para a empresa {nome_empresa}.")

def abrir_pdf(caminho_pdf):
    try:
        # Usando pdfplumber para abrir e extrair informações do PDF
        import pdfplumber
        with pdfplumber.open(caminho_pdf) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                print(f"Conteúdo da página {page_num + 1}:")
                print(text[:500])  # Exibe os primeiros 500 caracteres da página
                print("=" * 50)  # Separador para facilitar a leitura do conteúdo
    except Exception as e:
        print(f"Erro ao abrir o arquivo PDF {caminho_pdf}: {e}")

def processar_excel_e_abrir_pdf():
    # Caminho do diretório onde os arquivos Excel estão armazenados
    pasta_destino = "dividas ativas"
    
    # Iterando pelos arquivos na pasta 'dividas ativas' para pegar o nome do Excel
    for excel_file in os.listdir(pasta_destino):
        if excel_file.endswith('.xlsx'):
            # Extrair o nome da empresa a partir do nome do arquivo Excel (sem a extensão .xlsx)
            nome_empresa = os.path.splitext(excel_file)[0]
            
            # Remover " LTDA" do nome da empresa (para corresponder ao nome no PDF)
            nome_empresa = nome_empresa.replace(" LTDA", "")
            
            print(f"Nome da empresa extraído do Excel: {nome_empresa}")
            
            # Agora consultar o PDF correspondente
            consultar_pdf_da_empresa(nome_empresa)
            
def salvar_numeros_em_excel(lista_numeros, nome_arquivo, pasta_destino):
    # Salvar a lista de números em um arquivo Excel
    df = pd.DataFrame(lista_numeros, columns=["Inscrição da Dívida"])
    caminho_arquivo = os.path.join(pasta_destino, f"{nome_arquivo}.xlsx")
    df.to_excel(caminho_arquivo, index=False)
    print(f"Arquivo salvo em {caminho_arquivo}")


def descompactar_arquivo_zip(download_folder):
    # Descompactar o arquivo ZIP
    zip_file = None
    for file in os.listdir(download_folder):
        if file.endswith('.zip'):
            zip_file_path = os.path.join(download_folder, file)
            if not zip_file or os.path.getmtime(zip_file_path) > os.path.getmtime(zip_file):
                zip_file = zip_file_path

    if zip_file:
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
    descompactar_arquivo_zip(download_folder)
    pyautogui.press('Esc')
    sleep(2)

    lupa = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@class='btn btn btn-none rounded-pill m-0 icone-acao p-0 btn-none btn-none'][2]"))
    )

    lupa.click()

    #tentando clicar em dívidas ativas
    try: 
        divida_ativa = WebDriverWait(driver,5).until(
            EC.element_to_be_clickable((By.XPATH,"//div[@class='list-group-item active collapsed']"))
        )
        divida_ativa.click()

        # Extrair todos os números de todos os <div class="ml-50"> dentro da coluna específica
        numeros = driver.find_elements(By.XPATH, "//tr//td[@aria-colindex='1']//div[@class='ml-50']")
        lista_numeros = [numero.text for numero in numeros]

        empresa_element = driver.find_element(By.XPATH, "//h5[@id='pendencia-fiscal___BV_modal_title_']")
        nome_empresa_completo = empresa_element.text
        # Remover a parte "Pendência da situação fiscal - " do nome
        nome_empresa = nome_empresa_completo.replace("Pendência da situação fiscal - ", "").strip()
        nome_arquivo = re.sub(r'[\\/*?:"<>|]', "", nome_empresa)

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
    
    #tentando clicar em débitos(sief)
    try:    
        debitos_sief = WebDriverWait(driver,5).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'list-group-item') and contains(@class, 'collapsed')]//span[text()='Débito (Sief)']"))
        )
        debitos_sief.click()
        print('CLICADO EM DEBITOS SIEF')
    except Exception as i:
        print('n cliquei em debitos sie F')
        print(f'Erro{i}')


    sleep(3)
    pasta_debitos = os.path.join(os.getcwd(), 'debitos')

    sleep(2)
    driver.quit()

login()





