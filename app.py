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


def extrair_nome_empresa(nome_arquivo):
    """
    Extrai o nome da empresa do nome do arquivo PDF.
    Remove o prefixo 'situacao_fiscal--CNPJ-' e qualquer código no final.
    """
    nome_limpo = re.sub(r"situacao_fiscal--\d{14}-", "", nome_arquivo)  # Remove CNPJ e prefixo
    nome_limpo = re.sub(r"_[0-9]+\.pdf$", "", nome_limpo)  # Remove código final (se existir)
    
    return nome_limpo.strip()

def renomear_pdfs(pasta):
    """
    Itera sobre todos os PDFs da pasta e renomeia removendo prefixos e códigos desnecessários.
    """
    for arquivo in os.listdir(pasta):
        if arquivo.endswith(".pdf"):
            caminho_antigo = os.path.join(pasta, arquivo)
            nome_extraido = extrair_nome_empresa(arquivo)
            novo_nome = f"{nome_extraido}.pdf"
            caminho_novo = os.path.join(pasta, novo_nome)
            
            # Renomeia o arquivo
            os.rename(caminho_antigo, caminho_novo)
            print(f"Renomeado: {arquivo} → {novo_nome}")

# Defina a pasta onde estão os PDFs
pasta_pdfs = "debitos"

def consultar_pdf_da_empresa(nome_empresa, numeros_procurados):
    # Caminho da pasta onde os PDFs foram baixados/descompactados
    pasta_debitos = os.path.join(os.getcwd(), 'debitos')

    # Listar todos os arquivos PDF na pasta
    arquivos_pdf = [f for f in os.listdir(pasta_debitos) if f.endswith('.pdf')]

    # Tentar encontrar o PDF mais próximo usando fuzzy matching
    nome_pdf_proximo, score = process.extractOne(nome_empresa, arquivos_pdf)

    if score >= 80:  # Define um limite mínimo de similaridade (ajustável)
        caminho_pdf = os.path.join(pasta_debitos, nome_pdf_proximo)
        print(f"🔍 PDF encontrado: {nome_pdf_proximo} (Similaridade: {score}%)")

        # Aqui você pode usar uma biblioteca para abrir o PDF, por exemplo, pdfplumber
        abrir_pdf(caminho_pdf, numeros_procurados, nome_empresa)
    else:
        print(f"⚠️ Nenhum PDF encontrado para a empresa '{nome_empresa}'.")

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

    # Tentar encontrar o PDF mais próximo usando fuzzy matching
    nome_pdf_proximo, score = process.extractOne(nome_empresa, arquivos_pdf)

    if score >= 80:  # Define um limite mínimo de similaridade
        caminho_pdf = os.path.join(pasta_debitos, nome_pdf_proximo)
        print(f"🔍 PDF encontrado: {nome_pdf_proximo} (Similaridade: {score}%)")

        # Chamar a função para processar o PDF
        abrir_pdf_codigos(caminho_pdf, numeros_procurados, nome_empresa)
    else:
        print(f"⚠️ Nenhum PDF encontrado para a empresa '{nome_empresa}'.")

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

    sleep(6)
    pyautogui.press('F5')
    sleep(2)

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
    sleep(2)
    renomear_pdfs(pasta_pdfs)
    pyautogui.press('Esc')
    sleep(2)

    lupas = WebDriverWait(driver,5).until(
        EC.presence_of_all_elements_located((By.XPATH,"//button[@class='btn btn btn-none rounded-pill m-0 icone-acao p-0 btn-none btn-none'][2]"))
    )
    
    for lupa in lupas:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable(lupa)).click()


        #tentando clicar em dívidas ativas
        try: 
            divida_ativa = WebDriverWait(driver,5).until(
                EC.element_to_be_clickable((By.XPATH,"//div[@class='list-group-item active collapsed']"))
            )
            divida_ativa.click()

            # Extrair todos os números de todos os <div class="ml-50"> dentro da coluna específica
            numeros = driver.find_elements(By.XPATH, "//tr//td[@aria-colindex='1']//div[@class='ml-50']")
            lista_numeros = [numero.text for numero in numeros]

            #pegando nome da empresa 
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
            numeros = driver.find_elements(By.XPATH, "//tr//td[@aria-colindex='1']//div[@class='ml-50']")
            lista_numeros = [numero.text for numero in numeros]

            print("Códigos fiscais:", lista_numeros, "\n")

            pa_exercicio = driver.find_elements(By.XPATH, "//tr//td[@aria-colindex='2']//div[@class='ml-50']")
            lista_pa_exercicio = [pa.text for pa in pa_exercicio]
            print("Pa - exercício = ", lista_pa_exercicio, "\n")

            empresa_element = driver.find_element(By.XPATH, "//h5[@id='pendencia-fiscal___BV_modal_title_']")
            nome_empresa_completo = empresa_element.text
            # Remover a parte "Pendência da situação fiscal - " do nome
            nome_empresa = nome_empresa_completo.replace("Pendência da situação fiscal - ", "").strip()
            nome_arquivo = re.sub(r'[\\/*?:"<>|]', "", nome_empresa)
            pasta_debitos = os.path.join(os.getcwd(), 'debitos')
            pasta_codigos = "codigos fiscais"
            salvar_codigos_em_excel(lista_numeros, lista_pa_exercicio, nome_arquivo, pasta_codigos)
            processar_empresas_codigos()
            


        except Exception as i:
            print('n cliquei em debitos sie F')
            print(f'Erro{i}')

        pyautogui.press('esc')
        sleep(2)


    sleep(3)
    pasta_debitos = os.path.join(os.getcwd(), 'debitos')

    sleep(2)
    driver.quit()

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

# Adicionar imagens e mensagem de parceria
try:
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
atualizar_informacoes(info_frame)

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
