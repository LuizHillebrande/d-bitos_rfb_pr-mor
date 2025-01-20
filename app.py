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

def login():
    download_folder = os.getcwd()  # Defina o caminho completo da pasta debitos_rfb

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
    baixar_definitivamente = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@class='media mr-1']//div[@class='col-2']//svg[contains(@class, 'feather-download')]"))
    )
    baixar_definitivamente.click()




    sleep(5)
    driver.quit()

login()



