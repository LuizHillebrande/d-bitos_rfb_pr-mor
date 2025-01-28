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
import undetected_chromedriver as uc
import openpyxl as opx

wb_fgts = opx.load_workbook('EMPRESAS FGTS.xlsx')

# Para acessar a planilha
sheet_wb = wb_fgts['Página1']

imagem_alvo = r"certificado_esperado.png"

# Intervalo entre os cliques de tecla
intervalo = 0.5

def localizar_imagem_na_tela(imagem, confidence=0.8):
    """
    Tenta localizar a imagem na tela.
    :param imagem: Caminho da imagem a ser localizada.
    :param confidence: Nível de confiança para correspondência da imagem.
    :return: Coordenadas da imagem encontrada ou None.
    """
    try:
        # Localiza a imagem na tela
        localizacao = pyautogui.locateOnScreen(imagem, confidence=confidence)
        return localizacao
    except Exception as e:
        print(f"Erro ao localizar imagem: {e}")
        return None

def pressionar_ate_encontrar(imagem, intervalo=0.5):
    """
    Pressiona a seta para baixo até encontrar a imagem na tela.
    :param imagem: Caminho da imagem a ser localizada.
    :param intervalo: Intervalo entre as teclas pressionadas.
    """
    while True:
        localizacao = localizar_imagem_na_tela(imagem)
        if localizacao:
            print(f"Imagem encontrada nas coordenadas: {localizacao}")
            pyautogui.press('enter')
            sleep(3)
            break
        else:
            print("Imagem não encontrada. Pressionando seta para baixo...")
            pyautogui.press('down')
            time.sleep(intervalo)

def pegar_debitos_fgts():
    driver = uc.Chrome()
    driver.get("https://fgtsdigital.sistema.gov.br/portal/login")
    driver.maximize_window()

    sleep(2)

    entry_gov = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@class='br-button is-primary entrar']"))
    )
    entry_gov.click()


    entry_certificate = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@id='login-certificate']"))
    )
    entry_certificate.click()
    sleep(3)

    pyautogui.click(738,191, duration=2)
    sleep(3)

    pressionar_ate_encontrar(imagem_alvo, intervalo)

    sleep(5)

    definir = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@class='br-button is-primary']"))
    )
    definir.click

    for linha in sheet_wb.iter_rows(min_row=2, max_row=500):
        razao_social = linha[1].value
        cnpj = linha[2].value

        try: 
            trocar_perfil = WebDriverWait(driver,2).until(
            EC.element_to_be_clickable((By.XPATH,"//button[@class=' br-button secondary botao-barra-perfil']"))
            )
            trocar_perfil.click
        except:
            print('N tinha trocar perfil')


        dropdown_perfil = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH,"//div[@role='combobox']"))
        )
        dropdown_perfil.click()

        procurador_option = driver.find_element(By.XPATH, "//span[@class='ng-option-label' and text()='Procurador']")

        # Clicando no item
        ActionChains(driver).move_to_element(procurador_option).click().perform()

        input_cnpj = WebDriverWait(driver,5).until(
            EC.element_to_be_clickable((By.XPATH,"//input[@class='brx-input medium ng-untouched ng-pristine ng-invalid']"))
        )
        input_cnpj.click()
        input_cnpj.send_keys(cnpj)



        sleep(5)
        driver.quit()
pegar_debitos_fgts()
