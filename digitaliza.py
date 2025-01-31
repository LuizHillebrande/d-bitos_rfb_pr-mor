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
from tkinter import messagebox

def enviar_mensagens():
    driver = webdriver.Chrome()
    driver.get('https://app.digiliza.com.br/login')

    input_email = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='email']"))
    )
    input_email.click()
    input_email.send_keys('legal@contabilprimor.com.br')

    input_senha = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//input[@id='password']"))
    )
    input_senha.click()
    input_senha.send_keys('q7ne5k0la0VJ')

    enter = WebDriverWait(driver,5).until(
        EC.element_to_be_clickable((By.XPATH,"//button[@type='submit']"))
    )
    enter.click()

    sleep(5)
    messagebox.showinfo('Sucesso!', 'Mensagens enviadas com sucesso')
    driver.quit()

enviar_mensagens()
