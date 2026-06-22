import os
from time import sleep
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from utils.auth import login, acessar_banco, encerrar

load_dotenv()

# Configuração do navegador
options = Options()
options.headless = False

navegador = webdriver.Chrome(options=options)

url = "https://tcpoweb.pini.com.br/home/home.aspx"

isLoggedIn = False

# ===========================
# Fluxo da Aplicação
# ===========================

print("Iniciando aplicação...")

isLoggedIn = login(navegador, url)

if isLoggedIn:
    acessar_banco(navegador)
    sleep(15)
else:
    print("Aplicação não pode continuar sem login.")


encerrar(navegador, isLoggedIn)