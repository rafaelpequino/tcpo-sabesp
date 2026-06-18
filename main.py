import os
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from time import sleep

load_dotenv()

options = Options()
options.headless = False

navegador = webdriver.Chrome(options=options)

url = "https://tcpoweb.pini.com.br/home/home.aspx"

# Acessar a página
navegador.get(url)
sleep(1)

# Preencher o campo de usuário
inputUser = navegador.find_element(By.ID, "ctl00_header1_txtUsuario")
inputUser.send_keys(os.getenv("USER_INPUT"))
sleep(1)

# Preencher o campo de senha
inputPass = navegador.find_element(By.ID, "ctl00_header1_txtSenha")
inputPass.send_keys(os.getenv("PASS_INPUT"))
sleep(1)

# Clicar em "Entrar"
btnEntrar = navegador.find_element(By.ID, "ctl00_header1_btnAcessar")
btnEntrar.click()
sleep(1)