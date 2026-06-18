import os
from time import sleep

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


load_dotenv()

# Configuração do navegador
options = Options()
options.headless = False

navegador = webdriver.Chrome(options=options)

url = "https://tcpoweb.pini.com.br/home/home.aspx"

isLoggedIn = False


def login():
    global isLoggedIn

    try:
        print("Acessando página...")
        navegador.get(url)
        sleep(1)

        # Usuário
        print("Preenchendo campo de usuário...")
        inputUser = navegador.find_element(By.ID, "ctl00_header1_txtUsuario")
        inputUser.send_keys(os.getenv("USER_INPUT"))

        # Senha
        print("Preenchendo campo de senha...")
        inputPass = navegador.find_element(By.ID, "ctl00_header1_txtSenha")
        inputPass.send_keys(os.getenv("PASS_INPUT"))

        sleep(1)

        # Entrar
        print("Clicando em entrar...")
        btnEntrar = navegador.find_element(By.ID, "ctl00_header1_btnAcessar")
        btnEntrar.click()

        # Verifica se apareceu mensagem de aviso
        try:
            aviso = WebDriverWait(navegador, 2).until(
                EC.visibility_of_element_located((By.ID, "dialogboxhead"))
            )

            if aviso.text.strip() == "Aviso":
                print("Houve um problema no login!")
                isLoggedIn = False
                return

        except:
            # Nenhum aviso apareceu
            pass

        print("Login realizado com sucesso!")
        isLoggedIn = True

    except Exception as e:
        print(f"Erro ao acessar o site: {e}")
        isLoggedIn = False


def executar():
    """
    Coloque aqui as ações que devem ocorrer
    depois que estiver logado.
    """

    print("Executando rotina principal...")
    sleep(2)
    encerrar()


def encerrar():
    global isLoggedIn

    input("Pressione ENTER para encerrar...")
    try:
        print("Encerrando sessão...")
        if isLoggedIn:
            try:
                btnSair = navegador.find_element(By.ID, "ctl00_header2_Image4")
                btnSair.click()
                sleep(1)

            except Exception:
                print("Botão de sair não encontrado.")

        navegador.quit()
        print("Aplicação finalizada.")

    except Exception as e:
        print(f"Erro ao encerrar a sessão: {e}")


# ===========================
# Fluxo da Aplicação
# ===========================

print("Iniciando aplicação...")

login()

if isLoggedIn:
    executar()
else:
    print("Aplicação não pode continuar sem login.")


encerrar()