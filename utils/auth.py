import os
from time import sleep
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC


def login(navegador, url):
    try:
        print("Acessando página...")
        navegador.get(url)
        sleep(1)

        print("Preenchendo campo de usuário...")
        inputUser = navegador.find_element(By.ID, "ctl00_header1_txtUsuario")
        inputUser.send_keys(os.getenv("USER_INPUT"))

        print("Preenchendo campo de senha...")
        inputPass = navegador.find_element(By.ID, "ctl00_header1_txtSenha")
        inputPass.send_keys(os.getenv("PASS_INPUT"))

        sleep(1)

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
                return False
                
        except:
            # Nenhum aviso apareceu
            pass

        print("Login realizado com sucesso!")
        return True

    except Exception as e:
        print(f"Erro ao acessar o site: {e}")
        return False



def acessar_banco(navegador):
    sleep(1)
    try:
        print("Acessando Bases, composições e ferramentas...")
        btnEntrar = navegador.find_element(By.ID, "ctl00_MainContent_gvFuncionalidades_ctl02_btnFuncaoNome")
        btnEntrar.click()

        sleep(3)

        print("Abrindo menu de filtros...")
        btnEntrar = navegador.find_element(By.ID, "ctl00_MainContent_imgBtnBuscaAvancada")
        btnEntrar.click()

        sleep(5)
        
        print("Definindo filtros...")
        btnEntrar = navegador.find_element(By.ID, "ctl00_MainContent_chkBuscaBaseSelecionada")
        btnEntrar.click()

        print("Selecionando banco TCPO PINI...")
        selectEAP = Select(
            navegador.find_element(By.ID, "ctl00_MainContent_cboEAP")
        )
        selectEAP.select_by_value("TCPO_PINI|1|")

        print("Verificando banco selecionado...")
        bancoSelecionado = selectEAP.first_selected_option.get_attribute("value")
        if bancoSelecionado == "TCPO_PINI|1|":
            print("TCPO PINI selecionado corretamente.")
        else:
            print(f"Erro: valor atual é {bancoSelecionado}")
            encerrar(navegador, True)

    
    except Exception as e:
        print(f"Erro ao acessar o banco: {e}")
        encerrar(navegador, True)



def encerrar(navegador, isLoggedIn=True):
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
