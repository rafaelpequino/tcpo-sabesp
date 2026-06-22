import os
from turtle import delay
import pandas as pd
from time import sleep
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from utils.auth import login, acessar_banco, encerrar
import re

load_dotenv()

# Configuração do navegador
options = Options()
options.headless = False

navegador = webdriver.Chrome(options=options)

url = "https://tcpoweb.pini.com.br/home/home.aspx"

isLoggedIn = False


def exportar_insumos():
    insumos = {
        "Categoria": [],
        "Base": [],
        "Código": [],
        "Descrição": [],
        "Unidade": [],
        "Tipo": [],
        "Data Preço": [],
        "Preço": []
    }

    categorias = {
        "Materiais": "ctl00_MainContent_PiniTreeViewt304",
        "Mão de obra": "ctl00_MainContent_PiniTreeViewt305",
        "Mão de obra empreitada": "ctl00_MainContent_PiniTreeViewt306",
        "Serviços terceirizados": "ctl00_MainContent_PiniTreeViewt307",
        "Equipamentos - Aquisição": "ctl00_MainContent_PiniTreeViewt308",
        "Equipamentos - Locação": "ctl00_MainContent_PiniTreeViewt309"
    }

    try:
        print("Iniciando leitura de insumos...")

        navegador.find_element(By.ID, "ctl00_MainContent_PiniTreeViewt303").click()
        sleep(2)

        for categoria, id_elemento in categorias.items():
            btnCategoria = navegador.find_element(By.ID, id_elemento)
            btnCategoria.click()
            sleep(5)

            # Extrai o número total de páginas da categoria
            elemento = navegador.find_element(By.ID, "ctl00_MainContent_btnServicos")
            texto = elemento.get_attribute("value")
            paginas = int(re.search(r'Página \d+ de (\d+)', texto).group(1))

            print(f"Lendo categoria: {categoria}")
            sleep(2)

            # Considera que página 1 já está aberta
            for pagina in range(1, paginas + 1):
                print(f"Página {pagina}/{paginas}")
                sleep(1)

                if pagina == paginas:
                    break

                links = navegador.find_elements(
                    By.CSS_SELECTOR,
                    "#ctl00_MainContent_gvServicos tr.gridPager td table tbody tr td a"
                )

                proxima_pagina = str(pagina + 1)

                # tenta achar a próxima página
                link_proxima = None

                for link in links:
                    if link.text.strip() == proxima_pagina:
                        link_proxima = link
                        break

                if link_proxima:
                    link_proxima.click()

                else:
                    # só chega aqui quando a próxima página não está no bloco atual
                    botoes_reticencias = [
                        link for link in links
                        if link.text.strip() == "..."
                    ]

                    if botoes_reticencias:
                        print("Avançando bloco de páginas...")

                        # último ... = avançar
                        botoes_reticencias[-1].click()

                    else:
                        print("Não encontrei próxima página")
                        break

                sleep(2)







    except Exception as e:
        print(f"Erro ao exportar insumos: {e}")


# ===========================
# Fluxo da Aplicação
# ===========================

print("Iniciando aplicação...")

isLoggedIn = login(navegador, url)
if isLoggedIn:
    acessar_banco(navegador)
    sleep(5)
    exportar_insumos()
else:
    print("Aplicação não pode continuar sem login.")


encerrar(navegador)