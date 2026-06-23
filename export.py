import pandas as pd
from time import sleep
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, WebDriverException
from utils.auth import login, acessar_banco, encerrar
import re
import os
from datetime import datetime

load_dotenv()

# Configuração do navegador
options = Options()
options.headless = False
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-extensions")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

url = "https://tcpoweb.pini.com.br/home/home.aspx"

def salvar_excel(dados, status):
    pasta = "arquivos"

    if not os.path.exists(pasta):
        os.makedirs(pasta)

    agora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    nome_arquivo = f"insumos_{agora}_{status}.xlsx"

    caminho = os.path.join(pasta, nome_arquivo)

    cabecalho = [
        'Base',
        'Item',
        'Descrição',
        'Un.',
        'Tipo',
        'Data Preço',
        'Preço'
    ]

    df = pd.DataFrame(dados, columns=cabecalho)
    df.to_excel(caminho, index=False)

    print(f"Arquivo salvo: {caminho}")


def exportar_insumos(navegador):
    wait = WebDriverWait(navegador, 15)
    
    insumos = {
        #"Materiais": "ctl00_MainContent_PiniTreeViewt304",
        "Mão de obra": "ctl00_MainContent_PiniTreeViewt305",
        "Mão de obra empreitada": "ctl00_MainContent_PiniTreeViewt306",
        "Serviços terceirizados": "ctl00_MainContent_PiniTreeViewt307",
        "Equipamentos - Aquisição": "ctl00_MainContent_PiniTreeViewt308",
        "Equipamentos - Locação": "ctl00_MainContent_PiniTreeViewt309"
    }

    dados_excel = []

    try:
        print("Iniciando leitura de insumos...")

        # Desabilita elementos ocultos
        try:
            navegador.execute_script("""
                document.querySelectorAll('#ctl00_MainContent_PiniTreeView *').forEach(function(el) {
                    el.style.display = 'block';
                });
            """)
        except Exception as e:
            print(f"Aviso: Não foi possível modificar a árvore: {e}")
        
        sleep(1)

        for categoria, id_elemento in insumos.items():
            print(f"\n=== Processando Categoria: {categoria} ===")
            try:
                btnCategoria = wait.until(EC.element_to_be_clickable((By.ID, id_elemento)))
                btnCategoria.click()
                sleep(3)

                elemento = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_btnServicos")))
                texto = elemento.get_attribute("value")
                
                # Extrai número de páginas
                match = re.search(r'Página \d+ de (\d+)', texto)
                if not match:
                    print(f"Não foi possível extrair número de páginas. Texto: {texto}")
                    continue
                    
                paginas = int(match.group(1))
                print(f"Total de páginas a processar: {paginas}")
                sleep(1)

                for pagina in range(1, paginas + 1):
                    print(f"\nProcessando página {pagina}/{paginas}")
                    
                    try:
                        indice_tr = 2
                        max_retries = 3

                        while True:
                            try:
                                # Aguarda presença da tabela
                                wait.until(EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                ))

                                trs = navegador.find_elements(
                                    By.CSS_SELECTOR,
                                    "#ctl00_MainContent_gvServicos > tbody > tr"
                                )

                                if indice_tr >= len(trs) - 1:
                                    print(f"Fim da página. Total de linhas: {len(trs)}")
                                    break

                                tr = trs[indice_tr]
                                tds = tr.find_elements(By.TAG_NAME, "td")

                                if len(tds) < 3:
                                    print(f"Linha {indice_tr} não tem dados suficientes. Pulando...")
                                    indice_tr += 1
                                    continue

                                dados_item = [td.text for td in tds]
                                print(f"Item {indice_tr}: {dados_item[1]} - {dados_item[2][:50]}...")

                                # Clica no código do item com retry
                                retry_count = 0
                                while retry_count < max_retries:
                                    try:
                                        link = tds[1].find_element(By.TAG_NAME, "a")
                                        navegador.execute_script("arguments[0].scrollIntoView(true);", link)
                                        sleep(0.5)
                                        link.click()
                                        break
                                    except StaleElementReferenceException:
                                        retry_count += 1
                                        if retry_count >= max_retries:
                                            print(f"Erro: Elemento se tornou stale após {max_retries} tentativas")
                                            raise
                                        sleep(1)

                                # Aguarda dados de insumo
                                wait.until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvInsumo > tbody > tr:nth-child(2)")
                                ))
                                sleep(1)

                                tr_insumo = navegador.find_element(
                                    By.CSS_SELECTOR,
                                    "#ctl00_MainContent_gvInsumo > tbody > tr:nth-child(2)"
                                )

                                dados_insumo = [
                                    td.text for td in tr_insumo.find_elements(By.TAG_NAME, "td")[2:]
                                ]

                                dados_item.extend(dados_insumo)
                                dados_excel.append(dados_item)
                                print(f"Dados coletados: {len(dados_excel)} itens no total")

                                # Retorna para a lista com retry
                                retry_count = 0
                                while retry_count < max_retries:
                                    try:
                                        btn_voltar = navegador.find_element(
                                            By.ID,
                                            "ctl00_MainContent_btnServicos"
                                        )
                                        btn_voltar.click()
                                        break
                                    except StaleElementReferenceException:
                                        retry_count += 1
                                        if retry_count >= max_retries:
                                            raise
                                        sleep(1)

                                wait.until(EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                ))
                                sleep(1)

                                indice_tr += 1

                            except StaleElementReferenceException as e:
                                print(f"Elemento ficou stale. Tentando novamente...")
                                sleep(2)
                                continue
                            except TimeoutException as e:
                                print(f"Timeout ao processar linha {indice_tr}: {e}")
                                break
                            except Exception as e:
                                print(f"Erro ao processar linha {indice_tr}: {e}")
                                indice_tr += 1
                                continue

                        # Avança para próxima página
                        if pagina < paginas:
                            print(f"Avançando para próxima página...")
                            try:
                                links = navegador.find_elements(
                                    By.CSS_SELECTOR,
                                    "#ctl00_MainContent_gvServicos tr.gridPager td table tbody tr td a"
                                )

                                proxima_pagina = str(pagina + 1)
                                link_proxima = None

                                for link in links:
                                    if link.text.strip() == proxima_pagina:
                                        link_proxima = link
                                        break

                                if link_proxima:
                                    navegador.execute_script("arguments[0].scrollIntoView(true);", link_proxima)
                                    sleep(0.5)
                                    link_proxima.click()
                                else:
                                    # Procura pelos botões de reticências
                                    botoes_reticencias = [
                                        link for link in links
                                        if link.text.strip() == "..."
                                    ]

                                    if botoes_reticencias:
                                        print("Avançando bloco de páginas...")
                                        navegador.execute_script("arguments[0].scrollIntoView(true);", botoes_reticencias[-1])
                                        sleep(0.5)
                                        botoes_reticencias[-1].click()
                                    else:
                                        print("Não encontrei próxima página")
                                        break

                                wait.until(EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                ))
                                sleep(2)

                            except Exception as e:
                                print(f"Erro ao navegar para próxima página: {e}")
                                break

                    except KeyboardInterrupt:
                        raise
                    except Exception as e:
                        print(f"Erro ao processar página {pagina}: {e}")
                        continue

            except KeyboardInterrupt:
                raise
            except Exception as e:
                print(f"Erro ao processar categoria {categoria}: {e}")
                continue

        status = "FINALIZADO"

    except KeyboardInterrupt:
        print("\nProcesso interrompido pelo usuário.")
        status = "INTERROMPIDO"

    except WebDriverException as e:
        print(f"Erro crítico do WebDriver: {e}")
        status = "ERRO_WEBDRIVER"

    except Exception as e:
        print(f"Erro inesperado: {e}")
        status = "ERRO"

    finally:
        if dados_excel:
            salvar_excel(dados_excel, status)
        else:
            print("Nenhum dado foi coletado.")

        print(f"Processo {status.lower()}. Total de {len(dados_excel)} itens coletados.")


def main():
    """Função principal - executa o fluxo completo de login e exportação"""
    navegador = None
    isLoggedIn = False
    
    try:
        print("\n" + "=" * 70)
        print(" " * 15 + "SISTEMA DE EXPORTAÇÃO TCPO PINI")
        print("=" * 70 + "\n")
        
        # Inicializa navegador
        try:
            print("[→] Inicializando navegador ChromeDriver...")
            navegador = webdriver.Chrome(options=options)
            print("[✓] Navegador inicializado com sucesso\n")
        except WebDriverException as e:
            print(f"[✗] Erro ao inicializar ChromeDriver:")
            print(f"    {str(e)[:100]}...\n")
            print("⚠️  SOLUÇÕES:")
            print("   1. Verifique se o ChromeDriver está instalado")
            print("   2. Atualize ChromeDriver para a versão do seu Chrome")
            print("   3. Baixe em: https://chromedriver.chromium.org/\n")
            return False
        
        # Login
        print("[→] Realizando login...")
        isLoggedIn = login(navegador, url)
        
        if not isLoggedIn:
            print("[✗] Login falhou")
            print("    Verifique usuário e senha no arquivo .env\n")
            return False
        
        print("[✓] Login realizado\n")
        
        # Acessa banco
        print("[→] Acessando banco de dados TCPO PINI...")
        acessar_banco(navegador)
        print("[✓] Banco acessado\n")
        
        # Aguarda antes de exportar
        print("[→] Preparando para exportação (aguardando 5 segundos)...")
        sleep(5)
        
        # Exporta dados
        print("[→] Iniciando coleta de dados\n")
        print("-" * 70 + "\n")
        exportar_insumos(navegador)
        print("\n" + "-" * 70)
        print("[✓] Processo finalizado com sucesso!\n")
        return True
        
    except KeyboardInterrupt:
        print("\n\n[⚠] Processo interrompido pelo usuário\n")
        return False
    except WebDriverException as e:
        print(f"\n[✗] Erro crítico do ChromeDriver: {str(e)[:200]}\n")
        return False
    except Exception as e:
        print(f"\n[✗] Erro inesperado: {e}\n")
        import traceback
        traceback.print_exc()
        return False
    finally:
        if navegador:
            try:
                encerrar(navegador, isLoggedIn)
            except Exception as e:
                print(f"[✗] Erro ao encerrar navegador: {e}")
                try:
                    navegador.quit()
                except:
                    pass


if __name__ == "__main__":
    sucesso = main()
    exit(0 if sucesso else 1)