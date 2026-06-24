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

def aguardar_conexao(tentativa_num):
    """Aguarda o usuário reestabelecer a conexão"""
    print("\n" + "=" * 70)
    print("⚠️  CONEXÃO PERDIDA - AGUARDANDO RECONEXÃO")
    print("=" * 70)
    print(f"\nTentativa {tentativa_num} de reconexão")
    print("\nOpções:")
    print("1. Pressione ENTER para tentar continuar quando a internet voltar")
    print("2. Pressione 'Q' + ENTER para cancelar a automação")
    print("\n" + "-" * 70)
    
    resposta = input("\nAguardando sua resposta: ").strip().upper()
    
    if resposta == 'Q':
        print("\n✗ Automação cancelada pelo usuário.")
        return False
    
    print("\n✓ Tentando reconectar...")
    return True

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
    wait = WebDriverWait(navegador, 12)
    wait_rapido = WebDriverWait(navegador, 10)
    
    insumos = {
        "Materiais": "ctl00_MainContent_PiniTreeViewt304",
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

        for categoria, id_elemento in insumos.items():
            print(f"\n=== Processando Categoria: {categoria} ===")
            try:
                btnCategoria = wait.until(EC.element_to_be_clickable((By.ID, id_elemento)))
                btnCategoria.click()
                sleep(0.3)

                elemento = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_btnServicos")))
                texto = elemento.get_attribute("value")
                
                # Extrai número de páginas
                match = re.search(r'Página \d+ de (\d+)', texto)
                if not match:
                    print(f"Não foi possível extrair número de páginas. Texto: {texto}")
                    continue
                    
                paginas = int(match.group(1))
                print(f"Total de páginas a processar: {paginas}")

                for pagina in range(1, paginas + 1):
                    print(f"\nProcessando página {pagina}/{paginas}")
                    
                    try:
                        indice_tr = 2
                        max_retries = 3
                        pagina_processada = False
                        timeout_consecutivos = 0
                        indice_primeiro_timeout = None

                        while True:
                            try:
                                # Aguarda presença da tabela
                                wait_rapido.until(EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                ))

                                trs = navegador.find_elements(
                                    By.CSS_SELECTOR,
                                    "#ctl00_MainContent_gvServicos > tbody > tr"
                                )

                                if indice_tr >= len(trs) - 1:
                                    print(f"Fim da página. Total de linhas processadas: {indice_tr - 2}")
                                    pagina_processada = True
                                    break

                                tr = trs[indice_tr]
                                tds = tr.find_elements(By.TAG_NAME, "td")

                                if len(tds) < 3:
                                    indice_tr += 1
                                    continue

                                dados_item = [td.text for td in tds]

                                # Clica no código do item com retry
                                retry_count = 0
                                while retry_count < max_retries:
                                    try:
                                        link = tds[1].find_element(By.TAG_NAME, "a")
                                        link.click()
                                        break
                                    except StaleElementReferenceException:
                                        retry_count += 1
                                        if retry_count >= max_retries:
                                            raise
                                        sleep(0.2)

                                # Aguarda dados de insumo
                                wait_rapido.until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvInsumo > tbody > tr:nth-child(2)")
                                ))
                                sleep(0.2)

                                tr_insumo = navegador.find_element(
                                    By.CSS_SELECTOR,
                                    "#ctl00_MainContent_gvInsumo > tbody > tr:nth-child(2)"
                                )

                                dados_insumo = [
                                    td.text for td in tr_insumo.find_elements(By.TAG_NAME, "td")[2:]
                                ]

                                dados_item.extend(dados_insumo)
                                dados_excel.append(dados_item)
                                print(f"[{len(dados_excel)}] {dados_item[1]}")

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
                                        sleep(0.2)

                                # AGUARDA a página voltar COMPLETAMENTE antes de continuar
                                sleep(0.5)  # Aguarda física para página processar
                                wait_rapido.until(EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                ))
                                sleep(0.2)  # Aguarda mais um pouco para DOM estabilizar

                                indice_tr += 1

                            except StaleElementReferenceException as e:
                                sleep(0.2)
                                timeout_consecutivos = 0  # Reset contador
                                indice_primeiro_timeout = None
                                continue
                            except TimeoutException as e:
                                # Conta timeouts consecutivos
                                if indice_primeiro_timeout is None:
                                    indice_primeiro_timeout = indice_tr
                                
                                timeout_consecutivos += 1
                                
                                if timeout_consecutivos >= 3:
                                    # 3+ timeouts = problema de conexão
                                    if not aguardar_conexao(timeout_consecutivos):
                                        raise KeyboardInterrupt("Usuário cancelou após perda de conexão")
                                    
                                    # Volta para o primeiro item com timeout
                                    print(f"\n✓ Reconectado! Voltando para linha {indice_primeiro_timeout}...\n")
                                    indice_tr = indice_primeiro_timeout
                                    timeout_consecutivos = 0
                                    indice_primeiro_timeout = None
                                    sleep(1)
                                    continue
                                
                                # Menos de 3 timeouts - continue normalmente
                                indice_tr += 1
                                continue
                            except Exception as e:
                                print(f"Erro ao processar linha {indice_tr}: {e}")
                                indice_tr += 1
                                continue

                        # Avança para próxima página SOMENTE se ainda há páginas
                        if pagina_processada and pagina < paginas:
                            print(f"Avançando para página {pagina + 1}...")
                            tentativas_nav = 0
                            navegou = False
                            
                            while tentativas_nav < 2 and not navegou:
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
                                        link_proxima.click()
                                        navegou = True
                                    else:
                                        # Procura pelos botões de reticências
                                        botoes_reticencias = [
                                            link for link in links
                                            if link.text.strip() == "..."
                                        ]

                                        if botoes_reticencias:
                                            botoes_reticencias[-1].click()
                                            navegou = True
                                        else:
                                            print(f"Aviso: Não encontrei botão para próxima página")
                                            break

                                    if navegou:
                                        wait_rapido.until(EC.presence_of_all_elements_located(
                                            (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                        ))
                                        sleep(0.1)
                                        
                                except Exception as e:
                                    tentativas_nav += 1
                                    if tentativas_nav < 2:
                                        sleep(0.5)
                                        continue
                                    else:
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
