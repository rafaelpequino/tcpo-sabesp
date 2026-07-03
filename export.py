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
from utils import db
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
# Redução de consumo de memória
options.add_argument("--aggressive-cache-discard")
options.add_argument("--disk-cache-size=0")
options.add_argument("--media-cache-size=0")
options.add_argument("--js-flags=--max-old-space-size=512")
options.add_argument("--disable-background-networking")
options.add_argument("--disable-default-apps")
options.add_argument("--disable-sync")
options.add_argument("--disable-translate")
options.add_argument("--no-first-run")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
# Desabilita carregamento de imagens (maior economia de memória e banda)
options.add_experimental_option("prefs", {
    "profile.managed_default_content_settings.images": 2,
    "profile.default_content_setting_values.notifications": 2
})

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

    total_salvos = 0

    try:
        db.criar_tabela_insumos()
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

                # Se há mais de 3 páginas, pergunta a partir de qual começar
                pagina_inicial = 1
                if paginas > 3:
                    while True:
                        resposta = input(f"  A partir de qual página deseja começar? [1-{paginas}, ENTER para começar do início]: ").strip()
                        if resposta == "":
                            pagina_inicial = 1
                            break
                        if resposta.isdigit():
                            valor = int(resposta)
                            if 1 <= valor <= paginas:
                                pagina_inicial = valor
                                break
                        print(f"  Valor inválido. Digite um número entre 1 e {paginas}.")
                    
                    if pagina_inicial > 1:
                        print(f"\n  Navegue manualmente para a página {pagina_inicial} no navegador")
                        print(f"  e pressione ENTER quando estiver pronto.")
                        input("  Aguardando ENTER... ")
                        sleep(1)

                for pagina in range(pagina_inicial, paginas + 1):
                    print(f"\nProcessando página {pagina}/{paginas}")

                    # Limpa cache do navegador a cada 5 páginas para controlar uso de memória
                    if pagina % 5 == 0:
                        try:
                            navegador.execute_cdp_cmd("Network.clearBrowserCache", {})
                        except Exception:
                            pass

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

                                # Verifica se o item já foi extraído hoje
                                if db.item_ja_extraido_hoje(dados_item[1].strip()):
                                    print(f"  ↷ Pulando {dados_item[1].strip()} (já extraído hoje)")
                                    indice_tr += 1
                                    continue

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
                                db.salvar_insumo(
                                    base=dados_item[0],
                                    item=dados_item[1],
                                    descricao=dados_item[2],
                                    unidade=dados_item[3],
                                    tipo=dados_item[4],
                                    data_preco=dados_item[5],
                                    preco_str=dados_item[6]
                                )
                                total_salvos += 1
                                print(f"[{total_salvos}] {dados_item[1]}")

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
                            proxima_pagina_num = pagina + 1
                            print(f"Avançando para página {proxima_pagina_num}...")

                            def _pagina_atual():
                                """Lê o número da página atual a partir do texto 'Página X de Y' do botão"""
                                try:
                                    btn = navegador.find_element(By.ID, "ctl00_MainContent_btnServicos")
                                    texto = btn.get_attribute("value")
                                    m = re.search(r'Página (\d+) de', texto)
                                    return int(m.group(1)) if m else None
                                except Exception:
                                    return None

                            def _aguardar_tabela():
                                """Aguarda a tabela de serviços carregar"""
                                sleep(2)
                                wait_rapido.until(EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                ))
                                sleep(0.5)

                            def _esta_na_pagina(numero):
                                """Verifica se a página atual é a correta lendo 'Página X de Y' do botão"""
                                return _pagina_atual() == numero

                            def _pedir_navegacao_manual(numero):
                                """Congela e pede que o usuário navegue manualmente para a página"""
                                print("\n" + "=" * 70)
                                print("⚠️  NÃO FOI POSSÍVEL NAVEGAR AUTOMATICAMENTE")
                                print("=" * 70)
                                print(f"\nNão consegui ir para a página {numero} automaticamente.")
                                print(f"\nPOR FAVOR:")
                                print(f"  1. Navegue MANUALMENTE para a página {numero} no navegador")
                                print(f"  2. Aguarde a página carregar completamente")
                                print(f"  3. Pressione ENTER aqui no terminal para continuar")
                                print(f"     (ou digite 'Q' + ENTER para encerrar)\n")
                                resposta = input("Aguardando: ").strip().upper()
                                if resposta == 'Q':
                                    raise KeyboardInterrupt("Usuário cancelou após falha de navegação")
                                print("✓ Continuando...\n")

                            # --- Loop de navegação ---
                            max_tentativas = 6
                            tentativa = 0
                            navegou = False

                            while tentativa < max_tentativas and not navegou:
                                try:
                                    # Coleta os links de paginação atuais
                                    links_pager = navegador.find_elements(
                                        By.CSS_SELECTOR,
                                        "#ctl00_MainContent_gvServicos tr.gridPager td table tbody tr td a"
                                    )
                                    numero_str = str(proxima_pagina_num)

                                    # Procura link clicável com o número da próxima página
                                    link_alvo = next(
                                        (l for l in links_pager if l.text.strip() == numero_str),
                                        None
                                    )

                                    if link_alvo:
                                        # Encontrou o link — clica diretamente
                                        navegador.execute_script("arguments[0].scrollIntoView(true);", link_alvo)
                                        sleep(0.2)
                                        link_alvo.click()
                                        print(f"  Clicou no link '{proxima_pagina_num}', aguardando carregar...")
                                        _aguardar_tabela()

                                        if _esta_na_pagina(proxima_pagina_num):
                                            print(f"  ✓ Página {proxima_pagina_num} carregada com sucesso")
                                            navegou = True
                                        else:
                                            tentativa += 1
                                            print(f"  ⚠ Página ainda não mudou (tentativa {tentativa}/{max_tentativas})")
                                            sleep(0.5)
                                    else:
                                        # Link não visível — clica em "..." (que no ASP.NET já NAVEGA para o próximo grupo)
                                        botoes_reticencias = [l for l in links_pager if l.text.strip() == "..."]

                                        if botoes_reticencias:
                                            print(f"  Página {proxima_pagina_num} não está visível, clicando em '...'")
                                            botoes_reticencias[-1].click()
                                            _aguardar_tabela()

                                            pagina_obtida = _pagina_atual()
                                            if pagina_obtida == proxima_pagina_num:
                                                print(f"  ✓ Página {proxima_pagina_num} carregada via '...'")
                                                navegou = True
                                            else:
                                                tentativa += 1
                                                desc = f"página {pagina_obtida}" if pagina_obtida else "página desconhecida"
                                                print(f"  ⚠ '...' navegou para {desc}, não para {proxima_pagina_num} (tentativa {tentativa}/{max_tentativas})")
                                                sleep(0.5)
                                        else:
                                            tentativa += 1
                                            print(f"  ⚠ Sem link '{proxima_pagina_num}' e sem '...' visível (tentativa {tentativa})")
                                            sleep(0.5)

                                except StaleElementReferenceException:
                                    tentativa += 1
                                    sleep(0.5)
                                except Exception as e:
                                    print(f"  Erro: {str(e)[:80]}")
                                    tentativa += 1
                                    sleep(0.5)

                            if not navegou:
                                # Esgotou tentativas automáticas — congela e pede ajuda do usuário
                                _pedir_navegacao_manual(proxima_pagina_num)

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
        if total_salvos > 0:
            print(f"Processo {status.lower()}. Total de {total_salvos} itens salvos no banco.")
        else:
            print("Nenhum dado foi coletado.")


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
