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
import traceback

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

                # Guarda referência ao botão atual para detectar quando a página recarregar
                try:
                    btn_servicos_antigo = navegador.find_element(By.ID, "ctl00_MainContent_btnServicos")
                except NoSuchElementException:
                    btn_servicos_antigo = None

                btnCategoria.click()

                # Aguarda o elemento ficar stale (página recarregou) antes de ler os dados da nova categoria
                if btn_servicos_antigo is not None:
                    try:
                        wait.until(EC.staleness_of(btn_servicos_antigo))
                    except TimeoutException:
                        pass

                elemento = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_btnServicos")))
                texto = elemento.get_attribute("value")
                
                # Extrai número de páginas
                match = re.search(r'Página \d+ de (\d+)', texto)
                if not match:
                    print(f"Não foi possível extrair número de páginas. Texto: {texto}")
                    continue
                    
                paginas = int(match.group(1))
                print(f"Total de páginas a processar: {paginas}")

                # Sempre começa pela página 1
                pagina_inicial = 1

                # Categoria com página única: pula só a 1ª linha e lê até o fim (sem linha de paginação)
                # Categoria com múltiplas páginas: pula as 2 primeiras e para antes da última (linha de paginação)
                tr_inicio = 1 if paginas == 1 else 2
                tr_fim_offset = 0 if paginas == 1 else 1

                for pagina in range(pagina_inicial, paginas + 1):
                    print(f"\nProcessando página {pagina}/{paginas}")

                    # Limpa cache do navegador a cada 5 páginas para controlar uso de memória
                    if pagina % 5 == 0:
                        try:
                            navegador.execute_cdp_cmd("Network.clearBrowserCache", {})
                        except Exception:
                            pass

                    try:
                        indice_tr = tr_inicio
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

                                if indice_tr >= len(trs) - tr_fim_offset:
                                    print(f"Fim da página. Total de linhas processadas: {indice_tr - tr_inicio}")
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
                                    # Verifica de imediato: se já estamos na página alvo (ex.: click anterior funcionou
                                    # mas a verificação pós-click falhou por timing), encerra sem tentar de novo
                                    if _esta_na_pagina(proxima_pagina_num):
                                        print(f"  ✓ Página {proxima_pagina_num} carregada com sucesso")
                                        navegou = True
                                        break

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
                                        # Captura botão atual antes de clicar para detectar staleness do postback
                                        btn_antes = navegador.find_element(By.ID, "ctl00_MainContent_btnServicos")
                                        navegador.execute_script("arguments[0].scrollIntoView(true);", link_alvo)
                                        sleep(0.2)
                                        link_alvo.click()
                                        print(f"  Clicou no link '{proxima_pagina_num}', aguardando carregar...")

                                        # Aguarda o postback completar (btn fica stale) antes de verificar página
                                        try:
                                            wait.until(EC.staleness_of(btn_antes))
                                        except TimeoutException:
                                            pass
                                        _aguardar_tabela()

                                        if _esta_na_pagina(proxima_pagina_num):
                                            print(f"  ✓ Página {proxima_pagina_num} carregada com sucesso")
                                            navegou = True
                                        else:
                                            tentativa += 1
                                            print(f"  ⚠ Navegação incerta (tentativa {tentativa}/{max_tentativas}), verificando novamente...")
                                            sleep(0.5)
                                    else:
                                        # Link não visível — clica em "..." (que no ASP.NET já NAVEGA para o próximo grupo)
                                        botoes_reticencias = [l for l in links_pager if l.text.strip() == "..."]

                                        if botoes_reticencias:
                                            print(f"  Página {proxima_pagina_num} não está visível, clicando em '...'")
                                            btn_antes = navegador.find_element(By.ID, "ctl00_MainContent_btnServicos")
                                            botoes_reticencias[-1].click()

                                            try:
                                                wait.until(EC.staleness_of(btn_antes))
                                            except TimeoutException:
                                                pass
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
                        print(f"\n⚠️  Erro ao processar página {pagina}:")
                        print(f"    Tipo: {type(e).__name__}")
                        print(f"    Mensagem: {str(e)}")
                        if str(e).strip():
                            print(f"    Traceback:\n{traceback.format_exc()}")
                        continue

            except KeyboardInterrupt:
                raise
            except Exception as e:
                print(f"\n⚠️  Erro ao processar categoria {categoria}:")
                print(f"    Tipo: {type(e).__name__}")
                print(f"    Mensagem: {str(e)}")
                if str(e).strip():
                    print(f"    Traceback:\n{traceback.format_exc()}")
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


def exportar_servicos(navegador):
    wait = WebDriverWait(navegador, 12)
    wait_rapido = WebDriverWait(navegador, 10)

    # TODO: substitua os IDs pelos valores reais da árvore de navegação do TCPO PINI
    servicos = {
        "Canteiro de obras": ["ctl00_MainContent_PiniTreeViewt3"],
        "Demolições": ["ctl00_MainContent_PiniTreeViewt4"],
        "Limpeza terreno": ["ctl00_MainContent_PiniTreeViewt5"],
        "Locação": ["ctl00_MainContent_PiniTreeViewt6"],
        "Movimento de terra": ["ctl00_MainContent_PiniTreeViewt7"],
        "Sondagem": ["ctl00_MainContent_PiniTreeViewt8"],
        "Armaduras": ["ctl00_MainContent_PiniTreeViewt10", "ctl00_MainContent_PiniTreeViewt21"],
        "Concreto": ["ctl00_MainContent_PiniTreeViewt11", "ctl00_MainContent_PiniTreeViewt24", "ctl00_MainContent_PiniTreeViewt141"],
        "Composições TCPO": [
            "ctl00_MainContent_PiniTreeViewt13",
            "ctl00_MainContent_PiniTreeViewt15",
            "ctl00_MainContent_PiniTreeViewt27",
            "ctl00_MainContent_PiniTreeViewt32",
            "ctl00_MainContent_PiniTreeViewt34",
            "ctl00_MainContent_PiniTreeViewt42",
            "ctl00_MainContent_PiniTreeViewt73",
            "ctl00_MainContent_PiniTreeViewt126",
            "ctl00_MainContent_PiniTreeViewt132",
         ],
        "Fôrmas": ["ctl00_MainContent_PiniTreeViewt16", "ctl00_MainContent_PiniTreeViewt25"],
        "Fundações profundas": ["ctl00_MainContent_PiniTreeViewt17"],
        "Fundações superficiais": ["ctl00_MainContent_PiniTreeViewt18"],
        "Lastros": ["ctl00_MainContent_PiniTreeViewt19"],
        "Chumbadores": ["ctl00_MainContent_PiniTreeViewt22"],
        "Cimbramento": ["ctl00_MainContent_PiniTreeViewt23"],
        "Reparo e reforço estrutural": ["ctl00_MainContent_PiniTreeViewt28"],
        "Serviços relacionados": [
            "ctl00_MainContent_PiniTreeViewt29",
            "ctl00_MainContent_PiniTreeViewt35",
            "ctl00_MainContent_PiniTreeViewt40",
            "ctl00_MainContent_PiniTreeViewt57",
             "ctl00_MainContent_PiniTreeViewt74",
             "ctl00_MainContent_PiniTreeViewt107",
             "ctl00_MainContent_PiniTreeViewt113",
             "ctl00_MainContent_PiniTreeViewt117",
             "ctl00_MainContent_PiniTreeViewt128",
             "ctl00_MainContent_PiniTreeViewt152",
             "ctl00_MainContent_PiniTreeViewt167",
        ],
        "Calhas, rufos e condutores": ["ctl00_MainContent_PiniTreeViewt37"],
        "Domos / iluminação zenital": ["ctl00_MainContent_PiniTreeViewt38"],
        "Estruturas": ["ctl00_MainContent_PiniTreeViewt39"],
        "Alicerce": ["ctl00_MainContent_PiniTreeViewt44"],
        "Cobertura": ["ctl00_MainContent_PiniTreeViewt45"],
        "Geotêxtil": ["ctl00_MainContent_PiniTreeViewt46"],
        "Outras soluções": ["ctl00_MainContent_PiniTreeViewt47"],
        "Acústico": ["ctl00_MainContent_PiniTreeViewt49"],
        "Térmico": ["ctl00_MainContent_PiniTreeViewt50"],
        "Ferragens": ["ctl00_MainContent_PiniTreeViewt52"],
        "Grades e gradis": ["ctl00_MainContent_PiniTreeViewt53"],
        "Janelas": ["ctl00_MainContent_PiniTreeViewt54"],
        "Portas": ["ctl00_MainContent_PiniTreeViewt55"],
        "Portões": ["ctl00_MainContent_PiniTreeViewt56"],
        "Tubos e conexões em PVC": ["ctl00_MainContent_PiniTreeViewt59"], 
        "Tubos e conexões em CPVC": ["ctl00_MainContent_PiniTreeViewt60"],
        "Tubos e conexões em PPR": ["ctl00_MainContent_PiniTreeViewt61"],
        "Tubos e conexões em PEX": ["ctl00_MainContent_PiniTreeViewt62"],
        "Tubos e conexões em ferro galvanizado": ["ctl00_MainContent_PiniTreeViewt63"],
        "Tubos e conexões em cobre e bronze": ["ctl00_MainContent_PiniTreeViewt64"],
        "Tubos e conexões em ferro fundido": ["ctl00_MainContent_PiniTreeViewt65"],
        "Tubos e conexões cerâmicos": ["ctl00_MainContent_PiniTreeViewt66"],
        "Válvulas": ["ctl00_MainContent_PiniTreeViewt67"],
        "Registros": ["ctl00_MainContent_PiniTreeViewt68"],
        "Caixas, ralos e grelhas": ["ctl00_MainContent_PiniTreeViewt69"],
        "Cisternas e reservatórios": ["ctl00_MainContent_PiniTreeViewt70"],
        "Conjuntos elevatórios motor-bomba": ["ctl00_MainContent_PiniTreeViewt71"],
        "Alarmes": ["ctl00_MainContent_PiniTreeViewt76"],
        "Hidrantes": ["ctl00_MainContent_PiniTreeViewt77"],
        "Extintores": ["ctl00_MainContent_PiniTreeViewt78"],
        "Entrada de energia em caixa": ["ctl00_MainContent_PiniTreeViewt80"],
        "Entrada de energia em poste": ["ctl00_MainContent_PiniTreeViewt81"],
        "Transformadores": ["ctl00_MainContent_PiniTreeViewt82"],
        "Quadros de distribuição": ["ctl00_MainContent_PiniTreeViewt83"],
        "Disjuntores": ["ctl00_MainContent_PiniTreeViewt84"],
        "Chaves seccionadoras": ["ctl00_MainContent_PiniTreeViewt85"],
        "Canaletas para instalação aparente": ["ctl00_MainContent_PiniTreeViewt86"],
        "Duto em PEAD": ["ctl00_MainContent_PiniTreeViewt87"],
        "Eletrodutos de aço": ["ctl00_MainContent_PiniTreeViewt88"],
        "Eletrodutos flexíveis de PVC": ["ctl00_MainContent_PiniTreeViewt89"],
        "Eletrodutos de PVC": ["ctl00_MainContent_PiniTreeViewt90"],
        "Acessórios para perfilados e eletrocalhas": ["ctl00_MainContent_PiniTreeViewt91"],
        "Dutos": ["ctl00_MainContent_PiniTreeViewt92", "ctl00_MainContent_PiniTreeViewt120"],
        "Eletrocalhas": ["ctl00_MainContent_PiniTreeViewt93"],
        "Leitos": ["ctl00_MainContent_PiniTreeViewt94"],
         "Perfilados": ["ctl00_MainContent_PiniTreeViewt95"],
        "Caixas de passagem em aço": ["ctl00_MainContent_PiniTreeViewt96"],
        "Conduletes de alumínio": ["ctl00_MainContent_PiniTreeViewt97"],
        "Caixas de PVC": ["ctl00_MainContent_PiniTreeViewt98"],
        "Conexões e acessórios em aço": ["ctl00_MainContent_PiniTreeViewt99"],
        "Conexões e acessórios em PEAD": ["ctl00_MainContent_PiniTreeViewt100"],
        "Conexões e acessórios em PVC": ["ctl00_MainContent_PiniTreeViewt101"],
        "Conexões e acessórios para eletroduto de PVC flexível": ["ctl00_MainContent_PiniTreeViewt102"],
        "Cabos e fios": ["ctl00_MainContent_PiniTreeViewt103"],
        "Interruptores e tomadas": ["ctl00_MainContent_PiniTreeViewt104"],
        "Luminárias": ["ctl00_MainContent_PiniTreeViewt105"],
        "Iluminação de via pública": ["ctl00_MainContent_PiniTreeViewt106"],
        "Cabos": ["ctl00_MainContent_PiniTreeViewt109"],
        "Caixas e conduletes": ["ctl00_MainContent_PiniTreeViewt110"],
        "Placas e tomadas": ["ctl00_MainContent_PiniTreeViewt111"],
        "Segurança": ["ctl00_MainContent_PiniTreeViewt112"],
        "Aterramento - hastes e cordoalhas": ["ctl00_MainContent_PiniTreeViewt115"],
        "Captores": ["ctl00_MainContent_PiniTreeViewt116"],
        "Sinalizadores": ["ctl00_MainContent_PiniTreeViewt118"],
        "Rede frigorígena": ["ctl00_MainContent_PiniTreeViewt121"],
        "Chapisco": ["ctl00_MainContent_PiniTreeViewt123"],
        "Emboço": ["ctl00_MainContent_PiniTreeViewt124"],
        "Reboco": ["ctl00_MainContent_PiniTreeViewt127"],
        "Fibra mineral": ["ctl00_MainContent_PiniTreeViewt130"],
        "Madeira": ["ctl00_MainContent_PiniTreeViewt133", "ctl00_MainContent_PiniTreeViewt146"],
        "Metálicos": ["ctl00_MainContent_PiniTreeViewt134", "ctl00_MainContent_PiniTreeViewt158"],
        "PVC": ["ctl00_MainContent_PiniTreeViewt135"],
        "Alta resistência": ["ctl00_MainContent_PiniTreeViewt137"],
        "Borracha": ["ctl00_MainContent_PiniTreeViewt138"],
        "Cerâmicos": ["ctl00_MainContent_PiniTreeViewt139"],
        "Cimentados": ["ctl00_MainContent_PiniTreeViewt140"],
        "Elevados": ["ctl00_MainContent_PiniTreeViewt142"],
        "Epóxi": ["ctl00_MainContent_PiniTreeViewt143"],
        "Granilite": ["ctl00_MainContent_PiniTreeViewt144"],
        "Ladrilhos hidráulicos": ["ctl00_MainContent_PiniTreeViewt145"],
        "Pastilhas": ["ctl00_MainContent_PiniTreeViewt147", "ctl00_MainContent_PiniTreeViewt160"],
        "Pedras": ["ctl00_MainContent_PiniTreeViewt148"],
        "Podotáteis": ["ctl00_MainContent_PiniTreeViewt149"],
        "Têxteis": ["ctl00_MainContent_PiniTreeViewt150"],
        "Vinílicos": ["ctl00_MainContent_PiniTreeViewt151", "ctl00_MainContent_PiniTreeViewt161"],
        "Azulejo": ["ctl00_MainContent_PiniTreeViewt154"],
        "Cerâmica": ["ctl00_MainContent_PiniTreeViewt155"],
        "Massa única": ["ctl00_MainContent_PiniTreeViewt156"],
        "Melamínico": ["ctl00_MainContent_PiniTreeViewt157"],
        "Papel": ["ctl00_MainContent_PiniTreeViewt159"],
        "Esquadrias de madeira": ["ctl00_MainContent_PiniTreeViewt163"],
        "Esquadrias metálicas": ["ctl00_MainContent_PiniTreeViewt164"],
        "Paredes e tetos": ["ctl00_MainContent_PiniTreeViewt165"],
        "Pisos": ["ctl00_MainContent_PiniTreeViewt166"],
        "Tratamento de concreto": ["ctl00_MainContent_PiniTreeViewt168"],
        "Bacias sanitárias e bidês": ["ctl00_MainContent_PiniTreeViewt170"],
        "Bancadas": ["ctl00_MainContent_PiniTreeViewt171"],
        "Banheiras": ["ctl00_MainContent_PiniTreeViewt172"],
        "Barra de apoio": ["ctl00_MainContent_PiniTreeViewt173"],
        "Caixas de descarga": ["ctl00_MainContent_PiniTreeViewt174"],
        "Chuveiros e duchas": ["ctl00_MainContent_PiniTreeViewt175"],
        "Gabinetes": ["ctl00_MainContent_PiniTreeViewt176"],
        "Lavatórios e cubas": ["ctl00_MainContent_PiniTreeViewt177"],
        "Mictórios": ["ctl00_MainContent_PiniTreeViewt178"],
        "Misturadores": ["ctl00_MainContent_PiniTreeViewt179", "ctl00_MainContent_PiniTreeViewt237"],
        "Pias": ["ctl00_MainContent_PiniTreeViewt180"],
        "Piso-box": ["ctl00_MainContent_PiniTreeViewt181"],
        "Porta-papel": ["ctl00_MainContent_PiniTreeViewt182"],
        "Porta-toalha": ["ctl00_MainContent_PiniTreeViewt183"],
        "Pressurizador": ["ctl00_MainContent_PiniTreeViewt184"],
        "Saboneteira": ["ctl00_MainContent_PiniTreeViewt185"],
        "Tanques": ["ctl00_MainContent_PiniTreeViewt186"],
        "Torneiras": ["ctl00_MainContent_PiniTreeViewt187"],
        "Liso": ["ctl00_MainContent_PiniTreeViewt189"],
        "Fantasia": ["ctl00_MainContent_PiniTreeViewt190"],
        "Temperado": ["ctl00_MainContent_PiniTreeViewt191"],
        "Laminado": ["ctl00_MainContent_PiniTreeViewt192"],
        "Aramados": ["ctl00_MainContent_PiniTreeViewt193"],
        "Refletivo": ["ctl00_MainContent_PiniTreeViewt194"],
        "Espelhos": ["ctl00_MainContent_PiniTreeViewt195"],
        "Redes externas, drenagem - Redes e galerias": ["ctl00_MainContent_PiniTreeViewt197"],
        "Caixas de inspeção": ["ctl00_MainContent_PiniTreeViewt198"],
        "Poços de visita": ["ctl00_MainContent_PiniTreeViewt199"],
        "Bocas de lobo": ["ctl00_MainContent_PiniTreeViewt200"],
        "Canaletas": ["ctl00_MainContent_PiniTreeViewt201"],
        "Guias e sarjetas": ["ctl00_MainContent_PiniTreeViewt202"],
        "Redes externas - PEAD": ["ctl00_MainContent_PiniTreeViewt203"],
        "Redes externas - PVC PBA": ["ctl00_MainContent_PiniTreeViewt204"],
        "Redes externas - MPVC DEFoFo": ["ctl00_MainContent_PiniTreeViewt205"],
        "Redes externas - Serviços de implantação (escavação, lastro e reaterro)": ["ctl00_MainContent_PiniTreeViewt206"],
        "Redes externas, esgoto - PEAD": ["ctl00_MainContent_PiniTreeViewt207"],
        "Redes externas, esgoto - PVC": ["ctl00_MainContent_PiniTreeViewt208"],
        "Redes externas, esgoto - Serviços de implantação (escavação, lastro e reaterro)": ["ctl00_MainContent_PiniTreeViewt209"],
        "Redes externas aéreas - Postes": ["ctl00_MainContent_PiniTreeViewt210"],
        "Redes externas, energia - Cabos": ["ctl00_MainContent_PiniTreeViewt211"],
        "Redes externas enterradas - Dutos subterrâneos": ["ctl00_MainContent_PiniTreeViewt212"],
        "Redes externas - Iluminação": ["ctl00_MainContent_PiniTreeViewt213"],
        "Redes externas, energia - Serviços de implantação (escavação, lastro e reaterro)": ["ctl00_MainContent_PiniTreeViewt214"],
        "Pavimentação - Preparo e camadas intermediárias": ["ctl00_MainContent_PiniTreeViewt215"],
        "Pavimentos asfálticos": ["ctl00_MainContent_PiniTreeViewt216"],
        "Pavimentos de blocos": ["ctl00_MainContent_PiniTreeViewt217"],
        "Passeios, calçadas e acessibilidade": ["ctl00_MainContent_PiniTreeViewt218"],
        "Sinalização viária": ["ctl00_MainContent_PiniTreeViewt219"],
        "Fechamento perimétrico": ["ctl00_MainContent_PiniTreeViewt220"],
        "Paisagismo": ["ctl00_MainContent_PiniTreeViewt221"],
        "Recreação, esporte e lazer": ["ctl00_MainContent_PiniTreeViewt222"],
        "Horizontal": ["ctl00_MainContent_PiniTreeViewt224"],
        "Vertical": ["ctl00_MainContent_PiniTreeViewt225"],
        "Poços": ["ctl00_MainContent_PiniTreeViewt227"],
        "Fossas sépticas, filtros e sumidouros": ["ctl00_MainContent_PiniTreeViewt228"],
        "Ligações a redes de concessionárias de serviços públicos": ["ctl00_MainContent_PiniTreeViewt229"],
        "Serviços complementares": ["ctl00_MainContent_PiniTreeViewt230"],
        "Compressores": ["ctl00_MainContent_PiniTreeViewt232"],
        "Máquinas para pintura": ["ctl00_MainContent_PiniTreeViewt233"],
        "Compactadores de percussão": ["ctl00_MainContent_PiniTreeViewt234"],
        "Compactadores de placa vibratória": ["ctl00_MainContent_PiniTreeViewt235"],
        "Acabadoras": ["ctl00_MainContent_PiniTreeViewt236"],
        "Bombas para concreto e argamassas": ["ctl00_MainContent_PiniTreeViewt238"],
        "Desempenadeiras": ["ctl00_MainContent_PiniTreeViewt239"],
        "Usinas para concreto": ["ctl00_MainContent_PiniTreeViewt240"],
        "Vibradores de imersão": ["ctl00_MainContent_PiniTreeViewt241"],
        "Cortadoras e dobradoras": ["ctl00_MainContent_PiniTreeViewt242"],
        "Marteletes": ["ctl00_MainContent_PiniTreeViewt243"],
        "Bombas para drenagem": ["ctl00_MainContent_PiniTreeViewt244"],
        "Carregadeira": ["ctl00_MainContent_PiniTreeViewt245"],
        "Minicarregadeira": ["ctl00_MainContent_PiniTreeViewt246"],
        "Retroescavadeira": ["ctl00_MainContent_PiniTreeViewt247"],
        "Bate-estacas": ["ctl00_MainContent_PiniTreeViewt248"],
        "Máquinas para execução de estaca": ["ctl00_MainContent_PiniTreeViewt249"],
        "Campânulas": ["ctl00_MainContent_PiniTreeViewt250"],
        "Grupos geradores de energia elétrica": ["ctl00_MainContent_PiniTreeViewt251"],
        "Gruas fixas": ["ctl00_MainContent_PiniTreeViewt252"],
        "Gruas móveis": ["ctl00_MainContent_PiniTreeViewt253"],
        "Guinchos": ["ctl00_MainContent_PiniTreeViewt254"],
        "Guindastes": ["ctl00_MainContent_PiniTreeViewt255"],
        "Tratores": ["ctl00_MainContent_PiniTreeViewt256"],
        "Esmeris": ["ctl00_MainContent_PiniTreeViewt257"],
        "Furadeiras": ["ctl00_MainContent_PiniTreeViewt258"],
        "Motosserras": ["ctl00_MainContent_PiniTreeViewt259"],
        "Prensas hidráulicas": ["ctl00_MainContent_PiniTreeViewt260"],
        "Roçadeiras": ["ctl00_MainContent_PiniTreeViewt261"],
        "Serras": ["ctl00_MainContent_PiniTreeViewt262"],
        "Máquinas para corte e solda": ["ctl00_MainContent_PiniTreeViewt263"],
        "Talhas": ["ctl00_MainContent_PiniTreeViewt264"],
        "Termofusores": ["ctl00_MainContent_PiniTreeViewt265"],
        "Perfuratrizes": ["ctl00_MainContent_PiniTreeViewt266"],
        "Caldeiras": ["ctl00_MainContent_PiniTreeViewt267"],
        "Distribuidores de agregados": ["ctl00_MainContent_PiniTreeViewt268"],
        "Distribuidores para betume": ["ctl00_MainContent_PiniTreeViewt269"],
        "Distribuidores para lama asfáltica": ["ctl00_MainContent_PiniTreeViewt270"],
        "Escavadeiras": ["ctl00_MainContent_PiniTreeViewt271"],
        "Grades de disco": ["ctl00_MainContent_PiniTreeViewt272"],
        "Escreiperes": ["ctl00_MainContent_PiniTreeViewt273"],
        "Motoniveladoras": ["ctl00_MainContent_PiniTreeViewt274"],
        "Carregadeiras sobre pneus": ["ctl00_MainContent_PiniTreeViewt275"],
        "Pintura de faixas": ["ctl00_MainContent_PiniTreeViewt276"],
        "Rolos compactadores": ["ctl00_MainContent_PiniTreeViewt277"],
        "Tratores sobre esteiras": ["ctl00_MainContent_PiniTreeViewt278"],
        "Tratores sobre pneus": ["ctl00_MainContent_PiniTreeViewt279"],
        "Usinas misturadoras": ["ctl00_MainContent_PiniTreeViewt280"],
        "Usinas para asfalto": ["ctl00_MainContent_PiniTreeViewt281"],
        "Fresadoras": ["ctl00_MainContent_PiniTreeViewt282"],
        "Recicladoras de pavimento": ["ctl00_MainContent_PiniTreeViewt283"],
        "Tanques de asfalto": ["ctl00_MainContent_PiniTreeViewt284"],
        "Vassouras mecânicas": ["ctl00_MainContent_PiniTreeViewt285"],
        "Vibroacabadoras": ["ctl00_MainContent_PiniTreeViewt286"],
        "Seladoras": ["ctl00_MainContent_PiniTreeViewt287"],
        "Caminhões basculante": ["ctl00_MainContent_PiniTreeViewt288"],
        "Caminhões betoneira": ["ctl00_MainContent_PiniTreeViewt289"],
        "Caminhões carroceria de madeira": ["ctl00_MainContent_PiniTreeViewt290"],
        "Caminhões espargidores": ["ctl00_MainContent_PiniTreeViewt291"],
        "Caminhões com guindaste": ["ctl00_MainContent_PiniTreeViewt292"],
        "Caminhões com plataforma elevatória": ["ctl00_MainContent_PiniTreeViewt293"],
        "Caminhões tanque com hidrossemeador": ["ctl00_MainContent_PiniTreeViewt294"],
        "Caminhões tanque com irrigador": ["ctl00_MainContent_PiniTreeViewt295"],
        "Carretas semirreboque": ["ctl00_MainContent_PiniTreeViewt296"],
        "Cavalos mecânicos": ["ctl00_MainContent_PiniTreeViewt297"],
        "Cavalos mecânicos com carreta": ["ctl00_MainContent_PiniTreeViewt298"],
        "Empilhadeiras": ["ctl00_MainContent_PiniTreeViewt299"],
        "Veículos leves - furgões": ["ctl00_MainContent_PiniTreeViewt300"],
        "Veículos leves - automóveis": ["ctl00_MainContent_PiniTreeViewt301"],
        "Veículos leves - pick-up": ["ctl00_MainContent_PiniTreeViewt302"],
    }
    total_salvos = 0

    # Achata o dict (categoria → lista de IDs) em pares (categoria, id_elemento)
    categorias_ids = [
        (cat, id_el)
        for cat, ids in servicos.items()
        for id_el in ids
    ]

    try:
        db.criar_tabela_servicos()
        db.criar_tabela_composicoes()
        print("Iniciando leitura de serviços...")

        for categoria, id_elemento in categorias_ids:
            tipo_servico = categoria.upper()
            print(f"\n=== Processando Categoria: {categoria} | ID: {id_elemento} ===")
            try:
                # Expande toda a árvore ANTES de cada categoria
                try:
                    navegador.execute_script("""
                        document.querySelectorAll('#ctl00_MainContent_PiniTreeView *').forEach(function(el) {
                            el.style.display = 'block';
                        });
                    """)
                except Exception as e:
                    print(f"  Aviso: Não foi possível expandir árvore: {e}")

                btnCategoria = wait.until(EC.element_to_be_clickable((By.ID, id_elemento)))

                try:
                    btn_servicos_antigo = navegador.find_element(By.ID, "ctl00_MainContent_btnServicos")
                except NoSuchElementException:
                    btn_servicos_antigo = None

                btnCategoria.click()

                if btn_servicos_antigo is not None:
                    try:
                        wait.until(EC.staleness_of(btn_servicos_antigo))
                    except TimeoutException:
                        pass

                elemento = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_btnServicos")))
                texto = elemento.get_attribute("value")

                match = re.search(r'Página \d+ de (\d+)', texto)
                if not match:
                    print(f"Não foi possível extrair número de páginas. Texto: {texto}")
                    continue

                paginas = int(match.group(1))
                print(f"Total de páginas a processar: {paginas}")

                # Sempre começa pela página 1
                pagina_inicial = 1

                tr_inicio = 1 if paginas == 1 else 2
                tr_fim_offset = 0 if paginas == 1 else 1

                for pagina in range(pagina_inicial, paginas + 1):
                    print(f"\nProcessando página {pagina}/{paginas}")

                    if pagina % 5 == 0:
                        try:
                            navegador.execute_cdp_cmd("Network.clearBrowserCache", {})
                        except Exception:
                            pass

                    try:
                        indice_tr = tr_inicio
                        max_retries = 3
                        pagina_processada = False
                        timeout_consecutivos = 0
                        indice_primeiro_timeout = None

                        while True:
                            try:
                                wait_rapido.until(EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                ))

                                trs = navegador.find_elements(
                                    By.CSS_SELECTOR,
                                    "#ctl00_MainContent_gvServicos > tbody > tr"
                                )

                                if indice_tr >= len(trs) - tr_fim_offset:
                                    print(f"Fim da página. Total de linhas processadas: {indice_tr - tr_inicio}")
                                    pagina_processada = True
                                    break

                                tr = trs[indice_tr]
                                tds = tr.find_elements(By.TAG_NAME, "td")

                                if len(tds) < 4:
                                    indice_tr += 1
                                    continue

                                base      = tds[0].text.strip()
                                item      = tds[1].text.strip()
                                descricao = tds[2].text.strip()
                                unidade   = tds[3].text.strip()

                                # Clica no link do item com retry
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

                                # Aguarda carregamento da página do serviço
                                wait_rapido.until(EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_txtDataPreco")
                                ))
                                sleep(0.2)

                                # Coleta dados complementares do serviço
                                el_data = navegador.find_element(By.CSS_SELECTOR, "#ctl00_MainContent_txtDataPreco")
                                data_preco = el_data.get_attribute("value").strip()

                                el_preco = navegador.find_element(By.CSS_SELECTOR, "#ctl00_MainContent_lblValorTotalSemTaxa")
                                preco_str = el_preco.text.strip()

                                # Salva o serviço apenas se ainda não existir
                                if not db.servico_ja_extraido(item):
                                    db.salvar_servico(
                                        base=base,
                                        item=item,
                                        descricao=descricao,
                                        unidade=unidade,
                                        tipo=tipo_servico,
                                        data_preco=data_preco,
                                        preco_str=preco_str
                                    )
                                    total_salvos += 1
                                    print(f"[{total_salvos}] {item}")
                                else:
                                    print(f"  ↷ Serviço já existe: {item} (verificando composições...)")

                                # Lê tabela de composição (sempre, mesmo que o serviço já exista)
                                try:
                                    wait_rapido.until(EC.presence_of_element_located(
                                        (By.CSS_SELECTOR, "#ctl00_MainContent_gvComposicao > tbody > tr:nth-child(2)")
                                    ))

                                    trs_composicao = navegador.find_elements(
                                        By.CSS_SELECTOR,
                                        "#ctl00_MainContent_gvComposicao > tbody > tr"
                                    )

                                    # Pula header (linha 0) e as 3 últimas linhas (somatórios)
                                    linhas_validas = trs_composicao[1:-3] if len(trs_composicao) > 4 else []

                                    for tr_comp in linhas_validas:
                                        tds_comp = tr_comp.find_elements(By.TAG_NAME, "td")
                                        if len(tds_comp) < 8:
                                            continue

                                        item_insumo    = tds_comp[0].text.strip()
                                        desc_insumo    = tds_comp[1].text.strip()
                                        und_insumo     = tds_comp[2].text.strip()
                                        coef_str       = tds_comp[4].text.strip()
                                        
                                        # Tenta pegar preço unitário como input (tem field type="text")
                                        try:
                                            input_preco = tds_comp[5].find_element(By.TAG_NAME, "input")
                                            preco_unit_str = input_preco.get_attribute("value").strip()
                                        except NoSuchElementException:
                                            # Se não for input, pega como texto normal
                                            preco_unit_str = tds_comp[5].text.strip()
                                        
                                        preco_tot_str  = tds_comp[6].text.strip()
                                        consumo_str    = tds_comp[7].text.strip()

                                        # Verifica se a composição já existe
                                        if db.composicao_ja_existe(item, item_insumo):
                                            continue

                                        if not db.insumo_existe(item_insumo):
                                            db.salvar_insumo(
                                                base=base,
                                                item=item_insumo,
                                                descricao=desc_insumo,
                                                unidade=und_insumo,
                                                tipo="EXTRA",
                                                data_preco=data_preco,
                                                preco_str=preco_unit_str
                                            )
                                            print(f"    [+] Insumo EXTRA inserido: {item_insumo}")

                                        db.salvar_composicao(
                                            item_servico=item,
                                            item_insumo=item_insumo,
                                            data_preco=data_preco,
                                            coef_str=coef_str,
                                            preco_unit_str=preco_unit_str,
                                            preco_tot_str=preco_tot_str,
                                            consumo_str=consumo_str
                                        )
                                        print(f"    [✓] Composição adicionada: {item_insumo}")

                                except TimeoutException:
                                    print(f"    [!] Sem composição para {item}")

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

                                sleep(0.5)
                                wait_rapido.until(EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                ))
                                sleep(0.2)

                                indice_tr += 1

                            except StaleElementReferenceException:
                                sleep(0.2)
                                timeout_consecutivos = 0
                                indice_primeiro_timeout = None
                                continue
                            except TimeoutException:
                                if indice_primeiro_timeout is None:
                                    indice_primeiro_timeout = indice_tr

                                timeout_consecutivos += 1

                                if timeout_consecutivos >= 3:
                                    if not aguardar_conexao(timeout_consecutivos):
                                        raise KeyboardInterrupt("Usuário cancelou após perda de conexão")

                                    print(f"\n✓ Reconectado! Voltando para linha {indice_primeiro_timeout}...\n")
                                    indice_tr = indice_primeiro_timeout
                                    timeout_consecutivos = 0
                                    indice_primeiro_timeout = None
                                    sleep(1)
                                    continue

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
                                try:
                                    btn = navegador.find_element(By.ID, "ctl00_MainContent_btnServicos")
                                    texto = btn.get_attribute("value")
                                    m = re.search(r'Página (\d+) de', texto)
                                    return int(m.group(1)) if m else None
                                except Exception:
                                    return None

                            def _aguardar_tabela():
                                sleep(2)
                                wait_rapido.until(EC.presence_of_all_elements_located(
                                    (By.CSS_SELECTOR, "#ctl00_MainContent_gvServicos > tbody > tr")
                                ))
                                sleep(0.5)

                            def _esta_na_pagina(numero):
                                return _pagina_atual() == numero

                            def _pedir_navegacao_manual(numero):
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

                            max_tentativas = 6
                            tentativa = 0
                            navegou = False

                            while tentativa < max_tentativas and not navegou:
                                try:
                                    if _esta_na_pagina(proxima_pagina_num):
                                        print(f"  ✓ Página {proxima_pagina_num} carregada com sucesso")
                                        navegou = True
                                        break

                                    links_pager = navegador.find_elements(
                                        By.CSS_SELECTOR,
                                        "#ctl00_MainContent_gvServicos tr.gridPager td table tbody tr td a"
                                    )
                                    numero_str = str(proxima_pagina_num)

                                    link_alvo = next(
                                        (l for l in links_pager if l.text.strip() == numero_str),
                                        None
                                    )

                                    if link_alvo:
                                        btn_antes = navegador.find_element(By.ID, "ctl00_MainContent_btnServicos")
                                        navegador.execute_script("arguments[0].scrollIntoView(true);", link_alvo)
                                        sleep(0.2)
                                        link_alvo.click()
                                        print(f"  Clicou no link '{proxima_pagina_num}', aguardando carregar...")

                                        try:
                                            wait.until(EC.staleness_of(btn_antes))
                                        except TimeoutException:
                                            pass
                                        _aguardar_tabela()

                                        if _esta_na_pagina(proxima_pagina_num):
                                            print(f"  ✓ Página {proxima_pagina_num} carregada com sucesso")
                                            navegou = True
                                        else:
                                            tentativa += 1
                                            print(f"  ⚠ Navegação incerta (tentativa {tentativa}/{max_tentativas}), verificando novamente...")
                                            sleep(0.5)
                                    else:
                                        botoes_reticencias = [l for l in links_pager if l.text.strip() == "..."]

                                        if botoes_reticencias:
                                            print(f"  Página {proxima_pagina_num} não está visível, clicando em '...'")
                                            btn_antes = navegador.find_element(By.ID, "ctl00_MainContent_btnServicos")
                                            botoes_reticencias[-1].click()

                                            try:
                                                wait.until(EC.staleness_of(btn_antes))
                                            except TimeoutException:
                                                pass
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
                                _pedir_navegacao_manual(proxima_pagina_num)

                    except KeyboardInterrupt:
                        raise
                    except Exception as e:
                        print(f"\n⚠️  Erro ao processar página {pagina}:")
                        print(f"    Tipo: {type(e).__name__}")
                        print(f"    Mensagem: {str(e)}")
                        if str(e).strip():
                            print(f"    Traceback:\n{traceback.format_exc()}")
                        continue

            except KeyboardInterrupt:
                raise
            except Exception as e:
                print(f"\n⚠️  Erro ao processar {id_elemento} ({categoria}):")
                print(f"    Tipo: {type(e).__name__}")
                print(f"    Mensagem: {str(e)}")
                traceback.print_exc()
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
            print(f"Processo {status.lower()}. Total de {total_salvos} serviços salvos no banco.")
        else:
            print("Nenhum dado de serviço foi coletado.")


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
        #print("[→] Iniciando coleta de insumos\n")
        #print("-" * 70 + "\n")
        #exportar_insumos(navegador)
        #print("\n" + "-" * 70)
        #print("[✓] Insumos finalizados!\n")

        print("[→] Iniciando coleta de serviços e composições\n")
        print("-" * 70 + "\n")
        exportar_servicos(navegador)
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


main()