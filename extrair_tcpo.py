#!/usr/bin/env python3
"""
Extrator TCPOweb → Excel
========================
Automatiza a extração de todas as composições de preços do sistema
TCPO Web (tcpoweb.pini.com.br) para um arquivo Excel.

Pré-requisitos:
    pip install -r requirements.txt
    playwright install chromium

Credenciais (escolha uma opção):
    a) Crie um arquivo .env com:
           TCPO_USUARIO=seu_usuario
           TCPO_SENHA=sua_senha
    b) Digite quando o script solicitar no terminal.

Uso:
    python extrair_tcpo.py
"""

import os
import re
import time
import getpass

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ── Configurações ─────────────────────────────────────────────────────────────
BASE_URL   = "https://tcpoweb.pini.com.br"
OUTPUT_FILE = "tcpo_composicoes.xlsx"
HEADLESS   = False    # False = ver o navegador em tempo real; True = oculto
SLOW_MO    = 150      # ms entre ações (aumente para 300 se o site for lento)
TIMEOUT    = 25_000   # ms timeout para carregamentos de página
# ─────────────────────────────────────────────────────────────────────────────


# ── Credenciais ───────────────────────────────────────────────────────────────

def obter_credenciais():
    """Lê credenciais do .env ou solicita interativamente."""
    usuario = os.getenv("TCPO_USUARIO", "").strip()
    senha   = os.getenv("TCPO_SENHA",   "").strip()
    if not usuario:
        usuario = input("Usuário TCPO: ").strip()
    if not senha:
        senha = getpass.getpass("Senha TCPO: ")
    return usuario, senha


# ── Login ─────────────────────────────────────────────────────────────────────

def _preencher(alvo, seletores, valor):
    """Preenche o primeiro campo visível encontrado dentre os seletores."""
    for sel in seletores:
        try:
            loc = alvo.locator(sel)
            if loc.count() > 0 and loc.first.is_visible():
                loc.first.fill(valor)
                return True
        except Exception:
            continue
    return False


def _clicar(alvo, seletores):
    """Clica no primeiro elemento visível encontrado dentre os seletores."""
    for sel in seletores:
        try:
            loc = alvo.locator(sel)
            if loc.count() > 0 and loc.first.is_visible():
                loc.first.click()
                return True
        except Exception:
            continue
    return False


def fazer_login(page, usuario, senha):
    """Abre o sistema e realiza autenticação."""
    page.goto(f"{BASE_URL}/menu.aspx", wait_until="domcontentloaded", timeout=TIMEOUT)
    time.sleep(2)

    # Verifica se o formulário está dentro de um iframe
    alvo = page
    for frame in page.frames:
        try:
            if frame.locator("input[type='password']").count() > 0:
                alvo = frame
                break
        except Exception:
            continue

    _preencher(alvo, [
        "#txtUsuario", "#txtLogin", "#txtEmail",
        "input[name*='suario']", "input[name*='ser']", "input[name*='ogin']",
        "input[type='text']:first-of-type",
    ], usuario)

    _preencher(alvo, [
        "#txtSenha", "#txtPassword",
        "input[name*='enha']", "input[name*='ass']",
        "input[type='password']",
    ], senha)

    ok = _clicar(alvo, [
        "#btnEntrar", "#btnLogin", "#btnOk", "#btn_login",
        "input[type='submit']", "button[type='submit']",
        "input[value*='ntrar']", "a:has-text('Entrar')",
    ])
    if not ok:
        input("\n[!] Botão de login não encontrado automaticamente.\n"
              "    Faça o login manualmente no navegador e pressione Enter aqui para continuar...")

    page.wait_for_load_state("networkidle", timeout=TIMEOUT)
    time.sleep(2)

    # Verificar se apareceu o aviso de "Acesso negado / usuário já logado"
    _verificar_acesso_negado(page)

    # Página intermediária: clicar em "Composições e preços"
    _navegar_composicoes(page)


def _verificar_acesso_negado(page):
    """
    Detecta o modal 'Acesso negado. Este usuário já está utilizando esta
    aplicação em outro navegador...' e encerra o script com mensagem clara.
    """
    alvos = [page] + list(page.frames)
    for alvo in alvos:
        try:
            aviso = alvo.locator("text=Acesso negado")
            if aviso.count() > 0 and aviso.first.is_visible(timeout=2_000):
                # Tenta clicar em OK para fechar o modal antes de sair
                try:
                    alvo.locator("input[value='OK'], button:has-text('OK'), a:has-text('OK')").first.click()
                except Exception:
                    pass
                print("\n" + "=" * 58)
                print("  ACESSO NEGADO")
                print("  Este usuário já está logado em outro navegador")
                print("  ou computador. Feche a outra sessão e tente")
                print("  novamente.")
                print("=" * 58 + "\n")
                raise SystemExit(1)
        except SystemExit:
            raise
        except Exception:
            continue


def _navegar_composicoes(page):
    """
    Após o login o sistema exibe uma tela de boas-vindas/módulos.
    Localiza e clica no botão/link 'Composições e preços'.
    """
    seletores = [
        "a:has-text('Composições e preços')",
        "a:has-text('Composicoes e precos')",
        "td:has-text('Composições e preços')",
        "div:has-text('Composições e preços')",
        "span:has-text('Composições e preços')",
        "input[value*='omposições']",
    ]
    for sel in seletores:
        try:
            loc = page.locator(sel).first
            if loc.count() > 0 and loc.is_visible(timeout=3_000):
                print("    Clicando em 'Composições e preços'...")
                loc.click()
                page.wait_for_load_state("networkidle", timeout=TIMEOUT)
                time.sleep(2)
                return
        except Exception:
            continue

    # Tenta em iframes
    for frame in page.frames:
        for sel in seletores:
            try:
                loc = frame.locator(sel).first
                if loc.count() > 0 and loc.is_visible(timeout=2_000):
                    print("    Clicando em 'Composições e preços' (iframe)...")
                    loc.click()
                    page.wait_for_load_state("networkidle", timeout=TIMEOUT)
                    time.sleep(2)
                    return
            except Exception:
                continue

    print("    [!] Botão 'Composições e preços' não encontrado automaticamente.")
    input("    Clique manualmente em 'Composições e preços' e pressione Enter aqui para continuar...")


# ── Identificar frames ────────────────────────────────────────────────────────

# Fragmentos de URL que identificam cada frame pelo endereço (mais confiável
# que varrer o HTML completo, que pode retornar a página inteira por engano).
_URL_MENU    = ["TreeView", "treeview", "MenuLateral", "menuLateral"]
_URL_CONTENT = ["Pesq", "pesq", "Resultado", "resultado", "Composicao",
                "composicao", "Grid", "grid"]

# Palavras-chave no HTML como fallback (só usadas se a URL não resolver)
_HTML_MENU    = ["ctl00_TreeView", "TreeView1", "Canteiro de obras",
                 "Serviços Iniciais", "treeMenu"]
_HTML_CONTENT = ["Mostrando:", "ctl00_MainContent_grid", "gridDados",
                 "gridResultado", "tblComposicoes"]


def identificar_frames(page):
    """
    Retorna (frame_menu, frame_conteudo).
    Prioriza detecção por URL do frame; usa HTML como fallback.
    Se nenhum iframe for encontrado, devolve (page, page).
    """
    frame_menu     = None
    frame_conteudo = None

    # Ignora o frame principal (index 0) que é a página toda
    frames_filhos = [f for f in page.frames if f != page.main_frame]

    # 1ª tentativa: URL do frame
    for f in frames_filhos:
        url = f.url or ""
        if frame_menu is None and any(k in url for k in _URL_MENU):
            frame_menu = f
        if frame_conteudo is None and any(k in url for k in _URL_CONTENT):
            frame_conteudo = f

    # 2ª tentativa: conteúdo HTML
    if frame_menu is None or frame_conteudo is None:
        for f in frames_filhos:
            try:
                html = f.content()
            except Exception:
                continue
            if frame_menu is None and any(k in html for k in _HTML_MENU):
                frame_menu = f
            if frame_conteudo is None and any(k in html for k in _HTML_CONTENT):
                frame_conteudo = f

    return frame_menu or page, frame_conteudo or page


# ── Árvore de navegação ───────────────────────────────────────────────────────

# Seletor CSS para o container raiz da árvore ASP.NET no layout PINI.
# O TreeView é renderizado dentro de um <div> com id contendo 'TreeView'.
# Usar este container garante que NADA fora dele (header, nav, Vizca) seja tocado.
_SEL_TREE_CONTAINER = (
    "[id*='TreeView'], [id*='treeView'], "
    "[id*='TreeMenu'], [id*='treeMenu'], "
    "[id*='menuLateral'], [id*='MenuLateral']"
)

# Padrão ASP.NET TreeView para nós de expand/collapse: __doPostBack('t','sN\'...')
_EXPAND_ONCLICK = re.compile(
    r"TreeView_Toggle|__doPostBack\('[^']*','[tT]\d+\\", re.IGNORECASE
)
# Padrão de link folha (categoria): __doPostBack com parâmetro de nó simples
_LEAF_ONCLICK = re.compile(
    r"__doPostBack\s*\(", re.IGNORECASE
)
# Itens de navegação a excluir (header, user menu, etc.)
_NAV_EXCLUSOES = re.compile(
    r"vizca|sair|logout|entrar|login|info|ajuda|help|home|assine|contato",
    re.IGNORECASE
)


def _tree_container(frame):
    """
    Retorna o locator do container da árvore dentro do frame.
    Se não encontrar o container específico, retorna None — não usa
    o frame inteiro para evitar pegar elementos do cabeçalho.
    """
    loc = frame.locator(_SEL_TREE_CONTAINER)
    if loc.count() > 0:
        return loc.first
    return None


def expandir_arvore(frame_menu):
    """
    Expande todos os nós colapsados do ASP.NET TreeView usando JavaScript.

    No ASP.NET TreeView cada botão de expand/collapse tem um <img> com id
    no padrão '{prefix}t{nodeKey}' e o div de filhos correspondente tem id
    '{prefix}n{nodeKey}Nodes'. Se esse div estiver display:none, o nó está
    colapsado e clicamos na imagem para abrir.
    """
    print("  Expandindo nós da árvore...", end="", flush=True)
    iteracoes = 0
    while iteracoes < 60:
        expandiu = frame_menu.evaluate("""
            () => {
                let count = 0;
                // Todas as imagens de toggle do TreeView
                const imgs = document.querySelectorAll("img[onclick*='TreeView_Toggle']");
                for (const img of imgs) {
                    if (!img.id) continue;
                    // Transforma o id da imagem no id do div de filhos:
                    //   {prefix}t{key}  →  {prefix}n{key}Nodes
                    // Cobre nós simples (t0) e aninhados (t0_0, t0_0_1, ...)
                    const divId = img.id.replace(/t(\\d[\\d_]*)$/, 'n$1Nodes');
                    if (divId === img.id) continue; // padrão não casou
                    const nodesDiv = document.getElementById(divId);
                    if (nodesDiv && nodesDiv.style.display === 'none') {
                        img.click();
                        count++;
                    }
                }
                return count;
            }
        """)
        if not expandiu:
            break
        time.sleep(0.6)
        iteracoes += 1
    print(f" pronto ({iteracoes} rodadas).")


def listar_links_categorias(frame_menu):
    """
    Retorna lista de dicts {texto, href, onclick} para todos os links
    de categorias da árvore usando JavaScript.
    O JS percorre apenas o container do TreeView e retorna os links
    que NÃO são de expansão/colapso e NÃO são de navegação de sistema.
    """
    links_js = frame_menu.evaluate("""
        () => {
            // Localizar o container do TreeView pelo id
            const container = (
                document.querySelector("[id*='TreeView']") ||
                document.querySelector("[id*='TreeMenu']") ||
                document.querySelector("[id*='menuLateral']")
            );
            if (!container) return [];

            const NAV = /vizca|sair|logout|entrar|login|home|assine|contato|info/i;
            const result = [];
            const seen   = new Set();

            for (const a of container.querySelectorAll('a')) {
                const texto   = (a.innerText || '').trim();
                const href    = a.getAttribute('href')    || '';
                const onclick = a.getAttribute('onclick') || '';

                if (!texto || texto.length < 2)      continue;
                if (NAV.test(texto))                  continue;
                // Pular links de expand/collapse do TreeView
                if (/TreeView_Toggle/i.test(href))    continue;
                if (/TreeView_Toggle/i.test(onclick)) continue;
                // Pular links externos
                if (href.startsWith('http'))          continue;
                // Precisam ter algum destino
                if (!href && !onclick)                continue;

                const chave = texto + '|' + href + '|' + onclick;
                if (seen.has(chave)) continue;
                seen.add(chave);

                result.push({texto, href, onclick});
            }
            return result;
        }
    """)
    return links_js or []


# ── Extração de tabela ────────────────────────────────────────────────────────

def _frame_conteudo(page, frame_menu):
    """
    Retorna o frame que contém os resultados (o que NÃO é o menu e NÃO é o
    frame principal). Após um clique numa categoria, é neste frame que a
    tabela de resultados é carregada.
    """
    menu_url = getattr(frame_menu, "url", "") or ""
    for f in page.frames:
        if f is page.main_frame:
            continue
        url = f.url or ""
        if url == menu_url or not url or url == "about:blank":
            continue
        return f
    return page


def extrair_dados_categoria(frame_conteudo):
    """
    Extrai TODOS os itens da tabela de resultados, incluindo paginação.
    Usa JavaScript para localizar a tabela correta (ignora o painel de
    filtros que tem texto de checkboxes) e coleta todas as linhas de dados.
    """
    todos  = []
    pagina = 1

    while True:
        try:
            frame_conteudo.wait_for_load_state("networkidle", timeout=10_000)
        except Exception:
            pass
        time.sleep(0.5)

        linhas = frame_conteudo.evaluate("""
            () => {
                // Texto que identifica o painel de filtros — deve ser IGNORADO
                const FILTRO = /Procurar somente|FRASE EXATA|Listar somente COMP|Listar somente INSU/i;

                // Encontrar a tabela de resultados:
                // Critério principal: cabeçalho com Base + Item + Descrição ou Unidade
                // Critério secundário: tabela com mais linhas que não seja filtro
                let melhorTabela = null;
                let melhorLinhas = 0;

                for (const t of document.querySelectorAll('table')) {
                    // Pular tabelas que são o painel de filtros
                    const txt = (t.innerText || '');
                    if (FILTRO.test(txt) && !/TCPO|3R |Código/i.test(txt)) continue;

                    const rows = t.querySelectorAll('tr');
                    if (rows.length < 2) continue;

                    const header = (rows[0].innerText || '').toLowerCase();
                    const isResultTable = (
                        (header.includes('item') || header.includes('base')) &&
                        (header.includes('descri') || header.includes('unidade'))
                    );

                    if (isResultTable) {
                        melhorTabela = t;
                        break;  // encontrou pelo cabeçalho — parar aqui
                    }

                    // Fallback: tabela com mais linhas sem ser filtro
                    if (rows.length > melhorLinhas) {
                        melhorLinhas = rows.length;
                        melhorTabela = t;
                    }
                }

                if (!melhorTabela) return [];

                const result = [];
                const rows = melhorTabela.querySelectorAll('tr');
                // Pular linha de cabeçalho (row 0) e coletar todas as linhas de dados
                for (let i = 1; i < rows.length; i++) {
                    const cells = rows[i].querySelectorAll('td');
                    if (cells.length < 2) continue;
                    const vals = Array.from(cells).map(c => (c.innerText || '').trim());
                    // Pular linhas de paginação (contêm só números ou símbolos)
                    if (vals.every(v => v === '' || /^[\\d\\s<>»«\\.]+$/.test(v))) continue;
                    if (vals.some(v => v !== '')) result.push(vals);
                }
                return result;
            }
        """)

        n = len(linhas) if linhas else 0
        if n > 0:
            todos.extend(linhas)
            print(f"      Pág {pagina}: {n} itens")
        else:
            if pagina == 1:
                print(f"      Pág {pagina}: 0 itens (sem tabela de resultados)")
            break

        # ── Paginação ──────────────────────────────────────────────────────
        # O GridView ASP.NET renderiza os links de página dentro da tabela.
        # Tentamos via JS para pegar qualquer padrão de paginação.
        proxima = frame_conteudo.evaluate("""
            () => {
                // Links de próxima página: '>' ou '>>' ou 'Próxima' ou número de página
                const candidatos = Array.from(document.querySelectorAll('a'));
                for (const a of candidatos) {
                    const t = (a.innerText || '').trim();
                    const title = (a.title || '').toLowerCase();
                    if (t === '>' || t === '>>' || t === '»' ||
                        /próxima|próx|next/i.test(t) ||
                        /próxima|next/i.test(title)) {
                        a.click();
                        return true;
                    }
                }
                return false;
            }
        """)

        if proxima:
            pagina += 1
            time.sleep(1.0)
        else:
            break

    return todos


# ── Banco de dados (sempre TCPO PINI) ───────────────────────────────────────

BANCO_ALVO = "TCPO PINI"   # único banco que deve ser usado


def selecionar_banco_tcpo_pini(frame_menu):
    """
    Garante que o dropdown de bases esteja selecionado em 'TCPO PINI'.
    Tenta por texto exato, depois por texto parcial, depois por índice.
    """
    try:
        sel = frame_menu.locator("select").first
        if sel.count() == 0:
            print("  [!] Dropdown de banco não encontrado.")
            return

        opcoes = sel.locator("option").all()
        valor_alvo = None
        for op in opcoes:
            txt = op.inner_text().strip()
            if BANCO_ALVO.lower() in txt.lower():
                valor_alvo = op.get_attribute("value")
                break

        if valor_alvo is not None:
            sel.select_option(value=valor_alvo)
            print(f"  Banco selecionado: '{BANCO_ALVO}'")
        else:
            # Tenta selecionar pela label mesmo assim
            try:
                sel.select_option(label=BANCO_ALVO)
                print(f"  Banco selecionado por label: '{BANCO_ALVO}'")
            except Exception:
                disponiveis = [o.inner_text().strip() for o in opcoes]
                print(f"  [!] '{BANCO_ALVO}' não encontrado. Opções: {disponiveis}")
                print(f"  [!] Mantendo a seleção atual do dropdown.")
    except Exception as e:
        print(f"  [!] Erro ao selecionar banco: {e}")


# ── Filtros de busca avançada ─────────────────────────────────────────────────

def configurar_filtros_busca(page):
    """
    Abre a busca avançada (botão verde com ícone de chave, ao lado do campo
    de busca) e marca a opção 'Procurar somente na BASE SELECIONADA'.

    IMPORTANTE: os seletores são propositalmente escopados ao container do
    campo de busca para NÃO confundir com links de navegação (ex: menu do
    usuário 'Vizca') que também aparecem na página.
    """
    print("  Configurando filtros de busca avançada...")

    # ID exato do botão verde inspecionado no sistema
    SEL_BTN = "#ctl00_MainContent_imgBtnBuscaAvancada"

    alvos = [page] + list(page.frames)
    clicou_btn = False

    for alvo in alvos:
        try:
            loc = alvo.locator(SEL_BTN)
            if loc.count() > 0 and loc.first.is_visible(timeout=3_000):
                loc.first.click()
                time.sleep(0.8)
                clicou_btn = True
                break
        except Exception:
            pass

    if not clicou_btn:
        print("  [!] Botão de busca avançada não encontrado; tentando localizar o painel diretamente.")

    # Marcar checkbox 'Procurar somente na BASE SELECIONADA'
    cb_seletores = [
        "input[type='checkbox'][id*='BaseSelecionada']",
        "input[type='checkbox'][id*='baseSelecionada']",
        "input[type='checkbox'][name*='BaseSelecionada']",
        "label:has-text('BASE SELECIONADA') input",
        "label:has-text('BASE SELECIONADA') ~ input",
    ]
    # Busca também pelo texto do label próximo ao checkbox
    marcou = False
    for alvo in alvos:
        # Tenta seletores diretos
        for sel in cb_seletores:
            try:
                cb = alvo.locator(sel)
                if cb.count() > 0 and cb.first.is_visible(timeout=2_000):
                    if not cb.first.is_checked():
                        cb.first.check()
                    print("  ✓ 'Procurar somente na BASE SELECIONADA' marcado.")
                    marcou = True
                    break
            except Exception:
                continue
        if marcou:
            break

        # Fallback: procura todos os checkboxes e verifica o label próximo
        if not marcou:
            try:
                for cb in alvo.locator("input[type='checkbox']").all():
                    try:
                        # pega o texto do label associado ou do elemento pai
                        label_text = ""
                        cb_id = cb.get_attribute("id")
                        if cb_id:
                            lbl = alvo.locator(f"label[for='{cb_id}']")
                            if lbl.count() > 0:
                                label_text = lbl.first.inner_text()
                        if not label_text:
                            label_text = cb.evaluate(
                                "el => el.parentElement ? el.parentElement.innerText : ''"
                            )
                        if "BASE SELECIONADA" in label_text.upper():
                            if not cb.is_checked():
                                cb.check()
                            print("  ✓ 'Procurar somente na BASE SELECIONADA' marcado (fallback).")
                            marcou = True
                            break
                    except Exception:
                        continue
            except Exception:
                continue
        if marcou:
            break

    if not marcou:
        print("  [!] Checkbox 'BASE SELECIONADA' não encontrado; verifique manualmente.")

    time.sleep(0.5)


# ── Logout ────────────────────────────────────────────────────────────────────

def fazer_logout(page):
    """
    Clica em 'Sair' no canto superior direito para encerrar a sessão.
    Sempre chamado no bloco finally — executa mesmo em caso de erro ou Ctrl+C.
    """
    print("\n  Fazendo logout...")
    try:
        alvos = [page] + list(page.frames)
        for alvo in alvos:
            for sel in [
                "#lnkSair", "#linkSair", "a[id*='Sair']", "a[id*='sair']",
                "a[href*='Sair']", "a[href*='sair']",
                "a[href*='logout']", "a[href*='Logout']",
                "a:has-text('Sair')",
            ]:
                try:
                    loc = alvo.locator(sel)
                    if loc.count() > 0 and loc.first.is_visible(timeout=2_000):
                        loc.first.click()
                        page.wait_for_load_state("domcontentloaded", timeout=8_000)
                        print("  \u2713 Logout realizado.")
                        return
                except Exception:
                    continue
        print("  [!] Link 'Sair' n\u00e3o encontrado — feche a sess\u00e3o manualmente se necess\u00e1rio.")
    except Exception as e:
        print(f"  [!] Erro ao fazer logout: {e}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 58)
    print("   Extrator TCPOweb → Excel")
    print("=" * 58)

    usuario, senha = obter_credenciais()
    todos_registros = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
        ctx  = browser.new_context(viewport={"width": 1400, "height": 900})
        page = ctx.new_page()

        # 1. Login
        print("\n[1] Fazendo login...")
        fazer_login(page, usuario, senha)
        print("    Login OK.")

        try:
            # 2. Identificar estrutura
            print("\n[2] Identificando layout da página...")
            time.sleep(2)
            frame_menu, _ = identificar_frames(page)
            frame_url = getattr(frame_menu, "url", "") or ""
            print(f"    Frame menu: {frame_url or 'página principal'}")
            container = _tree_container(frame_menu)
            print(f"    Container da árvore: {'encontrado' if container else 'NÃO ENCONTRADO — inspecione o id do div da árvore'}")

            # 3. Selecionar banco TCPO PINI e configurar filtros
            print(f"\n[3] Selecionando banco '{BANCO_ALVO}'...")
            selecionar_banco_tcpo_pini(frame_menu)
            page.wait_for_load_state("networkidle", timeout=TIMEOUT)
            time.sleep(1.5)

            print("\n[4] Configurando filtros de busca avançada...")
            configurar_filtros_busca(page)

            frame_menu, _ = identificar_frames(page)
            banco = {"texto": BANCO_ALVO}

            print(f"\n  ╔══ Banco: {banco['texto']} ══╗")

            # Expandir árvore
            expandir_arvore(frame_menu)

            # Listar categorias
            links = listar_links_categorias(frame_menu)
            print(f"  {len(links)} categorias encontradas.")

            # Processar cada categoria
            processados = set()
            for i, lnk in enumerate(links, 1):
                texto = lnk["texto"]
                chave = (banco["texto"], texto, lnk["href"], lnk["onclick"])
                if chave in processados:
                    continue
                processados.add(chave)

                print(f"\n  [{i}/{len(links)}] {texto}")

                try:
                    fm, fc = identificar_frames(page)

                    href    = lnk["href"]
                    onclick = lnk["onclick"]

                    clicou = fm.evaluate("""
                        ([href, onclick, texto]) => {
                            const container = (
                                document.querySelector("[id*='TreeView']") ||
                                document.querySelector("[id*='TreeMenu']") ||
                                document.querySelector("[id*='menuLateral']")
                            );
                            if (!container) return false;
                            for (const a of container.querySelectorAll('a')) {
                                const aHref    = a.getAttribute('href')    || '';
                                const aOnclick = a.getAttribute('onclick') || '';
                                const aTexto   = (a.innerText || '').trim();
                                if (aTexto === texto &&
                                    (aHref === href || aOnclick === onclick)) {
                                    a.click();
                                    return true;
                                }
                            }
                            return false;
                        }
                    """, [href, onclick, texto])

                    if not clicou:
                        print("      → Link não encontrado no DOM, pulando.")
                        continue

                    time.sleep(SLOW_MO / 1000 + 0.8)

                    fc = _frame_conteudo(page, fm)
                    dados = extrair_dados_categoria(fc)

                    if not dados:
                        print("      → Sem tabela (nó de expansão ou categoria vazia).")
                        continue

                    for linha in dados:
                        while len(linha) < 4:
                            linha.append("")
                        todos_registros.append({
                            "Banco":      banco["texto"],
                            "Categoria":  texto,
                            "Base":       linha[0],
                            "Item":       linha[1],
                            "Descrição":  linha[2],
                            "Unidade":    linha[3] if len(linha) > 3 else "",
                        })

                    print(f"      ✓ {len(dados)} itens adicionados "
                          f"(total acumulado: {len(todos_registros)})")

                except PWTimeout:
                    print("      ✗ Timeout, continuando...")
                except Exception as e:
                    print(f"      ✗ Erro: {type(e).__name__}: {e}")

        finally:
            # Logout SEMPRE executado — mesmo em erro ou Ctrl+C
            fazer_logout(page)
            browser.close()

    # 5. Salvar Excel
    print("\n" + "=" * 58)
    print("Salvando resultados em Excel...")

    if todos_registros:
        df = pd.DataFrame(todos_registros)

        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            # Aba com todos os dados
            df.to_excel(writer, sheet_name="Todos", index=False)

            # Uma aba separada por banco de dados
            for banco_nome, grupo in df.groupby("Banco"):
                aba = str(banco_nome)[:31]          # Excel: máx 31 chars por aba
                grupo.drop(columns=["Banco"]).to_excel(writer, sheet_name=aba, index=False)

        print(f"\n✓ Arquivo salvo:   {OUTPUT_FILE}")
        print(f"✓ Total registros: {len(df):,}")
        print(f"✓ Categorias:      {df['Categoria'].nunique()}")
        print(f"✓ Bancos:          {df['Banco'].nunique()}")
    else:
        print("\n✗ Nenhum dado coletado.")
        print("  Dicas:")
        print("  - Verifique se o login foi bem-sucedido")
        print("  - Tente aumentar SLOW_MO para 300 no topo do script")
        print("  - Execute com HEADLESS = False para acompanhar visualmente")

    print("\nFinalizado!")


if __name__ == "__main__":
    main()
