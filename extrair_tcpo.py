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

    # Página intermediária: clicar em "Composições e preços"
    _navegar_composicoes(page)


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

_MENU_KEYWORDS    = ["TreeView", "Canteiro", "Serviços Iniciais", "treeMenu",
                     "menuLateral", "arvore", "navegacao"]
_CONTENT_KEYWORDS = ["Mostrando:", "ctl00_Content", "gridDados",
                     "gridResultado", "tblComposicoes"]


def identificar_frames(page):
    """
    Retorna (frame_menu, frame_conteudo).
    Se a página não usar framesets/iframes, devolve (page, page).
    """
    frame_menu     = None
    frame_conteudo = None

    for f in page.frames:
        try:
            html = f.content()
        except Exception:
            continue
        if frame_menu is None and any(k in html for k in _MENU_KEYWORDS):
            frame_menu = f
        if frame_conteudo is None and any(k in html for k in _CONTENT_KEYWORDS):
            frame_conteudo = f

    return frame_menu or page, frame_conteudo or page


# ── Árvore de navegação ───────────────────────────────────────────────────────

def expandir_arvore(frame_menu):
    """Clica em todos os ícones de expansão (+) até não restar nenhum."""
    print("  Expandindo nós da árvore...", end="", flush=True)
    iteracoes = 0
    while iteracoes < 60:
        nos = frame_menu.locator(
            "img[src*='plus'], img[src*='Expand'], img[src*='expand'], "
            "img[src*='collapsed'], a[title*='Expandir'], "
            "span.TreeExpand, td.TreeExpand"
        ).all()
        clicou = False
        for no in nos:
            try:
                if no.is_visible():
                    no.click()
                    time.sleep(0.35)
                    clicou = True
            except Exception:
                continue
        if not clicou:
            break
        iteracoes += 1
    print(f" pronto ({iteracoes} rodadas).")


# Padrão ASP.NET TreeView para nós de expand/collapse: __doPostBack('Ctrl','s0\...')
_EXPAND_PATTERN = re.compile(r"__doPostBack\([^)]*'[sSlLcC]\d+\\", re.IGNORECASE)


def listar_links_categorias(frame_menu):
    """
    Retorna lista de dicts {texto, href, onclick} para todos os links
    que parecem ser folhas da árvore (navegam para conteúdo).
    """
    links  = []
    vistos = set()

    for el in frame_menu.locator("a").all():
        try:
            texto   = el.inner_text().strip()
            href    = el.get_attribute("href")    or ""
            onclick = el.get_attribute("onclick") or ""

            if not texto or len(texto) < 2:
                continue

            # Pular nós de expansão/colapso do ASP.NET TreeView
            if _EXPAND_PATTERN.search(onclick):
                continue

            # Pular âncoras sem destino
            if not href and not onclick:
                continue

            chave = (texto, href, onclick)
            if chave in vistos:
                continue
            vistos.add(chave)

            links.append({"texto": texto, "href": href, "onclick": onclick})
        except Exception:
            continue

    return links


# ── Extração de tabela ────────────────────────────────────────────────────────

def _localizar_tabela(frame):
    """Encontra a tabela de composições pelo cabeçalho."""
    for t in frame.locator("table").all():
        try:
            header = t.locator("tr:first-child").inner_text().lower()
            if any(p in header for p in ["descrição", "item", "base", "unidade"]):
                return t
        except Exception:
            continue
    return None


def extrair_dados_categoria(frame_conteudo):
    """Extrai todas as linhas da categoria atual, incluindo paginação."""
    todos  = []
    pagina = 1

    while True:
        try:
            frame_conteudo.wait_for_load_state("networkidle", timeout=8_000)
        except Exception:
            pass
        time.sleep(0.4)

        tabela = _localizar_tabela(frame_conteudo)
        if not tabela:
            break

        linhas = tabela.locator("tr").all()
        n = 0
        for linha in linhas:
            cols = linha.locator("td").all()
            if len(cols) >= 2:
                vals = [c.inner_text().strip() for c in cols]
                if any(vals):
                    todos.append(vals)
                    n += 1

        print(f"      Pág {pagina}: {n} itens")

        # Verificar se há próxima página
        prox = frame_conteudo.locator(
            "a:text-is('Próxima'), a:text-is('>>'), a:text-is('>'), "
            "a[title*='Próxima'], a[title*='próxima'], input[value*='róxima']"
        )
        if prox.count() > 0 and prox.first.is_visible():
            prox.first.click()
            pagina += 1
            time.sleep(0.6)
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
    Abre a busca avançada (botão verde com ícone de chave) e marca
    a opção 'Procurar somente na BASE SELECIONADA'.
    """
    print("  Configurando filtros de busca avançada...")

    # Clicar no botão verde de busca avançada
    btn_seletores = [
        "img[src*='avancada']", "img[src*='Avancada']", "img[src*='advanced']",
        "img[src*='chave']",    "img[src*='wrench']",   "img[src*='tool']",
        "a[title*='vançada']",  "a[title*='Advanced']",
        "input[title*='vançada']",
        # botão verde identificado pela cor/classe
        "button.green", "a.btnAvancada", "#btnBuscaAvancada",
        # fallback: imagem logo após o campo de busca
        "#imgBuscaAvancada", "img.buscaAvancada",
    ]

    alvos = [page] + list(page.frames)
    clicou_btn = False
    for alvo in alvos:
        for sel in btn_seletores:
            try:
                loc = alvo.locator(sel)
                if loc.count() > 0 and loc.first.is_visible(timeout=2_000):
                    loc.first.click()
                    time.sleep(0.8)
                    clicou_btn = True
                    break
            except Exception:
                continue
        if clicou_btn:
            break

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

        # 2. Identificar estrutura
        print("\n[2] Identificando layout da página...")
        time.sleep(2)
        frame_menu, _ = identificar_frames(page)
        print(f"    Frame menu identificado: {frame_menu.name or 'página principal'}")

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

                # Localizar o link pelo texto exato para evitar referências obsoletas
                el = fm.locator("a").filter(
                    has_text=re.compile(r"^\s*" + re.escape(texto) + r"\s*$")
                ).first

                try:
                    el.scroll_into_view_if_needed(timeout=3_000)
                except Exception:
                    pass

                if not el.is_visible():
                    print("      → Elemento não visível, pulando.")
                    continue

                el.click()
                time.sleep(SLOW_MO / 1000 + 0.5)

                _, fc = identificar_frames(page)
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
