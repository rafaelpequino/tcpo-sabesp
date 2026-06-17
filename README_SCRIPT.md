# O que o script `extrair_tcpo.py` faz — passo a passo técnico

Este documento descreve o que o script executa internamente, em ordem, do início ao fim.

---

## Visão geral

O script abre um navegador Chromium, faz login no TCPOweb, navega por toda a árvore de categorias, entra em cada item para coletar os dados de detalhe (código, tipo, sub-itens, preços) e salva tudo em um arquivo Excel organizado por abas.

---

## Etapas de execução

### 1 — Credenciais

- Lê `TCPO_USUARIO` e `TCPO_SENHA` do arquivo `.env`.
- Se algum dos campos estiver vazio, solicita no terminal (senha digitada sem eco).

---

### 2 — Login

- Abre o navegador Chromium (visível, `HEADLESS = False`).
- Navega para `https://tcpoweb.pini.com.br`.
- Preenche os campos de usuário e senha e clica em **Entrar**.
- Após o login, verifica se apareceu o modal **"Acesso negado"**:
  - Se sim → clica OK e encerra o script imediatamente (`SystemExit`).
  - Se não → clica no link **"Composições e preços"** para entrar no módulo correto.

---

### 3 — Identificação dos frames

- A página usa **frameset** (dois frames lado a lado).
- O script identifica qual frame é o **menu** (URL contém `Menu.aspx`) e qual é o **conteúdo** (todos os outros frames não-vazios).

---

### 4 — Seleção do banco de dados

- No frame do menu, localiza o `<select>` de bancos.
- Seleciona a opção **"TCPO PINI"** por texto exato (ou parcial, ou índice como fallback).
- Aguarda o recarregamento da página após a seleção.

---

### 5 — Filtros de busca avançada

- Clica no botão verde **"Busca Avançada"** (`#ctl00_MainContent_imgBtnBuscaAvancada`).
- Marca o checkbox **"Procurar somente na BASE SELECIONADA"**.
- Essa configuração garante que só apareçam itens do banco TCPO PINI.

---

### 6 — Expansão da árvore de navegação

- Localiza o container da `TreeView` ASP.NET no menu.
- Percorre todos os `<img>` cujo `onclick` contém `TreeView_Toggle`.
- Para cada nó colapsado (filho com `display:none`), clica para expandir.
- Repete o processo em rodadas até não restar nenhum nó colapsado.

---

### 7 — Listagem de categorias

- Após a expansão, percorre todos os `<a>` dentro do container da árvore.
- Filtra fora links de navegação geral (Sair, Login, Home, etc.) e links de toggle.
- Para cada link de categoria, captura:
  - **texto** — nome da categoria (ex: "Canteiro de obras")
  - **href** / **onclick** — como acionar o link
  - **grupo** — texto do nó raiz imediato (ex: "Serviços", "Insumos", "Composições auxiliares")
- Resultado: lista de ~274 categorias.

---

### 8 — Loop por categoria

Para cada categoria da lista:

#### 8.1 — Clicar na categoria
- Busca o `<a>` correspondente no DOM do menu pelo texto + href/onclick.
- Se não encontrar (árvore colapsada após navegação): **re-expande** a árvore e tenta novamente.
- Fallback final: executa o `href`/`onclick` diretamente via `eval()`.

#### 8.2 — Extrair tabela de resultados
- No frame de conteúdo, localiza a tabela que tem "item"/"base" + "descrição"/"unidade" no cabeçalho.
- Extrai todas as linhas, capturando também o `href` do link de detalhe de cada item.
- **Paginação automática**: se existir um link `>` ou `Próxima`, clica e repete até a última página.

#### 8.3 — Coletar detalhe de cada item
Para cada linha da tabela:

1. **Navega para o detalhe**:
   - Se o link é `javascript:__doPostBack(...)`: localiza o `<a>` pelo href exato e chama `click()` no contexto do frame; aguarda o link **"Voltar para:"** aparecer para confirmar que a página de detalhe carregou.
   - Se for URL normal: `frame.goto()`.

2. **Extrai o cabeçalho** via regex sobre o `innerText` da página:
   - Código, Tipo (SERVIÇO COMPOSTO, MATERIAL, MÃO DE OBRA, etc.), Unidade, BIM
   - Descrição
   - Região de preços (do `<select>` selecionado)
   - Data de preços
   - Sem Taxas (R$), Com Taxas (R$), LS (%), BDI (%)

3. **Extrai sub-itens** da tabela de composição (cabeçalho com "código" + "descrição"):
   - Código, Descrição, Un, Class, Coef, Preço unitário (R$) sem taxas, Total (R$) sem taxas, Consumo

4. **Volta para a listagem** sem destruir o frameset:
   - Inspeciona o link "Voltar para:" para ver se tem `target="_top"` (que recarregaria tudo).
   - Se sim: faz `frame.goto(href)` direto no content-frame, sem propagar para o frameset.
   - Se for `javascript:`: executa via `eval()` no frame.

---

### 9 — Montagem dos registros

Cada item gera **uma ou mais linhas** no dataset, identificadas pela coluna `Tipo_Linha`:

| `Tipo_Linha` | Quando | Colunas Sub-* |
|---|---|---|
| `Composição` | Item com sub-itens | Vazias |
| `Insumo/Serviço` | Item sem sub-itens | Vazias |
| `Sub-item` | Cada componente de uma composição | Preenchidas |

Todos os registros têm também o campo `Grupo` (Serviços / Insumos / Composições auxiliares).

---

### 10 — Logout (sempre executado)

- O `finally` garante que o logout ocorra mesmo em caso de erro ou Ctrl+C.
- Clica no link **"Sair"** por seletor (`a[href*='Sair']`, depois por texto).

---

### 11 — Geração do arquivo Excel

- Cria o DataFrame com todos os registros coletados.
- Monta a coluna auxiliar `Aba` mapeando `Grupo` → nome da aba:
  - texto com "servi" → `Serviços`
  - texto com "insumo" → `Insumos`
  - texto com "compos" → `Composições auxiliares`
  - outros → `Outros`
- Salva um arquivo `.xlsx` com uma aba por grupo; colunas `Aba` e `Grupo` são descartadas.
- O arquivo é salvo em `extraidos/tcpo_YYYY-MM-DD_HH-MM-SS.xlsx`.

---

## Colunas do Excel gerado

| Coluna | Descrição |
|---|---|
| `Tipo_Linha` | `Composição`, `Insumo/Serviço` ou `Sub-item` |
| `Banco` | Sempre "TCPO PINI" |
| `Categoria` | Nome da categoria da árvore (ex: "Canteiro de obras") |
| `Base` | Código da base (coluna 1 da tabela de listagem) |
| `Item` | Código do item (ex: `02.101.000057.SER`) |
| `Descrição` | Descrição do item pai |
| `Unidade` | Unidade de medida do item pai |
| `Código` | Código extraído da página de detalhe |
| `Tipo` | Tipo (SERVIÇO COMPOSTO, MATERIAL, MÃO DE OBRA, etc.) |
| `BIM` | Código BIM |
| `Região` | Região de preços selecionada |
| `Data Preços` | Data da última atualização de preços |
| `Sem Taxas (R$)` | Valor sem leis sociais e BDI |
| `Com Taxas (R$)` | Valor com leis sociais e BDI |
| `LS (%)` | Percentual de Leis Sociais |
| `BDI (%)` | Percentual de BDI |
| `Sub-Código` | Código do sub-item (linhas `Sub-item`) |
| `Sub-Descrição` | Descrição do sub-item |
| `Sub-Un` | Unidade do sub-item |
| `Sub-Class` | Classificação (MOD, MAT, EQP, etc.) |
| `Sub-Coef` | Coeficiente de utilização |
| `Sub-Preço Unit (R$)` | Preço unitário sem taxas |
| `Sub-Total (R$)` | Total sem taxas |
| `Sub-Consumo` | Consumo |

---

## Parâmetros configuráveis (topo do script)

| Variável | Padrão | Efeito |
|---|---|---|
| `HEADLESS` | `False` | `True` oculta o navegador (mais rápido) |
| `SLOW_MO` | `150` ms | Pausa entre ações — aumente se o site estiver lento |
| `TIMEOUT` | `25000` ms | Tempo máximo para carregamento de cada página |
| `OUTPUT_DIR` | `extraidos` | Pasta onde os arquivos Excel são salvos |
