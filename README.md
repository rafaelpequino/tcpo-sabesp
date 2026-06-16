# Extrator TCPOweb → Excel

Automação para extrair todas as composições de preços do sistema [TCPOweb (PINI)](https://tcpoweb.pini.com.br) e salvar em um arquivo Excel.

---

## Pré-requisitos

- **Python 3.9 ou superior** instalado e disponível no PATH  
  Verifique com: `python --version`

---

## Passo a passo

### 1. Abrir o terminal na pasta do projeto

No Windows Explorer, navegue até a pasta `automacao-tcpo`, clique na barra de endereço, digite `powershell` e pressione Enter.

Ou abra o PowerShell e execute:
```powershell
cd "C:\Users\rafae\OneDrive\Documentos\Sabesp\Sistemas\automacao-tcpo"
```

---

### 2. Instalar as dependências (somente na primeira vez)

```powershell
pip install -r requirements.txt
playwright install chromium
```

> O segundo comando baixa o navegador Chromium que será controlado automaticamente. Pode demorar alguns minutos.

---

### 3. Configurar as credenciais

**Opção A — Arquivo `.env` (recomendado, não precisa digitar toda vez):**

1. Copie o arquivo `.env.example` e renomeie a cópia para `.env`
2. Abra o `.env` e preencha:
   ```
   TCPO_USUARIO=seu_usuario_aqui
   TCPO_SENHA=sua_senha_aqui
   ```

**Opção B — Digitar no terminal a cada execução:**

Deixe o `.env` de lado. O script vai pedir usuário e senha interativamente ao ser iniciado.

---

### 4. Executar o script

```powershell
python extrair_tcpo.py
```

---

## O que acontece durante a execução

O script abre uma janela do navegador Chromium e executa automaticamente as seguintes etapas:

| Etapa | O que faz |
|-------|-----------|
| **[1] Login** | Preenche usuário e senha e entra no sistema |
| **[2] Composições e preços** | Clica no módulo correto na tela inicial |
| **[3] Seleção de banco** | Garante que o dropdown esteja em **TCPO PINI** |
| **[4] Filtros de busca** | Abre a busca avançada (botão verde) e marca **"Procurar somente na BASE SELECIONADA"** |
| **[5] Expansão da árvore** | Clica em todos os nós `+` da navegação lateral até expandir tudo |
| **[6] Extração** | Clica em cada categoria, lê a tabela de resultados e avança páginas automaticamente se houver paginação |
| **[7] Excel** | Salva tudo no arquivo `tcpo_composicoes.xlsx` |

> Você pode acompanhar tudo visualmente no navegador que se abre. Para executar em modo oculto (mais rápido), mude `HEADLESS = False` para `HEADLESS = True` no topo do script.

---

## Resultado gerado

Ao final, será criado o arquivo **`tcpo_composicoes.xlsx`** na mesma pasta do script, com a seguinte estrutura:

### Abas do Excel

| Aba | Conteúdo |
|-----|----------|
| `Todos` | Todos os registros extraídos de todas as categorias |
| `TCPO PINI` | Os mesmos dados, sem a coluna "Banco" |

### Colunas

| Coluna | Exemplo |
|--------|---------|
| `Banco` | TCPO PINI |
| `Categoria` | Canteiro de obras |
| `Base` | TCPO |
| `Item` | 3R 02 54 00 00 00 00 08 |
| `Descrição` | Abrigo provisório de madeira com dois pavimentos... |
| `Unidade` | m² |

---

## Solução de problemas

| Situação | O que fazer |
|----------|-------------|
| Login não foi feito automaticamente | O script pausa e exibe uma mensagem no terminal; faça o login manualmente no navegador aberto e pressione **Enter** no terminal |
| Botão "Composições e preços" não encontrado | Clique manualmente e pressione **Enter** |
| Site lento / erros de timeout | Aumente o valor de `SLOW_MO` de `150` para `300` no topo do script |
| Nenhum dado coletado | Execute com `HEADLESS = False` para acompanhar o navegador e identificar onde está travando |
| `playwright` não reconhecido | Execute `playwright install chromium` novamente |

---

## Arquivos do projeto

```
automacao-tcpo/
├── extrair_tcpo.py      ← Script principal
├── requirements.txt     ← Dependências Python
├── .env.example         ← Modelo para as credenciais
├── .env                 ← Suas credenciais (NÃO compartilhe este arquivo)
└── tcpo_composicoes.xlsx← Resultado gerado após a execução
```
