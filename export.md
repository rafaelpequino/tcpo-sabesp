# `export.py`

## Objetivo

`export.py` automatiza a coleta de dados do portal TCPO PINI usando Selenium e grava os resultados em um SQL Server por meio de `utils.db`. O fluxo abre o Chrome, faz login, seleciona o banco **TCPO PINI** e percorre as categorias disponíveis no portal.

## Fluxo principal

A execução começa em `main()`:

1. Carrega as variáveis do arquivo `.env`.
2. Inicializa o ChromeDriver com opções para reduzir consumo de memória e desabilitar imagens/notificações.
3. Acessa `https://tcpoweb.pini.com.br/home/home.aspx` e chama `login()`.
4. Solicita no terminal o tipo de exportação: `1` para serviços e composições ou `2` para insumos.
5. Chama `acessar_banco()` para abrir “Bases, composições e ferramentas”, ativar a busca avançada e selecionar `TCPO_PINI|1|`.
6. Executa somente a rotina escolhida (`exportar_servicos()` ou `exportar_insumos()`).
7. Ao terminar, aguarda uma confirmação no terminal e encerra a sessão/navegador com `encerrar()`.

## Exportação de serviços e composições

`exportar_servicos()` mantém um dicionário extenso de categorias e IDs dos nós da árvore do portal. Para cada categoria:

- expande a árvore de navegação;
- clica no nó correspondente;
- identifica o total de páginas pelo texto `Página X de Y`;
- percorre as linhas da tabela `gvServicos`, ignorando cabeçalho e linha de paginação;
- abre cada serviço e coleta base, código/item, descrição, unidade, data de preço e preço total;
- grava o serviço em `servicos`, desde que o item ainda não exista;
- lê a tabela `gvComposicao` e grava cada composição em `composicoes`;
- insere na tabela `insumos` os insumos encontrados na composição que ainda não existirem, marcando-os como `EXTRA`;
- retorna à lista e avança para a próxima página.

A página seguinte é localizada pelo link numérico do paginador ou, quando necessário, pelo botão `...`. O script confirma a página lendo novamente o texto `Página X de Y`. Depois de até seis tentativas sem sucesso, solicita que o usuário navegue manualmente e pressione ENTER.

## Exportação de insumos

`exportar_insumos()` usa seis categorias específicas: materiais, mão de obra, mão de obra empreitada, serviços terceirizados e equipamentos de aquisição/locação. Para cada item, abre a página de detalhes, coleta os dados e grava na tabela `insumos`.

Antes de salvar, verifica `db.item_ja_extraido_hoje()`. Apesar do nome, essa verificação considera registros dos últimos **7 dias**, evitando nova coleta recente do mesmo item.

## Banco de dados

As funções de persistência estão em `utils.db` e usam `DB_CONNECTION_STRING` do `.env`. O módulo cria, se necessário, as tabelas `servicos`, `insumos` e `composicoes`, converte preços brasileiros como `1.234,56` para decimal e registra a data de extração no fuso de São Paulo.

Principais relações gravadas:

- `servicos`: dados gerais, preço do serviço, status de conclusão (`Finalizado`) e memorial descritivo (`Memorial_Conteudo`, `Memorial_Criterio`, `Memorial_Normas`, `Memorial_Observacoes`);
- `insumos`: materiais/recursos associados e seus preços;
- `composicoes`: vínculo entre serviço e insumo, coeficiente, preço unitário, preço total e consumo.

## Tratamento de falhas

O script usa `WebDriverWait`, retries para elementos obsoletos (`StaleElementReferenceException`) e tratamento por categoria/página para continuar a coleta quando uma linha falha. Após três timeouts consecutivos, considera possível perda de conexão, pausa e pede ao usuário para reconectar. O usuário pode digitar `Q` para cancelar.

Ao final, informa se o processo foi finalizado, interrompido ou terminou com erro. Falhas críticas do WebDriver são capturadas separadamente.

## Pré-requisitos

- Google Chrome e ChromeDriver compatível;
- Python com as dependências de `requirements.txt`;
- arquivo `.env` contendo `USER_INPUT`, `PASS_INPUT` e `DB_CONNECTION_STRING`;
- SQL Server acessível e um driver ODBC instalado;
- acesso ao portal TCPO PINI.

O script depende de IDs HTML específicos do portal. Mudanças na estrutura ou nos IDs do site podem interromper a navegação e exigirão atualização dos dicionários e seletores em `export.py`.
