import os
import pyodbc
import pandas as pd
import pytz
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()


def _conectar():
    conn_str = os.getenv("DB_CONNECTION_STRING")
    # Aceita tanto o formato pyodbc (DRIVER=...) quanto o formato .NET (Server=...)
    if conn_str and "DRIVER=" not in conn_str.upper():
        drivers_sql = [d for d in pyodbc.drivers() if "SQL Server" in d]
        driver = next(
            (d for d in ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "SQL Server"]
             if d in drivers_sql),
            drivers_sql[0] if drivers_sql else "SQL Server"
        )
        conn_str = f"DRIVER={{{driver}}};{conn_str}"
    # Normaliza atributos booleanos: True/False → yes/no (pyodbc não aceita True/False)
    import re as _re
    for attr in ['Trusted_Connection', 'TrustServerCertificate', 'Encrypt']:
        conn_str = _re.sub(rf'{attr}\s*=\s*True', f'{attr}=yes', conn_str, flags=_re.IGNORECASE)
        conn_str = _re.sub(rf'{attr}\s*=\s*False', f'{attr}=no', conn_str, flags=_re.IGNORECASE)
    return pyodbc.connect(conn_str)


def criar_tabela_insumos():
    """Cria a tabela insumos no banco se ainda não existir."""
    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute("""
        IF NOT EXISTS (
            SELECT * FROM sys.objects
            WHERE object_id = OBJECT_ID(N'[dbo].[insumos]') AND type = N'U'
        )
        BEGIN
            CREATE TABLE [dbo].[insumos] (
                CodInsumo    INT            NOT NULL IDENTITY(1,1) PRIMARY KEY,
                BaseExtraida VARCHAR(100)   NULL,
                Item         VARCHAR(100)   NULL,
                Descricao    VARCHAR(255)   NULL,
                Unidade      VARCHAR(20)    NULL,
                Tipo         VARCHAR(100)   NULL,
                DataPreco    VARCHAR(20)    NULL,
                Preco        DECIMAL(12, 2) NULL,
                DtExtracao   DATETIME       NULL
            )
        END
    """)
    conn.commit()
    conn.close()
    print("Tabela 'insumos' verificada/criada com sucesso.")


def item_ja_extraido_hoje(item: str) -> bool:
    """Retorna True se já existe um registro com esse Item extraído hoje (horário de Brasília)."""
    brasilia = pytz.timezone("America/Sao_Paulo")
    hoje = datetime.now(brasilia).date()

    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT 1 FROM insumos WHERE Item = ? AND CAST(DtExtracao AS DATE) = ?",
        item,
        str(hoje)
    )
    existe = cursor.fetchone() is not None
    conn.close()
    return existe


def salvar_insumo(base: str, item: str, descricao: str, unidade: str,
                  tipo: str, data_preco: str, preco_str: str):
    """Salva um insumo extraído no banco de dados."""
    brasilia = pytz.timezone("America/Sao_Paulo")
    dt_extracao = datetime.now(brasilia).replace(tzinfo=None)  # pyodbc não suporta tz-aware

    # Converte formato de preço brasileiro (1.234,56) para decimal
    try:
        preco = float(preco_str.replace(".", "").replace(",", "."))
    except (ValueError, AttributeError):
        preco = None

    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO insumos (BaseExtraida, Item, Descricao, Unidade, Tipo, DataPreco, Preco, DtExtracao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        base, item, descricao, unidade, tipo, data_preco, preco, dt_extracao
    )
    conn.commit()
    conn.close()


def exportar_banco_para_excel():
    """Exporta todos os registros do banco para um arquivo Excel em /arquivos."""
    pasta = "arquivos"
    if not os.path.exists(pasta):
        os.makedirs(pasta)

    brasilia = pytz.timezone("America/Sao_Paulo")
    agora = datetime.now(brasilia).strftime("%Y-%m-%d_%H-%M-%S")
    nome_arquivo = f"insumos_db_{agora}.xlsx"
    caminho = os.path.join(pasta, nome_arquivo)

    conn = _conectar()
    df = pd.read_sql(
        """
        SELECT
            CodInsumo,
            BaseExtraida  AS [Base],
            Item,
            Descricao     AS [Descrição],
            Unidade       AS [Un.],
            Tipo,
            DataPreco     AS [Data Preço],
            Preco         AS [Preço],
            DtExtracao    AS [Dt. Extração]
        FROM insumos
        ORDER BY DtExtracao DESC
        """,
        conn
    )
    conn.close()

    df.to_excel(caminho, index=False)
    print(f"Dados exportados para: {caminho}")
    return caminho
