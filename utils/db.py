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


def _corrigir_tipo_coluna_datapreco(tabela: str):
    """Corrige o tipo da coluna DataPreco para VARCHAR(20) se estiver com tipo errado."""
    conn = _conectar()
    cursor = conn.cursor()
    try:
        # Verifica o tipo atual da coluna
        cursor.execute(f"""
            SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS 
            WHERE TABLE_NAME = '{tabela}' AND COLUMN_NAME = 'DataPreco'
        """)
        resultado = cursor.fetchone()
        if resultado:
            tipo_atual = resultado[0]
            if tipo_atual.upper() not in ('VARCHAR', 'NVARCHAR', 'TEXT'):
                print(f"  ⚠️  Corrigindo tipo da coluna DataPreco em '{tabela}'...")
                print(f"     Tipo atual: {tipo_atual} → Tipo correto: VARCHAR(20)")
                cursor.execute(f"ALTER TABLE [dbo].[{tabela}] ALTER COLUMN DataPreco VARCHAR(20) NULL")
                conn.commit()
                print(f"  ✓ Coluna DataPreco corrigida com sucesso.")
    except Exception as e:
        print(f"  ⚠️  Não foi possível corrigir coluna DataPreco: {e}")
    finally:
        conn.close()


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
    _corrigir_tipo_coluna_datapreco('insumos')


def item_ja_extraido_hoje(item: str) -> bool:
    """Retorna True se já existe um registro com esse Item extraído nos últimos 7 dias (horário de Brasília)."""
    brasilia = pytz.timezone("America/Sao_Paulo")
    hoje = datetime.now(brasilia).date()

    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT 1 FROM insumos WHERE Item = ? AND CAST(DtExtracao AS DATE) >= DATEADD(day, -7, ?)",
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


def criar_tabela_servicos():
    """Cria a tabela servicos no banco se ainda não existir."""
    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute("""
        IF NOT EXISTS (
            SELECT * FROM sys.objects
            WHERE object_id = OBJECT_ID(N'[dbo].[servicos]') AND type = N'U'
        )
        BEGIN
            CREATE TABLE [dbo].[servicos] (
                CodServico   INT            NOT NULL IDENTITY(1,1) PRIMARY KEY,
                BaseExtraida VARCHAR(100)   NULL,
                Item         VARCHAR(100)   NULL,
                Descricao    VARCHAR(255)   NULL,
                Unidade      VARCHAR(20)    NULL,
                Tipo         VARCHAR(100)   NULL,
                DataPreco    VARCHAR(20)    NULL,
                Preco        DECIMAL(12,2)  NULL,
                DtExtracao   DATETIME       NULL
            )
        END
    """)
    conn.commit()
    conn.close()
    print("Tabela 'servicos' verificada/criada com sucesso.")
    _corrigir_tipo_coluna_datapreco('servicos')


def _corrigir_codcomposicao_identity():
    """Corrige a coluna CodComposicao para ser IDENTITY se ainda não estiver."""
    conn = _conectar()
    cursor = conn.cursor()
    try:
        # Verifica se a coluna CodComposicao tem IDENTITY ativado
        cursor.execute("""
            SELECT COLUMNPROPERTY(OBJECT_ID(N'[dbo].[composicoes]'), 'CodComposicao', 'IsIdentity')
        """)
        resultado = cursor.fetchone()
        if resultado and resultado[0] == 0:  # 0 = não tem identity
            print(f"  ⚠️  Corrigindo coluna CodComposicao (adicionando IDENTITY)...")
            # Recriar a tabela com IDENTITY corrigido
            cursor.execute("DROP TABLE [dbo].[composicoes]")
            cursor.execute("""
                CREATE TABLE [dbo].[composicoes] (
                    CodComposicao  INT            NOT NULL IDENTITY(1,1) PRIMARY KEY,
                    ItemServico    VARCHAR(100)   NULL,
                    ItemInsumo     VARCHAR(100)   NULL,
                    DataPreco      VARCHAR(20)    NULL,
                    Coeficiente    DECIMAL(12,6)  NULL,
                    PrecoUnitario  DECIMAL(12,2)  NULL,
                    PrecoTotal     DECIMAL(12,2)  NULL,
                    Consumo        DECIMAL(12,6)  NULL
                )
            """)
            conn.commit()
            print(f"  ✓ Coluna CodComposicao corrigida com IDENTITY.")
    except Exception as e:
        print(f"  ⚠️  Erro ao corrigir CodComposicao: {e}")
    finally:
        conn.close()


def criar_tabela_composicoes():
    """Cria a tabela composicoes no banco se ainda não existir."""
    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute("""
        IF NOT EXISTS (
            SELECT * FROM sys.objects
            WHERE object_id = OBJECT_ID(N'[dbo].[composicoes]') AND type = N'U'
        )
        BEGIN
            CREATE TABLE [dbo].[composicoes] (
                CodComposicao  INT            NOT NULL IDENTITY(1,1) PRIMARY KEY,
                ItemServico    VARCHAR(100)   NULL,
                ItemInsumo     VARCHAR(100)   NULL,
                DataPreco      VARCHAR(20)    NULL,
                Coeficiente    DECIMAL(12,6)  NULL,
                PrecoUnitario  DECIMAL(12,2)  NULL,
                PrecoTotal     DECIMAL(12,2)  NULL,
                Consumo        DECIMAL(12,6)  NULL
            )
        END
    """)
    conn.commit()
    conn.close()
    print("Tabela 'composicoes' verificada/criada com sucesso.")
    _corrigir_codcomposicao_identity()
    _corrigir_tipo_coluna_datapreco('composicoes')


def servico_ja_extraido(item: str) -> bool:
    """Verifica se o item já existe na tabela servicos."""
    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM servicos WHERE Item = ?", item)
    existe = cursor.fetchone() is not None
    conn.close()
    return existe


def insumo_existe(item: str) -> bool:
    """Verifica se o item já existe na tabela insumos (sem filtro de data)."""
    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM insumos WHERE Item = ?", item)
    existe = cursor.fetchone() is not None
    conn.close()
    return existe


def salvar_servico(base: str, item: str, descricao: str, unidade: str,
                   tipo: str, data_preco: str, preco_str: str):
    """Salva um serviço extraído no banco de dados."""
    brasilia = pytz.timezone("America/Sao_Paulo")
    dt_extracao = datetime.now(brasilia).replace(tzinfo=None)

    try:
        preco = float(preco_str.replace(".", "").replace(",", "."))
    except (ValueError, AttributeError):
        preco = None

    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO servicos (BaseExtraida, Item, Descricao, Unidade, Tipo, DataPreco, Preco, DtExtracao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        base, item, descricao, unidade, tipo, data_preco, preco, dt_extracao
    )
    conn.commit()
    conn.close()


def composicao_ja_existe(item_servico: str, item_insumo: str) -> bool:
    """Verifica se a composição de um serviço com um insumo já existe."""
    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT 1 FROM composicoes WHERE ItemServico = ? AND ItemInsumo = ?",
        item_servico, item_insumo
    )
    existe = cursor.fetchone() is not None
    conn.close()
    return existe


def salvar_composicao(item_servico: str, item_insumo: str, data_preco: str,
                      coef_str: str, preco_unit_str: str, preco_tot_str: str,
                      consumo_str: str):
    """Salva uma linha de composição no banco de dados."""
    def _parse(val):
        try:
            # Remove espaços, tabs, quebras de linha
            val = str(val).strip()
            if not val or val.upper() == 'NULL':
                return None
            return float(val.replace(".", "").replace(",", "."))
        except (ValueError, AttributeError, TypeError):
            return None

    conn = _conectar()
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO composicoes
            (ItemServico, ItemInsumo, DataPreco, Coeficiente, PrecoUnitario, PrecoTotal, Consumo)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        item_servico, item_insumo, data_preco,
        _parse(coef_str), _parse(preco_unit_str),
        _parse(preco_tot_str), _parse(consumo_str)
    )
    conn.commit()
    conn.close()
