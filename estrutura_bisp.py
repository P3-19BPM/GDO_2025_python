import pyodbc
import pandas as pd
import os

# --- Configurações de conexão ---
dsn = 'Sample Cloudera Impala DSN'  # Altere se necessário
user = os.getenv('DB_USERNAME')
pwd = os.getenv('DB_PASSWORD')

# --- Caminho de saída ---
output_csv_path = r"E:\GDO\colunas_tabelas_impala.csv"

# --- Conexão ---
conn = pyodbc.connect(f'DSN={dsn};UID={user};PWD={pwd};', autocommit=True)
cursor = conn.cursor()

# --- Obter bancos ---
cursor.execute("SHOW DATABASES")
databases = [row[0] for row in cursor.fetchall() if row[0].startswith("db_")]

# --- Coleta ---
dados_colunas = []

for db in databases:
    print(f"🔍 Banco: {db}")
    try:
        cursor.execute(f"SHOW TABLES IN {db}")
        tabelas = [row[0] for row in cursor.fetchall()]
        for tabela in tabelas:
            try:
                cursor.execute(f"DESCRIBE {db}.{tabela}")
                for coluna in cursor.fetchall():
                    dados_colunas.append([
                        db, tabela, coluna[0], coluna[1]
                    ])
            except Exception as e:
                print(f"❌ Erro ao descrever {db}.{tabela}: {e}")
    except Exception as e:
        print(f"❌ Erro ao acessar tabelas de {db}: {e}")

# --- Exportar para CSV ---
df = pd.DataFrame(dados_colunas, columns=[
    "Database", "Tabela", "Coluna", "Tipo de Dados"
])
df.to_csv(output_csv_path, index=False, encoding='utf-8')
print(f"\n✅ Arquivo CSV gerado em: {output_csv_path}")
