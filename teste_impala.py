import pyodbc
import os
from dotenv import load_dotenv

load_dotenv()

dsn_name = 'Sample Cloudera Impala DSN'
username = os.getenv('DB_USERNAME')
password = os.getenv('DB_PASSWORD')

try:
    connection_string = f'DSN={dsn_name};UID={username};PWD={password};'
    conn = pyodbc.connect(connection_string, autocommit=True)
    print("Conexão bem-sucedida!")
    conn.close()
except Exception as e:
    print(f"Erro na conexão: {e}")
