import pandas as pd
import json

# Caminho do CSV de entrada e JSON de saída
csv_path = r"E:\GitHub\GDO_2025_python\colunas_tabelas_impala.csv"
json_path = r"E:\GitHub\GDO_2025_python\colunas_tabelas_impala.json"

# Carrega o CSV
df = pd.read_csv(csv_path)

# Cria estrutura aninhada
estrutura = {}

for _, row in df.iterrows():
    db = row["Database"]
    tabela = row["Tabela"]
    coluna = row["Coluna"]
    tipo = row["Tipo de Dados"]

    if db not in estrutura:
        estrutura[db] = {}
    if tabela not in estrutura[db]:
        estrutura[db][tabela] = []

    estrutura[db][tabela].append({
        "coluna": coluna,
        "tipo": tipo
    })

# Salva como JSON
with open(json_path, "w", encoding="utf-8") as f:
    json.dump(estrutura, f, indent=2, ensure_ascii=False)

print(f"✅ JSON estruturado salvo em: {json_path}")
