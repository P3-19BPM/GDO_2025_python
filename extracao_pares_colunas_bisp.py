import math
import time
import pyodbc
import json
import os
from dotenv import load_dotenv
load_dotenv()


# Configurações
dsn = 'Sample Cloudera Impala DSN'
user = os.getenv('DB_USERNAME')
pwd = os.getenv('DB_PASSWORD')

# Caminhos
base_dir = r"E:\GitHub\GDO_2025_python"
input_json = os.path.join(base_dir, "relacionamentos_colunas_pares.json")
output_json = os.path.join(base_dir, "valores_relacionados.json")

# Lê o JSON com os pares
with open(input_json, "r", encoding="utf-8") as f:
    relacionamentos = json.load(f)

# Monta a lista plana com todos os pares
pares_para_executar = []
for db, tabelas in relacionamentos.items():
    for tabela, pares in tabelas.items():
        for par in pares:
            pares_para_executar.append(
                (db, tabela, par["coluna_chave"], par["coluna_descricao"]))

# Calcula totais
total_pares = len(pares_para_executar)
lote_tamanho = 10
total_lotes = math.ceil(total_pares / lote_tamanho)

print(f"🔍 Total de pares para processar: {total_pares}")
print(f"📦 Lotes de {lote_tamanho} → Total de lotes: {total_lotes}")

# Confirmação do usuário
while True:
    resposta = input(
        "Deseja iniciar a extração? (y para iniciar / q para abortar): ").strip().lower()
    if resposta == "y":
        break
    elif resposta == "q":
        print("❌ Extração abortada pelo usuário.")
        exit()
    else:
        print("⚠️ Opção inválida! Digite 'y' para iniciar ou 'q' para abortar.")

# Só conecta agora, após confirmação
conn = pyodbc.connect(f'DSN={dsn};UID={user};PWD={pwd};', autocommit=True)
cursor = conn.cursor()

resultados = {}

# Processa em lotes
for i in range(0, total_pares, lote_tamanho):
    lote = pares_para_executar[i:i + lote_tamanho]
    numero_lote = i // lote_tamanho + 1
    print(
        f"\n🚀 Executando lote {numero_lote}/{total_lotes} ({len(lote)} pares)...")

    for db, tabela, col_chave, col_desc in lote:
        try:
            query = f"""
                SELECT DISTINCT {col_chave}, {col_desc}
                FROM {db}.{tabela}
                WHERE {col_chave} IS NOT NULL AND {col_desc} IS NOT NULL
            """

            # Adiciona filtro de data se a coluna existir
            try:
                cursor.execute(f"DESCRIBE {db}.{tabela}")
                colunas_tabela = [row[0].lower() for row in cursor.fetchall()]
                if "data_hora_fato" in colunas_tabela:
                    query += " AND data_hora_fato >= '2020-01-01'"
            except:
                pass

            query += " LIMIT 500"

            cursor.execute(query)
            linhas = cursor.fetchall()

            if db not in resultados:
                resultados[db] = {}
            if tabela not in resultados[db]:
                resultados[db][tabela] = []

            resultados[db][tabela].append({
                "coluna_chave": col_chave,
                "coluna_descricao": col_desc,
                "valores": [{"codigo": row[0], "descricao": row[1]} for row in linhas]
            })

            print(
                f"✅ {db}.{tabela} - {col_chave} + {col_desc} ({len(linhas)} valores)")

        except Exception as e:
            print(f"❌ Erro em {db}.{tabela} - {col_chave}/{col_desc}: {e}")

    # Salva parcial
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(resultados, f, indent=2, ensure_ascii=False)

    print(f"💾 Lote {numero_lote} salvo em {output_json}")
    time.sleep(0.5)

cursor.close()
conn.close()

print("\n🎯 Extração concluída!")
