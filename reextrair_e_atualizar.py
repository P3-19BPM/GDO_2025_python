import time
import pyodbc
import json
import os
from dotenv import load_dotenv

# --- Configurações ---
load_dotenv()
dsn = 'Sample Cloudera Impala DSN'
user = os.getenv('DB_USERNAME')
pwd = os.getenv('DB_PASSWORD')

# --- Caminhos dos Arquivos ---
base_dir = r"E:\GitHub\GDO_2025_python"
# 1. Arquivo com a lista de tarefas (os pares que atingiram o limite)
arquivo_de_tarefas = os.path.join(
    base_dir, "pares_finais_para_reextracao.json")
# 2. Arquivo de resultados que será lido e atualizado
arquivo_de_resultados = os.path.join(base_dir, "valores_relacionados.json")

# --- Lógica Principal ---


def carregar_json(caminho_arquivo):
    """Função auxiliar para carregar um arquivo JSON com segurança."""
    try:
        with open(caminho_arquivo, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        print(
            f"❌ ERRO CRÍTICO: O arquivo '{caminho_arquivo}' não foi encontrado. Abortando.")
        return None
    except json.JSONDecodeError:
        print(
            f"❌ ERRO CRÍTICO: O arquivo '{caminho_arquivo}' não é um JSON válido. Abortando.")
        return None


def salvar_json(dados, caminho_arquivo):
    """Função auxiliar para salvar dados em um arquivo JSON."""
    with open(caminho_arquivo, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)
    print(f"\n💾 Arquivo '{caminho_arquivo}' salvo com sucesso!")


def reextrair_e_atualizar():
    """
    Função principal que orquestra a leitura, re-extração e atualização dos dados.
    """
    # Carrega a lista de pares que precisam ser re-extraídos
    pares_para_reextrair = carregar_json(arquivo_de_tarefas)
    if not pares_para_reextrair:
        return

    # Carrega o arquivo de resultados completo que será atualizado
    resultados_completos = carregar_json(arquivo_de_resultados)
    if not resultados_completos:
        return

    total_pares = len(pares_para_reextrair)
    print(
        f"🔍 Encontrados {total_pares} pares para re-extrair com limite de 5000.")

    # Confirmação do usuário
    resposta = input(
        f"Deseja iniciar a re-extração e atualizar o arquivo '{arquivo_de_resultados}'? (y/n): ").strip().lower()
    if resposta != 'y':
        print("❌ Operação abortada pelo usuário.")
        return

    # Conecta ao banco de dados
    try:
        conn = pyodbc.connect(
            f'DSN={dsn};UID={user};PWD={pwd};', autocommit=True)
        cursor = conn.cursor()
        print("\n✅ Conexão com o banco de dados estabelecida.")
    except Exception as e:
        print(
            f"❌ ERRO DE CONEXÃO: Não foi possível conectar ao banco de dados: {e}")
        return

    # Itera sobre cada tarefa da lista
    for i, tarefa in enumerate(pares_para_reextrair):
        tabela_completa = tarefa['tabela']
        col_chave = tarefa['coluna_id']
        col_desc = tarefa['coluna_descricao']

        # Divide o nome da tabela em 'db' e 'tabela'
        db, tabela = tabela_completa.split('.', 1)

        print(
            f"\n🚀 Processando {i+1}/{total_pares}: {tabela_completa} ({col_chave} / {col_desc})...")

        try:
            # Monta a query com o novo limite
            query = f"""
                SELECT DISTINCT `{col_chave}`, `{col_desc}`
                FROM `{db}`.`{tabela}`
                WHERE `{col_chave}` IS NOT NULL AND `{col_desc}` IS NOT NULL
                LIMIT 5000
            """

            cursor.execute(query)
            novas_linhas = cursor.fetchall()
            print(f"  -> Extraídos {len(novas_linhas)} novos valores.")

            # Converte os novos dados para o formato de dicionário
            novos_valores = [{"codigo": row[0], "descricao": row[1]}
                             for row in novas_linhas]

            # --- Lógica de Atualização ---
            # Encontra o local exato no JSON de resultados para substituir os dados
            encontrado_e_atualizado = False
            if db in resultados_completos and tabela in resultados_completos[db]:
                for item_existente in resultados_completos[db][tabela]:
                    if item_existente['coluna_chave'] == col_chave and item_existente['coluna_descricao'] == col_desc:
                        item_existente['valores'] = novos_valores
                        encontrado_e_atualizado = True
                        print(f"  -> ✅ Dados atualizados com sucesso no JSON.")
                        break

            if not encontrado_e_atualizado:
                print(
                    f"  -> ⚠️ AVISO: O par {tabela_completa} - {col_chave}/{col_desc} não foi encontrado no arquivo de resultados. Nenhum dado foi atualizado.")

        except Exception as e:
            print(f"  -> ❌ ERRO na query para {tabela_completa}: {e}")

        time.sleep(0.5)  # Pequena pausa entre as queries

    # Fecha a conexão com o banco
    cursor.close()
    conn.close()
    print("\n🔌 Conexão com o banco de dados fechada.")

    # Salva o arquivo JSON final com todas as atualizações
    salvar_json(resultados_completos, arquivo_de_resultados)


# --- Ponto de Entrada do Script ---
if __name__ == "__main__":
    reextrair_e_atualizar()
    print("\n🎯 Processo de re-extração e atualização concluído!")
