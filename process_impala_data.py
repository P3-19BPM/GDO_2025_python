from dotenv import load_dotenv
from openpyxl import load_workbook
from shapely.geometry import Point
import os
import pyodbc
import geopandas as gpd
import pandas as pd

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# --- Configuração de conexão ODBC ---
dsn_name = 'Sample Cloudera Impala DSN'
username = os.getenv('DB_USERNAME')
password = os.getenv('DB_PASSWORD')

# --- Caminhos Comuns ---
# Ajuste estes caminhos para o seu ambiente local
geojson_path = r"E:\QGis\csv\Mapas_Tratados\SubSetores_19BPM_GeoJSON.json"
sql_scripts_folder = r"E:\GitHub\GDO_2025_python\sql_scripts"

# --- Configurações para GDO_2025 ---
gdo_output_excel_path = r"E:\GDO\GDO_2025\Monitoramento_GDO_2025.xlsx"
gdo_sql_files_mapping = {
    "BD_IMV2025": "BD_IMV2025.sql",
    "BD_ICVPe": "BD_ICVPe.sql",
    "BD_ICVPa": "BD_ICVPa.sql",
    "BD_POG": "BD_POG.sql",
    "BD_PL": "BD_PL.sql",
    "BD_IRTD": "BD_IRTD.sql"
}

# --- Configurações para INT_Comunitaria_2025 ---
int_cum_output_excel_path = r"E:\GDO\GDO_2025\Monitoramento_INT_Comunitaria_2025.xlsx"
int_cum_sql_files_mapping = {
    "BD_MRPP_INT_CUM": "BD_MRPP_INT_CUM.sql",
    "BD_RC_INT_CUM": "BD_RC_INT_CUM.sql",
    "BD_VCP_INT_CUM": "BD_VCP_INT_CUM.sql",
    "BD_VTC_DENOMINADOR_INT_CUM": "BD_VTC_DENOMINADOR_INT_CUM.sql",
    "BD_VTC_NUMERADOR_INT_CUM": "BD_VTC_NUMERADOR_INT_CUM.sql",
    "BD_VT_DENOMINADOR_INT_CUM": "BD_VT_DENOMINADOR_INT_CUM.sql",
    "BD_VT_NUMERADOR_INT_CUM": "BD_VT_NUMERADOR_INT_CUM.sql"
}

# --- Caminhos para CSVs de GDO ---
gdo_csv_paths = {
    "BD_IMV2025": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_1.csv",
    "BD_ICVPe": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_2.csv",
    "BD_ICVPa": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_3.csv",
    "BD_POG": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_4.csv",
    "BD_PL": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_5.csv",
    "BD_IRTD": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\GDO_2025_6.csv",
}

# --- Caminhos para CSVs de INT_Comunitaria ---
int_cum_csv_paths = {
    "BD_MRPP_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_1.csv",
    "BD_RC_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_2.csv",
    "BD_VCP_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_3.csv",
    "BD_VTC_DENOMINADOR_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_4.csv",
    "BD_VTC_NUMERADOR_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_5.csv",
    "BD_VT_DENOMINADOR_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_6.csv",
    "BD_VT_NUMERADOR_INT_CUM": r"E:\Painel_PowerBI\BASE_GDO\BD_2025\INT_COM_2025_7.csv"
}


# --- Função para ler o conteúdo de um arquivo SQL ---
def read_sql_file(filepath):
    if not os.path.exists(filepath):
        print(f"❌ Erro: Arquivo SQL não encontrado em {filepath}")
        return None
    with open(filepath, 'r', encoding='utf-8') as file:
        return file.read()

# --- Função para processar cada consulta SQL e retornar um DataFrame ---


def fetch_data_from_impala(sql_query, dsn, user, pwd):
    conn = None
    cursor = None
    try:
        connection_string = f'DSN={dsn};UID={user};PWD={pwd};'
        conn = pyodbc.connect(connection_string, autocommit=True)
        cursor = conn.cursor()
        print(f"⏳ Executando consulta SQL...")
        cursor.execute(sql_query)

        if cursor.description is None:
            print("⚠️ ATENÇÃO: cursor.description é None. A consulta pode não ter retornado resultados ou falhou no servidor.")
            print(f"Consulta SQL executada: \n{sql_query[:200]}...")
            try:
                rows = cursor.fetchall()
                if not rows:
                    print("⚠️ Nenhuma linha retornada pela consulta.")
                else:
                    print(
                        f"⚠️ Linhas retornadas, mas sem descrição de coluna: {len(rows)} linhas.")
            except Exception as fetch_e:
                print(
                    f"❌ Erro ao tentar buscar linhas após cursor.description ser None: {fetch_e}")
            raise ValueError(
                "Não foi possível obter a descrição das colunas da consulta SQL.")

        columns = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
        df = pd.DataFrame.from_records(rows, columns=columns)
        print(f"✅ Consulta finalizada: {df.shape[0]} registros obtidos.")
        return df
    except Exception as e:
        print(f"❌ Erro ao conectar ou executar consulta no Impala: {e}")
        raise
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

# --- Função para Processar Cada DataFrame e realizar a junção espacial ---


def process_dataframe_for_spatial_join(df, geojson_path):
    try:
        # 1. Tratamento de coordenadas nulas e conversão para numérico
        # Cria uma cópia para evitar SettingWithCopyWarning
        df_processed = df.copy()

        # Converte as colunas de latitude e longitude para tipo numérico
        # errors='coerce' irá transformar valores não numéricos em NaN
        df_processed['numero_latitude'] = pd.to_numeric(
            df_processed['numero_latitude'], errors='coerce')
        df_processed['numero_longitude'] = pd.to_numeric(
            df_processed['numero_longitude'], errors='coerce')

        # Remove linhas onde as coordenadas são NaN após a conversão
        initial_rows = len(df_processed)
        df_processed.dropna(
            subset=['numero_latitude', 'numero_longitude'], inplace=True)
        if len(df_processed) < initial_rows:
            print(
                f"⚠️ {initial_rows - len(df_processed)} linhas removidas devido a coordenadas nulas ou inválidas.")

        # 2. Tratamento de REDS duplicados (mantém a primeira ocorrência)
        # Assumindo que 'numero_ocorrencia' é a coluna que identifica o REDS
        if 'numero_ocorrencia' in df_processed.columns:
            initial_rows_after_na = len(df_processed)
            df_processed.drop_duplicates(
                subset=['numero_ocorrencia'], inplace=True)
            if len(df_processed) < initial_rows_after_na:
                print(
                    f"⚠️ {initial_rows_after_na - len(df_processed)} REDS duplicados removidos (mantida a primeira ocorrência).")
        else:
            print(
                "⚠️ Coluna 'numero_ocorrencia' não encontrada para verificar duplicados.")

        # Verificar se as colunas necessárias existem após os tratamentos
        required_columns = ['numero_longitude', 'numero_latitude']
        missing_columns = [
            col for col in required_columns if col not in df_processed.columns]

        if missing_columns:
            print(
                f"\n[ERRO] DataFrame não possui as colunas esperadas: {required_columns} após tratamento.")
            print(f"Colunas faltando: {missing_columns}")
            print(f"Colunas disponíveis: {df_processed.columns.tolist()}\n")
            raise ValueError(f"Colunas faltando: {missing_columns}")

        # Criar GeoDataFrame
        # Garante que Point receba valores escalares float
        df_processed['geometry'] = df_processed.apply(lambda row: Point(
            float(row['numero_longitude']), float(row['numero_latitude'])), axis=1)
        points_gdf = gpd.GeoDataFrame(
            df_processed, geometry='geometry', crs='EPSG:4326')

        # Carregar e alinhar o GeoJSON
        polygons_gdf = gpd.read_file(geojson_path)
        if polygons_gdf.crs != 'EPSG:4326':
            polygons_gdf = polygons_gdf.to_crs('EPSG:4326')

        # Junção espacial
        print("📍 Realizando junção espacial...")
        result_gdf = gpd.sjoin(points_gdf, polygons_gdf,
                               how='left', predicate='within')

        # Limpar colunas desnecessárias (como visuais do GeoJSON)
        colunas_a_remover = ['index_right',
                             'fill-opacity', 'stroke-opacity', 'stroke']
        result_gdf.drop(columns=[
            col for col in colunas_a_remover if col in result_gdf.columns], inplace=True)

        # Assegurar que as colunas 'name', 'PELOTAO', 'CIA_PM' existam antes de tentar acessá-las
        # Se elas não existirem no resultado da junção, elas serão adicionadas como NaN
        for col in ['name', 'PELOTAO', 'CIA_PM']:
            if col not in result_gdf.columns:
                result_gdf[col] = None

        # Selecionar as colunas originais do DataFrame processado mais as novas colunas da junção
        # Garantir que as colunas de saída existam no result_gdf antes de selecioná-las
        output_columns = [col for col in df_processed.columns if col !=
                          'geometry'] + ['name', 'PELOTAO', 'CIA_PM']
        output_columns = [
            col for col in output_columns if col in result_gdf.columns]
        result = result_gdf[output_columns]

        result = result.rename(columns={
            'name': 'SubSetor',
            'PELOTAO': 'Pelotao',
            'CIA_PM': 'CIA_PM'
        })

        if 'data_hora_fato' in result.columns:
            result['data_fato'] = pd.to_datetime(
                result['data_hora_fato'], errors='coerce').dt.date

        return result

    except Exception as e:
        print(
            f"\n[ERRO] Falha ao processar o DataFrame para junção espacial: {str(e)}\n")
        raise

# --- Função para processar um conjunto de dados e salvar no Excel ---


def process_data_set(output_excel_path, sql_mapping, geojson_path, dsn, user, pwd, dataset_name="", csv_paths=None):
    print(f"\n{'='*50}")
    print(f"Iniciando processamento para: {dataset_name}")
    print(f"{'='*50}\n")

    # Cria o diretório para os CSVs se ele não existir
    if csv_paths:
        # Pega o diretório do primeiro CSV para criar a pasta base
        # Garante que o dicionário não está vazio antes de tentar acessar
        if csv_paths:
            first_csv_path = list(csv_paths.values())[0]
            csv_dir = os.path.dirname(first_csv_path)
            os.makedirs(csv_dir, exist_ok=True)
            print(f"DEBUG: Diretório para CSVs criado/verificado: {csv_dir}")

    try:
        # Use pd.ExcelWriter com mode='a' para anexar/substituir abas
        # Se o arquivo não existir, ele será criado automaticamente
        with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            for sheet_name, sql_filename in sql_mapping.items():
                sql_filepath = os.path.join(sql_scripts_folder, sql_filename)
                print(
                    f"\nProcessando aba: {sheet_name} a partir de {sql_filepath}")

                sql_query = read_sql_file(sql_filepath)
                if sql_query is None:
                    print(
                        f"⚠️ Pulando {sheet_name} devido a arquivo SQL não encontrado ou vazio.")
                    continue

                try:
                    raw_data_df = fetch_data_from_impala(
                        sql_query, dsn, user, pwd)

                    # --- SALVAR RAW DATA EM CSV (ANTES DO PROCESSAMENTO GEOESPACIAL) ---
                    if csv_paths and sheet_name in csv_paths:
                        csv_output_path = csv_paths[sheet_name]
                        raw_data_df.to_csv(
                            csv_output_path, index=False, encoding='utf-8')
                        print(
                            f"✅ Dados brutos salvos em CSV: {csv_output_path}")
                    # --- FIM SALVAR RAW DATA EM CSV ---

                    processed_data = process_dataframe_for_spatial_join(
                        raw_data_df, geojson_path)
                    processed_data.to_excel(
                        writer, sheet_name=sheet_name, index=False)
                except Exception as e:
                    print(f"❌ Erro ao processar {sheet_name}: {e}")
                    # Continua para a próxima aba mesmo se uma falhar
                    continue
        print(
            f"\n✅ Arquivo {dataset_name} atualizado com sucesso em: {output_excel_path}")
    except FileNotFoundError:
        # Se o arquivo não existir, crie um novo com mode='w'
        with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='w') as writer:
            for sheet_name, sql_filename in sql_mapping.items():
                sql_filepath = os.path.join(sql_scripts_folder, sql_filename)
                print(
                    f"\nProcessando aba: {sheet_name} a partir de {sql_filepath}")

                sql_query = read_sql_file(sql_filepath)
                if sql_query is None:
                    print(
                        f"⚠️ Pulando {sheet_name} devido a arquivo SQL não encontrado ou vazio.")
                    continue

                try:
                    raw_data_df = fetch_data_from_impala(
                        sql_query, dsn, user, pwd)

                    # --- SALVAR RAW DATA EM CSV (ANTES DO PROCESSAMENTO GEOESPACIAL) ---
                    if csv_paths and sheet_name in csv_paths:
                        csv_output_path = csv_paths[sheet_name]
                        raw_data_df.to_csv(
                            csv_output_path, index=False, encoding='utf-8')
                        print(
                            f"✅ Dados brutos salvos em CSV: {csv_output_path}")
                    # --- FIM SALVAR RAW DATA EM CSV ---

                    processed_data = process_dataframe_for_spatial_join(
                        raw_data_df, geojson_path)
                    processed_data.to_excel(
                        writer, sheet_name=sheet_name, index=False)
                except Exception as e:
                    print(f"❌ Erro ao processar {sheet_name}: {e}")
                    continue
        print(
            f"\n✅ Novo arquivo {dataset_name} criado em: {output_excel_path}")
    except Exception as e:
        print(
            f"\n❌ [ERRO FATAL] Ocorreu um erro durante o processamento de {dataset_name}: {e}")


# --- Execução Principal ---
if __name__ == "__main__":
    # Processar dados GDO
    process_data_set(
        output_excel_path=gdo_output_excel_path,
        sql_mapping=gdo_sql_files_mapping,
        geojson_path=geojson_path,
        dsn=dsn_name,
        user=username,
        pwd=password,
        dataset_name="GDO_2025",
        csv_paths=gdo_csv_paths  # Adicionado
    )

    # Processar dados de Interação Comunitária
    process_data_set(
        output_excel_path=int_cum_output_excel_path,
        sql_mapping=int_cum_sql_files_mapping,
        geojson_path=geojson_path,
        dsn=dsn_name,
        user=username,
        pwd=password,
        dataset_name="INT_Comunitaria_2025",
        csv_paths=int_cum_csv_paths  # Adicionado
    )

    print("\nProcessamento de todos os conjuntos de dados concluído.")
