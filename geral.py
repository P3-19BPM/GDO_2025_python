import pandas as pd
import geopandas as gpd
import pyodbc
import os
from shapely.geometry import Point
from openpyxl import load_workbook

# --- Configuração de conexão ODBC (Mantenha suas credenciais aqui ou use variáveis de ambiente) ---
# Certifique-se de que este DSN está configurado corretamente no seu sistema
dsn_name = 'Sample Cloudera Impala DSN'
username = 'Vai ser meu CPF, pretendo colocar em um .env'
password = 'Aqui deve ir minha Senha, deve ir em um .env'

# --- Caminhos Comuns ---
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
    "BD_IRTD": "BD_IRTD.sql",
    # Este SQL pode ser removido se for exclusivo do INT_CUM
    "BD_INTERACAO_COMUNITARIA": "BD_INTERACAO_COMUNITARIA.sql"
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

# --- Função para ler o conteúdo de um arquivo SQL ---


def read_sql_file(filepath):
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

        # --- Adicionar depuração aqui ---
        if cursor.description is None:
            print("⚠️ ATENÇÃO: cursor.description é None. A consulta pode não ter retornado resultados ou falhou no servidor.")
            # Imprime os primeiros 200 caracteres da query
            print(f"Consulta SQL executada: \n{sql_query[:200]}...")
            # Tente buscar linhas mesmo que description seja None, para ver se há algum erro ou resultado inesperado
            try:
                # Isso pode levantar uma exceção se não houver resultados válidos
                rows = cursor.fetchall()
                if not rows:
                    print("⚠️ Nenhuma linha retornada pela consulta.")
                else:
                    print(
                        f"⚠️ Linhas retornadas, mas sem descrição de coluna: {len(rows)} linhas.")
                    # Se houver linhas, mas sem descrição, isso é um cenário incomum e pode indicar um problema grave.
                    # Você pode querer inspecionar as primeiras linhas: print(rows[:5])
            except Exception as fetch_e:
                print(
                    f"❌ Erro ao tentar buscar linhas após cursor.description ser None: {fetch_e}")
            # Levantar um erro para interromper o processamento, pois não podemos criar o DataFrame sem colunas
            raise ValueError(
                "Não foi possível obter a descrição das colunas da consulta SQL.")

        columns = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
        df = pd.DataFrame.from_records(rows, columns=columns)
        print(f"✅ Consulta finalizada: {df.shape[0]} registros obtidos.")
        return df
    except Exception as e:
        print(f"❌ Erro ao conectar ou executar consulta no Impala: {e}")
        raise  # Re-lança a exceção para que o erro fatal seja capturado no nível superior
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


# --- Função para Processar Cada DataFrame e realizar a junção espacial ---


def process_dataframe_for_spatial_join(df, geojson_path):
    try:
        # Remover registros com coordenadas ausentes
        df.dropna(subset=['numero_latitude', 'numero_longitude'], inplace=True)

        # Verificar se as colunas necessárias existem
        required_columns = ['numero_longitude', 'numero_latitude']
        missing_columns = [
            col for col in required_columns if col not in df.columns]

        if missing_columns:
            print(
                f"\n[ERRO] DataFrame não possui as colunas esperadas: {required_columns}")
            print(f"Colunas faltando: {missing_columns}")
            print(f"Colunas disponíveis: {df.columns.tolist()}\n")
            raise ValueError(f"Colunas faltando: {missing_columns}")

        # Criar GeoDataFrame
        df['geometry'] = df.apply(lambda row: Point(
            row['numero_longitude'], row['numero_latitude']), axis=1)
        points_gdf = gpd.GeoDataFrame(df, geometry='geometry', crs='EPSG:4326')

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
                # Ou pd.NA, dependendo da sua preferência
                result_gdf[col] = None

        output_columns = list(df.columns) + ['name', 'PELOTAO', 'CIA_PM']
        # Garantir que as colunas de saída existam no result_gdf antes de selecioná-las
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


def process_data_set(output_excel_path, sql_mapping, geojson_path, dsn, user, pwd, dataset_name=""):
    print(f"\n{'='*50}")
    print(f"Iniciando processamento para: {dataset_name}")
    print(f"{'='*50}\n")
    try:
        # Se o arquivo já existir, carregar as abas existentes
        with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            for sheet_name, sql_filename in sql_mapping.items():
                sql_filepath = os.path.join(sql_scripts_folder, sql_filename)
                print(
                    f"\nProcessando aba: {sheet_name} a partir de {sql_filepath}")

                # 1. Ler o SQL do arquivo
                sql_query = read_sql_file(sql_filepath)

                # 2. Executar a consulta no Impala e obter o DataFrame
                raw_data_df = fetch_data_from_impala(sql_query, dsn, user, pwd)

                # 3. Processar o DataFrame para junção espacial
                processed_data = process_dataframe_for_spatial_join(
                    raw_data_df, geojson_path)

                # Salvar ou substituir a aba no Excel
                processed_data.to_excel(
                    writer, sheet_name=sheet_name, index=False)
        print(
            f"\n✅ Arquivo {dataset_name} atualizado com sucesso em: {output_excel_path}")
    except FileNotFoundError:
        # Se o arquivo não existir, criar um novo
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            for sheet_name, sql_filename in sql_mapping.items():
                sql_filepath = os.path.join(sql_scripts_folder, sql_filename)
                print(
                    f"\nProcessando aba: {sheet_name} a partir de {sql_filepath}")

                # 1. Ler o SQL do arquivo
                sql_query = read_sql_file(sql_filepath)

                # 2. Executar a consulta no Impala e obter o DataFrame
                raw_data_df = fetch_data_from_impala(sql_query, dsn, user, pwd)

                # 3. Processar o DataFrame para junção espacial
                processed_data = process_dataframe_for_spatial_join(
                    raw_data_df, geojson_path)

                # Salvar no Excel
                processed_data.to_excel(
                    writer, sheet_name=sheet_name, index=False)
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
        dataset_name="GDO_2025"
    )

    # Processar dados de Interação Comunitária
    process_data_set(
        output_excel_path=int_cum_output_excel_path,
        sql_mapping=int_cum_sql_files_mapping,
        geojson_path=geojson_path,
        dsn=dsn_name,
        user=username,
        pwd=password,
        dataset_name="INT_Comunitaria_2025"
    )

    print("\nProcessamento de todos os conjuntos de dados concluído.")
